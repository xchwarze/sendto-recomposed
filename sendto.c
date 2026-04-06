/*
 * sendto.c – Unicode-only reworked version of SendTo+ (MSVC)
 * Copyright (c) 2025 DSR! <xchwarze@gmail.com>
 * based on https://github.com/lifenjoiner/sendto-plus
 */

#include <sdkddkver.h>
#define WIN32_LEAN_AND_MEAN
#ifndef _WIN32_WINNT
#define _WIN32_WINNT _WIN32_WINNT_WIN7
#define _WIN32_IE _WIN32_IE_IE80
#endif

#include <windows.h>
#include <stdlib.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <commoncontrols.h> /* for IID_IImageList */
#include <shellapi.h>
#include <strsafe.h>
#include <stdbool.h>

#pragma comment(lib, "comctl32.lib")   // commctrl.h – InitCommonControlsEx, ImageList_*, etc.
#pragma comment(lib, "shell32.lib")    // shlobj.h, shobjidl.h – SHGetKnownFolderPath, IShellItem, etc.
#pragma comment(lib, "shlwapi.lib")    // shlwapi.h – PathIsDirectoryW, StrCmpLogicalW, etc.
#pragma comment(lib, "ole32.lib")      // COM: CoCreateInstance, etc.

#define MAX_DEPTH 5
#define MAX_LOCAL_PATH 32767
#define MENU_POOL_SIZE 64

/** Binary cache file signature: "STC\0" (SendTo Cache). */
#define CACHE_MAGIC  0x00435453
#define CACHE_VERSION 1

static LPSHELLFOLDER desktopShellFolder = NULL;
static HDC hdcIconCache = NULL;

/* -------------------------------------------------------------------------- */
/* Utility macros                                                             */
/* -------------------------------------------------------------------------- */

#define SAFE_RELEASE(p)     do { if (p) { (p)->lpVtbl->Release(p); (p) = NULL; } } while (0)
#define ERR_BOX(msg)        MessageBoxW(NULL, msg, L"SendTo+", MB_OK|MB_ICONERROR)


/* -------------------------------------------------------------------------- */
/* Helpers                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * OptInDarkPopupMenus
 *
 * Enable process-wide dark theming for popup/context menus on
 * Windows 10 1809 + by calling uxtheme ordinals 135/136.
 * Safe on older builds: missing exports are ignored.
 */
static void OptInDarkPopupMenus(void)
{
    typedef enum _PreferredAppMode {
        AppMode_Default     = 0,
        AppMode_AllowDark   = 1,
        AppMode_ForceDark   = 2,
        AppMode_ForceLight  = 3,
        AppMode_Max
    } PreferredAppMode;

    HMODULE uxThemeModule = LoadLibraryW(L"uxtheme.dll");
    if (!uxThemeModule) {
        return;
    }

    // Ordinals are stable since 1809
    typedef PreferredAppMode (WINAPI *SetPreferredAppMode_t)(PreferredAppMode);
    typedef BOOL (WINAPI *FlushMenuThemes_t)(void);

    SetPreferredAppMode_t pSetPreferredAppMode =
        (SetPreferredAppMode_t)GetProcAddress(uxThemeModule, MAKEINTRESOURCEA(135));
    FlushMenuThemes_t pFlushMenuThemes =
        (FlushMenuThemes_t)GetProcAddress(uxThemeModule, MAKEINTRESOURCEA(136));

    if (pSetPreferredAppMode) {
        pSetPreferredAppMode(AppMode_AllowDark);
    }

    if (pFlushMenuThemes) {
        pFlushMenuThemes();
    }

    // Release the library – the theme mode is already applied process-wide
    FreeLibrary(uxThemeModule);
}

/* ---- Undocumented types for SetWindowCompositionAttribute ---------------- */
typedef enum {
    ACCENT_DISABLED                    = 0,
    ACCENT_ENABLE_GRADIENT             = 1,
    ACCENT_ENABLE_TRANSPARENTGRADIENT  = 2,
    ACCENT_ENABLE_BLURBEHIND           = 3,
    ACCENT_ENABLE_ACRYLICBLURBEHIND    = 4,   /* acrylic – W10 1809+ / W11  */
    ACCENT_ENABLE_HOSTBACKDROP         = 5,    /* Mica – W11 only            */
} ACCENT_STATE;
 
typedef struct {
    ACCENT_STATE AccentState;
    DWORD        AccentFlags;
    DWORD        GradientColor;   /* AABBGGRR format */
    DWORD        AnimationId;
} ACCENT_POLICY;
 
typedef struct {
    DWORD  Attribute;    /* 19 = WCA_ACCENT_POLICY */
    PVOID  Data;
    ULONG  DataSize;
} WINDOWCOMPOSITIONATTRIBDATA;
 
#define WCA_ACCENT_POLICY 19
 
typedef BOOL (WINAPI *SetWindowCompositionAttribute_t)(HWND, WINDOWCOMPOSITIONATTRIBDATA *);
 
/** Resolved once; NULL if unavailable. */
static SetWindowCompositionAttribute_t pSetWindowCompositionAttribute = NULL;
 
/** TRUE after first resolution attempt (succeed or fail). */
static bool g_acrylicResolved = false;
 
/**
 * EnsureAcrylicProc – lazy-resolve SetWindowCompositionAttribute from user32.
 *
 * Safe to call many times; only loads once.
 */
static void EnsureAcrylicProc(void)
{
    if (g_acrylicResolved) return;
    g_acrylicResolved = true;
 
    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (hUser32) {
        pSetWindowCompositionAttribute =
            (SetWindowCompositionAttribute_t)GetProcAddress(
                hUser32, "SetWindowCompositionAttribute");
    }
}
 
/**
 * IsSystemDarkMode – check if Windows is set to dark app mode.
 *
 * Reads the registry value directly; works on W10 1809+ and W11.
 * Returns FALSE on older builds or if the key doesn't exist.
 */
static BOOL IsSystemDarkMode(void)
{
    DWORD value = 0;
    DWORD size  = sizeof(value);
    LONG lr = RegGetValueW(
        HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
        L"AppsUseLightTheme",
        RRF_RT_REG_DWORD,
        NULL,
        &value,
        &size
    );
 
    /* AppsUseLightTheme: 0 = dark, 1 = light */
    return (lr == ERROR_SUCCESS && value == 0);
}
 
/**
 * ApplyAcrylicToWindow – set acrylic backdrop on an HWND.
 *
 * @param hwnd   Target window (typically a #32768 menu window).
 */
static void ApplyAcrylicToWindow(HWND hwnd)
{
    if (!pSetWindowCompositionAttribute) return;
 
    /*
     * Tint colour in AABBGGRR:
     *   Dark mode:   semi-transparent dark    →  0xB0000000
     *   Light mode:  semi-transparent white   →  0xB0FFFFFF
     *
     * Adjust alpha (first byte) to taste:
     *   0x01 = almost fully transparent (maximum blur)
     *   0xCC = mostly opaque (subtle blur)
     *   0xB0 = balanced (matches system menu feel)
     */
    DWORD tintColor = IsSystemDarkMode() ? 0xB0000000 : 0xB0FFFFFF;
 
    ACCENT_POLICY accent = {
        .AccentState   = ACCENT_ENABLE_ACRYLICBLURBEHIND,
        .AccentFlags   = 2,   /* ACCENT_FLAG_DRAW_ALL_BORDERS (undocumented) */
        .GradientColor = tintColor,
        .AnimationId   = 0,
    };
 
    WINDOWCOMPOSITIONATTRIBDATA data = {
        .Attribute = WCA_ACCENT_POLICY,
        .Data      = &accent,
        .DataSize  = sizeof(accent),
    };
 
    pSetWindowCompositionAttribute(hwnd, &data);
}
 
/**
 * AcrylicMenuEventProc – WinEvent callback that applies acrylic to popup
 *                         menu windows (#32768) as they appear.
 *
 * Installed just before TrackPopupMenuEx, uninstalled right after it returns.
 */
static void CALLBACK AcrylicMenuEventProc(
    HWINEVENTHOOK hHook,
    DWORD         event,
    HWND          hwnd,
    LONG          idObject,
    LONG          idChild,
    DWORD         dwEventThread,
    DWORD         dwmsEventTime
) {
    (void)hHook;
    (void)event;
    (void)idChild;
    (void)dwEventThread;
    (void)dwmsEventTime;
 
    if (idObject != OBJID_WINDOW || !hwnd) return;
 
    /* Only act on popup menu windows */
    WCHAR className[16];
    if (GetClassNameW(hwnd, className, ARRAYSIZE(className)) &&
        wcscmp(className, L"#32768") == 0)
    {
        ApplyAcrylicToWindow(hwnd);
    }
}


/* -------------------------------------------------------------------------- */
/* Dynamic array for menu items                                               */
/* -------------------------------------------------------------------------- */

/**
 * MenuEntry – one "Send To" item.
 *
 * @member path  Heap-alloc'd absolute path via _wcsdup (owner; freed by VectorDestroy).
 * @member icon  32-bit ARGB bitmap for the menu (may be NULL).
 */
typedef struct {
    PWSTR   path;
    HBITMAP icon;
} MenuEntry;

/**
 * MenuVector – simple grow-only array.
 *
 * @member items     Pointer to contiguous buffer (realloc'd).
 * @member count     Elements currently stored.
 * @member capacity  Allocated slots in @items.
 */
typedef struct {
    MenuEntry *items;
    UINT       count;
    UINT       capacity;
} MenuVector;

/**
 * Pointer to the active menu-item vector; used by SendToWndProc to
 * lazily resolve icons on WM_INITMENUPOPUP. Set once in RunSendTo().
 */
static MenuVector *g_menuItems = NULL;

/**
 * VectorEnsureCapacity – make room for at least @need elements.
 *
 * @param vec   Target vector.
 * @param need  Desired minimum capacity.
 * @return      true on success, false on OOM.
 */
static bool VectorEnsureCapacity(MenuVector *vec, UINT need)
{
    if (need <= vec->capacity) {
        return true;
    }

    // Amortized growth: double current capacity (or start at 64), but at least 'need'
    UINT newCap = vec->capacity ? vec->capacity * 2 : MENU_POOL_SIZE;
    if (newCap < need) {
        newCap = need;
    }

    MenuEntry *tmp = realloc(vec->items, newCap * sizeof *tmp);
    if (!tmp) {
        return false;
    }

    // Zero the new tail so later clean-up is safe.
    ZeroMemory(tmp + vec->capacity,
               (newCap - vec->capacity) * sizeof *tmp);

    vec->items    = tmp;
    vec->capacity = newCap;

    return true;
}

/**
 * VectorPush – append a new entry (takes ownership of resources).
 *
 * @param vec   Vector to modify.
 * @param path  Heap path (caller must not reuse).
 * @param icon  Bitmap handle (may be NULL, ownership transferred).
 * @return      true on success, false on OOM — if false the caller still
 *              owns @path/@icon and must free/delete them.
 */
static bool VectorPush(MenuVector *vec, PCWSTR path, HBITMAP icon)
{
    // If capacity growth fails, we do *not* consume the resources.
    if (!VectorEnsureCapacity(vec, vec->count + 1)) {
        return false;
    }

    // Heap-duplicate the path string; caller retains no ownership.
    PWSTR dupPath = _wcsdup(path);
    if (!dupPath) {
        return false;
    }

    vec->items[vec->count].path = dupPath;
    vec->items[vec->count].icon = icon;
    vec->count++;

    return true;
}

/**
 * VectorDestroy – free all stored paths/bitmaps and reset vector to zero.
 *
 * @param vec  Vector to wipe.
 */
static void VectorDestroy(MenuVector *vec)
{
    for (UINT i = 0; i < vec->count; ++i) {
        free(vec->items[i].path);
        if (vec->items[i].icon) {
            DeleteObject(vec->items[i].icon);
        }
    }
    free(vec->items);
    ZeroMemory(vec, sizeof *vec);
}


/* -------------------------------------------------------------------------- */
/* Icons tools                                                                */
/* -------------------------------------------------------------------------- */

/**
 * InitBitmapInfo32 – fill a BITMAPINFO structure for a top-down 32-bit DIB.
 *
 * Centralises the BITMAPINFO setup used by CreateDIBSection32, IconCacheLookup,
 * and IconCacheStore so the format is defined in exactly one place.
 *
 * @param bmi     Pointer to BITMAPINFO to initialise (caller-allocated).
 * @param width   Desired bitmap width in pixels.
 * @param height  Desired bitmap height in pixels.
 */
static void InitBitmapInfo32(BITMAPINFO *bmi, int width, int height)
{
    ZeroMemory(bmi, sizeof *bmi);
    bmi->bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi->bmiHeader.biWidth       = width;
    bmi->bmiHeader.biHeight      = -height;  // negative => top-down orientation
    bmi->bmiHeader.biPlanes      = 1;
    bmi->bmiHeader.biBitCount    = 32;
    bmi->bmiHeader.biCompression = BI_RGB;
}

/**
 * CreateDIBSection32 – allocate a top-down 32-bit DIB of given size.
 *
 * @param width    Desired bitmap width in pixels.
 * @param height   Desired bitmap height in pixels.
 * @param outBits  Optional; if non-NULL, receives the pointer to the raw pixel
 *                 buffer.  Pass NULL when the caller doesn't need direct access
 *                 (e.g. DibFromIcon, which draws via GDI instead).
 * @return         New HBITMAP or NULL on failure.
 */
static HBITMAP CreateDIBSection32(int width, int height, PVOID *outBits)
{
    BITMAPINFO bmi;
    InitBitmapInfo32(&bmi, width, height);

    PVOID pBits = NULL;
    HBITMAP hbm = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, &pBits, NULL, 0);
    if (outBits) {
        *outBits = pBits;
    }

    return hbm;
}

/**
 * DibFromIcon – convert an HICON to a 32-bit ARGB HBITMAP.
 *
 * @param iconHandle  Source HICON (ownership transferred; this function destroys it).
 * @return            HBITMAP or NULL on failure.
 */
static HBITMAP DibFromIcon(HICON iconHandle)
{
    if (!iconHandle) {
        return NULL;
    }

    ICONINFO iconInfo;
    if (!GetIconInfo(iconHandle, &iconInfo)) {
        // failed to extract bitmap handles
        DestroyIcon(iconHandle);
        return NULL;
    }

    // get dimensions from the color bitmap
    BITMAP bmpMetrics;
    if (!GetObject(iconInfo.hbmColor, sizeof bmpMetrics, &bmpMetrics)) {
        DeleteObject(iconInfo.hbmColor);
        DeleteObject(iconInfo.hbmMask);
        DestroyIcon(iconHandle);
        return NULL;
    }

    // allocate DIB section; pixel pointer not needed (drawing via GDI)
    HBITMAP dibBitmap = CreateDIBSection32(bmpMetrics.bmWidth, bmpMetrics.bmHeight, NULL);
    if (dibBitmap) {
        // use global temporary DC
        HGDIOBJ oldObj = SelectObject(hdcIconCache, dibBitmap);
        DrawIconEx(hdcIconCache, 0, 0, iconHandle, bmpMetrics.bmWidth, bmpMetrics.bmHeight, 0, NULL, DI_NORMAL);
        SelectObject(hdcIconCache, oldObj);
    }

    // Cleanup original icon and bitmaps
    DeleteObject(iconInfo.hbmColor);
    DeleteObject(iconInfo.hbmMask);
    DestroyIcon(iconHandle);

    return dibBitmap;
}

/**
 * IconForItem – retrieve shell small icon for a file or directory.
 *
 * Works for both files and directories: SHGetFileInfoW resolves the
 * appropriate icon in either case, including custom folder icons set
 * via desktop.ini and shortcut (.lnk) target icons.
 *
 * @param filePath  Null-terminated wide string path to a file or directory.
 * @return          32-bit ARGB HBITMAP, or NULL on failure.
 */
static HBITMAP IconForItem(PCWSTR filePath)
{
    SHFILEINFOW info;
    UINT flags;

    // primary: real icon from the shell (resolves .lnk targets, desktop.ini, etc.)
    flags = SHGFI_ICON | SHGFI_SMALLICON;
    if (SHGetFileInfoW(filePath, FILE_ATTRIBUTE_NORMAL, &info, sizeof(info), flags)) {
        HBITMAP result = DibFromIcon(info.hIcon);
        if (result) {
            return result;
        }
    }

    // fallback: system image list (includes non-existent/virtual items)
    flags = SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON;
    if (SHGetFileInfoW(filePath, FILE_ATTRIBUTE_NORMAL, &info, sizeof(info), flags)) {
        HBITMAP result = DibFromIcon(info.hIcon);
        if (result) {
            return result;
        }
    }

    return NULL;
}


/* -------------------------------------------------------------------------- */
/* Persistent icon cache                                                      */
/* -------------------------------------------------------------------------- */

/**
 * IconCacheEntry – one cached icon with its invalidation key.
 *
 * @member path       Absolute path of the file this icon belongs to.
 * @member lastWrite  Last-write timestamp of the file when the icon was resolved.
 * @member width      Bitmap width in pixels.
 * @member height     Bitmap height in pixels.
 * @member pixels     Heap-alloc'd raw 32-bit ARGB pixel data (width*height*4 bytes).
 */
typedef struct {
    WCHAR    path[MAX_PATH];
    FILETIME lastWrite;
    int      width;
    int      height;
    BYTE     *pixels;
} IconCacheEntry;

/**
 * IconCache – in-memory store of cached icons loaded from / saved to disk.
 *
 * @member entries   Pointer to contiguous buffer of cache entries.
 * @member count     Number of entries currently stored.
 * @member capacity  Allocated slots in @entries.
 * @member dirty     TRUE if any entry was added or updated since last save.
 */
typedef struct {
    IconCacheEntry *entries;
    UINT            count;
    UINT            capacity;
    bool            dirty;
} IconCache;

/** Global icon cache instance; only active when /C flag is passed. */
static IconCache g_iconCache = { 0 };

/** Whether persistent icon caching is enabled (set via /C flag). */
static bool g_useCacheFlag = false;

/**
 * GetFileLastWriteTime – retrieve the last-write FILETIME for a path.
 *
 * @param path       Null-terminated wide string path.
 * @param outTime    Receives the FILETIME on success.
 * @return           TRUE on success, FALSE on failure.
 */
static BOOL GetFileLastWriteTime(PCWSTR path, FILETIME *outTime)
{
    WIN32_FILE_ATTRIBUTE_DATA attrs;
    if (!GetFileAttributesExW(path, GetFileExInfoStandard, &attrs)) {
        return FALSE;
    }
    *outTime = attrs.ftLastWriteTime;
    return TRUE;
}

/**
 * ResolveCacheFilePath – build the path to "sendto.cache" next to the executable.
 *
 * @param outPath  Buffer of at least MAX_PATH WCHARs to receive the result.
 * @return         TRUE on success, FALSE on failure.
 */
static BOOL ResolveCacheFilePath(WCHAR outPath[MAX_PATH])
{
    if (!GetModuleFileNameW(NULL, outPath, MAX_PATH)) {
        return FALSE;
    }
    PathRemoveFileSpecW(outPath);
    return PathAppendW(outPath, L"sendto.cache");
}

/**
 * IconCacheLoad – read the cache file from disk into g_iconCache.
 *
 * Silently succeeds (with zero entries) if the file doesn't exist or is
 * corrupt.  Only entries whose on-disk format matches CACHE_VERSION are loaded.
 *
 * On partial reads (corrupt file mid-stream), only the successfully-parsed
 * entries are kept and capacity is trimmed to match, so IconCacheSave will
 * not write a stale header count.
 */
static void IconCacheLoad(void)
{
    WCHAR cacheFile[MAX_PATH];
    if (!ResolveCacheFilePath(cacheFile)) {
        return;
    }

    HANDLE hFile = CreateFileW(
        cacheFile, GENERIC_READ, FILE_SHARE_READ, NULL,
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL
    );

    if (hFile == INVALID_HANDLE_VALUE) {
        return;
    }

    // read and validate header
    DWORD bytesRead;
    DWORD magic = 0, version = 0, entryCount = 0;
    if (!ReadFile(hFile, &magic, sizeof magic, &bytesRead, NULL) || magic != CACHE_MAGIC) {
        goto done;
    }

    if (!ReadFile(hFile, &version, sizeof version, &bytesRead, NULL) || version != CACHE_VERSION) {
        goto done;
    }

    if (!ReadFile(hFile, &entryCount, sizeof entryCount, &bytesRead, NULL)) {
        goto done;
    }

    // allocate entries
    g_iconCache.entries = calloc(entryCount, sizeof *g_iconCache.entries);
    if (!g_iconCache.entries) {
        goto done;
    }

    g_iconCache.capacity = entryCount;
    for (DWORD i = 0; i < entryCount; ++i) {
        IconCacheEntry *e = &g_iconCache.entries[i];

        // read path length (in WCHARs including null terminator)
        DWORD pathLen = 0;
        if (!ReadFile(hFile, &pathLen, sizeof pathLen, &bytesRead, NULL)) goto done;
        if (pathLen == 0 || pathLen > MAX_PATH) goto done;

        // read path
        if (!ReadFile(hFile, e->path, pathLen * sizeof(WCHAR), &bytesRead, NULL)) goto done;

        // read timestamp
        if (!ReadFile(hFile, &e->lastWrite, sizeof e->lastWrite, &bytesRead, NULL)) goto done;

        // read dimensions
        if (!ReadFile(hFile, &e->width, sizeof e->width, &bytesRead, NULL)) goto done;
        if (!ReadFile(hFile, &e->height, sizeof e->height, &bytesRead, NULL)) goto done;

        if (e->width <= 0 || e->height <= 0 || e->width > 256 || e->height > 256) goto done;

        // read pixel data
        DWORD pixelSize = (DWORD)(e->width * e->height * 4);
        e->pixels = malloc(pixelSize);
        if (!e->pixels) goto done;
        if (!ReadFile(hFile, e->pixels, pixelSize, &bytesRead, NULL) || bytesRead != pixelSize) {
            free(e->pixels);
            e->pixels = NULL;
            goto done;
        }

        g_iconCache.count++;
    }

done:
    CloseHandle(hFile);

    // Trim capacity to match the number of entries actually loaded.
    // If the file was truncated mid-stream, capacity was set to the header's
    // entryCount but only count entries were parsed successfully.  Aligning
    // the two prevents IconCacheSave from writing a misleading header count.
    g_iconCache.capacity = g_iconCache.count;
}

/**
 * IconCacheSave – write the current g_iconCache contents to disk.
 *
 * Only writes if the cache has been marked dirty (new or updated entries).
 * The header entry-count reflects only entries with valid pixel data, so
 * the reader will never expect more records than were actually written.
 */
static void IconCacheSave(void)
{
    if (!g_iconCache.dirty || g_iconCache.count == 0) {
        return;
    }

    WCHAR cacheFile[MAX_PATH];
    if (!ResolveCacheFilePath(cacheFile)) {
        return;
    }

    HANDLE hFile = CreateFileW(
        cacheFile, GENERIC_WRITE, 0, NULL,
        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL
    );
    if (hFile == INVALID_HANDLE_VALUE) {
        return;
    }

    // Count entries with valid pixel data — only these will be serialised.
    // Entries with NULL pixels (e.g. from a partially-loaded cache or a
    // failed IconCacheStore) are skipped during writing, so the header must
    // advertise the real number of records the reader will encounter.
    DWORD validCount = 0;
    for (UINT i = 0; i < g_iconCache.count; ++i) {
        if (g_iconCache.entries[i].pixels) {
            validCount++;
        }
    }

    if (validCount == 0) {
        CloseHandle(hFile);
        return;
    }

    DWORD written;

    // write header
    DWORD magic   = CACHE_MAGIC;
    DWORD version = CACHE_VERSION;
    WriteFile(hFile, &magic,      sizeof magic,      &written, NULL);
    WriteFile(hFile, &version,    sizeof version,    &written, NULL);
    WriteFile(hFile, &validCount, sizeof validCount, &written, NULL);

    for (UINT i = 0; i < g_iconCache.count; ++i) {
        IconCacheEntry *e = &g_iconCache.entries[i];
        if (!e->pixels) continue;

        // write path length + path
        DWORD pathLen = (DWORD)(wcslen(e->path) + 1);
        WriteFile(hFile, &pathLen, sizeof pathLen, &written, NULL);
        WriteFile(hFile, e->path, pathLen * sizeof(WCHAR), &written, NULL);

        // write timestamp
        WriteFile(hFile, &e->lastWrite, sizeof e->lastWrite, &written, NULL);

        // write dimensions
        WriteFile(hFile, &e->width,  sizeof e->width,  &written, NULL);
        WriteFile(hFile, &e->height, sizeof e->height, &written, NULL);

        // write pixel data
        DWORD pixelSize = (DWORD)(e->width * e->height * 4);
        WriteFile(hFile, e->pixels, pixelSize, &written, NULL);
    }

    CloseHandle(hFile);
}

/**
 * IconCacheDestroy – free all heap memory held by g_iconCache.
 */
static void IconCacheDestroy(void)
{
    for (UINT i = 0; i < g_iconCache.count; ++i) {
        free(g_iconCache.entries[i].pixels);
    }
    free(g_iconCache.entries);
    ZeroMemory(&g_iconCache, sizeof g_iconCache);
}

/**
 * IconCacheEnsureCapacity – make room for at least @need entries in g_iconCache.
 *
 * Mirrors the amortised-doubling strategy of VectorEnsureCapacity, starting
 * at 128 slots when the cache is empty.
 *
 * @param need  Desired minimum capacity.
 * @return      true on success, false on OOM (existing data is untouched).
 */
static bool IconCacheEnsureCapacity(UINT need)
{
    if (need <= g_iconCache.capacity) {
        return true;
    }

    UINT newCap = g_iconCache.capacity ? g_iconCache.capacity * 2 : 128;
    if (newCap < need) {
        newCap = need;
    }

    IconCacheEntry *tmp = realloc(g_iconCache.entries, newCap * sizeof *tmp);
    if (!tmp) {
        return false;
    }

    ZeroMemory(tmp + g_iconCache.capacity,
               (newCap - g_iconCache.capacity) * sizeof *tmp);

    g_iconCache.entries  = tmp;
    g_iconCache.capacity = newCap;

    return true;
}

/**
 * IconCacheLookup – search for a cached icon matching @path and its current
 *                   last-write timestamp.
 *
 * @param path   Null-terminated wide string path of the file to look up.
 * @return       A new HBITMAP created from cached pixel data if a valid entry
 *               exists, or NULL if not found or stale.
 */
static HBITMAP IconCacheLookup(PCWSTR path)
{
    FILETIME ft;
    if (!GetFileLastWriteTime(path, &ft)) {
        return NULL;
    }

    for (UINT i = 0; i < g_iconCache.count; ++i) {
        IconCacheEntry *e = &g_iconCache.entries[i];
        if (_wcsicmp(e->path, path) != 0) {
            continue;
        }

        if (CompareFileTime(&e->lastWrite, &ft) != 0) {
            continue;
        }

        if (!e->pixels) {
            continue;
        }

        // rebuild HBITMAP from cached pixels via shared helper
        PVOID pBits = NULL;
        HBITMAP hbm = CreateDIBSection32(e->width, e->height, &pBits);
        if (hbm && pBits) {
            memcpy(pBits, e->pixels, (size_t)(e->width * e->height * 4));
        }

        return hbm;
    }

    return NULL;
}

/**
 * IconCacheStore – add or update a cache entry for the given path and bitmap.
 *
 * Extracts the raw 32-bit pixel data from @hbm via GetDIBits and stores it
 * alongside the file's current last-write timestamp.  Marks the cache dirty.
 *
 * @param path  Null-terminated wide string path of the file.
 * @param hbm   32-bit ARGB HBITMAP whose pixels will be copied into the cache.
 */
static void IconCacheStore(PCWSTR path, HBITMAP hbm)
{
    if (!hbm) return;

    FILETIME ft;
    if (!GetFileLastWriteTime(path, &ft)) {
        return;
    }

    // get bitmap dimensions
    BITMAP bm;
    if (!GetObject(hbm, sizeof bm, &bm)) return;
    if (bm.bmWidth <= 0 || bm.bmHeight <= 0) return;

    DWORD pixelSize = (DWORD)(bm.bmWidth * bm.bmHeight * 4);
    BYTE *pixels = malloc(pixelSize);
    if (!pixels) return;

    // extract pixel data from the HBITMAP via shared helper
    BITMAPINFO bmi;
    InitBitmapInfo32(&bmi, bm.bmWidth, bm.bmHeight);

    HDC hdc = GetDC(NULL);
    int scanlines = GetDIBits(hdc, hbm, 0, bm.bmHeight, pixels, &bmi, DIB_RGB_COLORS);
    ReleaseDC(NULL, hdc);

    if (scanlines == 0) {
        free(pixels);
        return;
    }

    // check if entry already exists (stale) and update in-place
    for (UINT i = 0; i < g_iconCache.count; ++i) {
        IconCacheEntry *e = &g_iconCache.entries[i];
        if (_wcsicmp(e->path, path) != 0) continue;

        free(e->pixels);
        e->lastWrite = ft;
        e->width     = bm.bmWidth;
        e->height    = bm.bmHeight;
        e->pixels    = pixels;
        g_iconCache.dirty = true;
        return;
    }

    // new entry — grow array if needed
    if (!IconCacheEnsureCapacity(g_iconCache.count + 1)) {
        free(pixels);
        return;
    }

    IconCacheEntry *e = &g_iconCache.entries[g_iconCache.count++];
    StringCchCopyW(e->path, MAX_PATH, path);
    e->lastWrite = ft;
    e->width     = bm.bmWidth;
    e->height    = bm.bmHeight;
    e->pixels    = pixels;
    g_iconCache.dirty = true;
}

/**
 * CachedIconForItem – resolve a file or directory icon, using the persistent
 *                     cache when available.  On cache miss, falls back to
 *                     IconForItem() and stores the result for next time.
 *
 * @param filePath  Null-terminated wide string path to a file or directory.
 * @return          32-bit ARGB HBITMAP, or NULL on failure.
 */
static HBITMAP CachedIconForItem(PCWSTR filePath)
{
    if (g_useCacheFlag) {
        HBITMAP cached = IconCacheLookup(filePath);
        if (cached) return cached;
    }

    HBITMAP icon = IconForItem(filePath);
    if (icon && g_useCacheFlag) {
        IconCacheStore(filePath, icon);
    }

    return icon;
}


/* -------------------------------------------------------------------------- */
/* Foreground window helpers                                                  */
/* -------------------------------------------------------------------------- */

/**
 * FindMainWindowCtx – EnumWindows callback context for FindMainWindowByPid.
 *
 * @member pid     Target process ID to search for.
 * @member result  Receives the first matching HWND, or NULL if not found yet.
 */
typedef struct {
    DWORD pid;
    HWND  result;
} FindMainWindowCtx;

/**
 * FindMainWindowCb – EnumWindows callback: finds the first visible, ownerless
 *                    top-level window belonging to @ctx->pid.
 *
 * "Main window" heuristic: visible + no owner + matching PID.
 * Stops enumeration as soon as a match is found.
 *
 * @param hwnd  Handle to the window currently being enumerated.
 * @param lp    Pointer to a FindMainWindowCtx cast to LPARAM.
 * @return      FALSE to stop enumeration (match found), TRUE to continue.
 */
static BOOL CALLBACK FindMainWindowCb(HWND hwnd, LPARAM lp)
{
    FindMainWindowCtx *ctx = (FindMainWindowCtx *)lp;

    DWORD wPid = 0;
    GetWindowThreadProcessId(hwnd, &wPid);

    // visible, no owner, owned by the target PID → this is the main window
    if (wPid == ctx->pid && IsWindowVisible(hwnd) && !GetWindow(hwnd, GW_OWNER)) {
        ctx->result = hwnd;
        return FALSE; // stop enumeration
    }

    return TRUE;
}

/**
 * FindMainWindowByPid – locate the main visible top-level window of a process.
 *
 * @param pid  Process ID to search for.
 * @return     HWND on success, NULL if no matching window was found.
 */
static HWND FindMainWindowByPid(DWORD pid)
{
    FindMainWindowCtx ctx = { pid, NULL };
    EnumWindows(FindMainWindowCb, (LPARAM)&ctx);
    return ctx.result;
}

/**
 * ForceWindowToForeground – bring @hwnd to the foreground, bypassing Windows'
 *                           foreground-lock restrictions.
 *
 * Windows normally blocks SetForegroundWindow() unless the calling thread
 * already owns the foreground.  Temporarily attaching our thread's input to
 * the target thread's input queue satisfies that restriction and lets the
 * call succeed.
 *
 * @param hwnd  Window to activate.  No-op if NULL or no longer valid.
 */
static void ForceWindowToForeground(HWND hwnd)
{
    if (!hwnd || !IsWindow(hwnd)) {
        return;
    }

    HWND  hwndFg    = GetForegroundWindow();
    DWORD fgThread  = hwndFg ? GetWindowThreadProcessId(hwndFg, NULL) : 0;
    DWORD tgtThread = GetWindowThreadProcessId(hwnd, NULL);

    // attach input queues so SetForegroundWindow is not rejected
    BOOL attached = FALSE;
    if (fgThread && fgThread != tgtThread) {
        attached = AttachThreadInput(fgThread, tgtThread, TRUE);
    }

    ShowWindow(hwnd, SW_RESTORE);
    SetForegroundWindow(hwnd);
    BringWindowToTop(hwnd);

    // always detach, even if SetForegroundWindow failed
    if (attached) {
        AttachThreadInput(fgThread, tgtThread, FALSE);
    }
}

/**
 * ActivateTargetWindow – find and bring to the foreground the window of a
 *                        newly-spawned process (no-args / HandleOpenTarget branch).
 *
 * Waits for @hProcess to become idle, then locates its main window by PID.
 * Retries up to ~2 s to allow slow apps to create their window after idle.
 * Once found, activates it via ForceWindowToForeground.
 *
 * @param hProcess  Handle of the newly-spawned process (must be non-NULL).
 */
static void ActivateTargetWindow(HANDLE hProcess)
{
    // Wait until the app is ready for input, then find its main window.
    // Retry loop covers apps that create their window slightly after
    // WaitForInputIdle returns.
    WaitForInputIdle(hProcess, 3000);

    HWND hwndTarget = NULL;
    DWORD pid = GetProcessId(hProcess);
    for (int retry = 0; retry < 20 && !hwndTarget; ++retry) {
        hwndTarget = FindMainWindowByPid(pid);
        if (!hwndTarget) {
            Sleep(100); // 20 × 100 ms = 2 s max wait
        }
    }

    ForceWindowToForeground(hwndTarget);
}


/* -------------------------------------------------------------------------- */
/* Foreground event hook (drag-and-drop activation)                           */
/* -------------------------------------------------------------------------- */

/**
 * HWND captured by the WinEvent hook when a window takes the foreground
 * during or immediately after a COM drag-and-drop operation.
 *
 * Set by ForegroundEventHookProc, consumed by HandleSendFiles.
 * Protected by single-threaded apartment — no synchronisation needed.
 */
static HWND g_dropForegroundHwnd = NULL;

/**
 * ForegroundEventHookProc – WinEvent callback that records the most recent
 *                           window to receive foreground activation.
 *
 * Installed just before ExecuteDragDrop and uninstalled right after the
 * message-pump wait.  Captures any window that the shell (or the launched
 * application) activates as a result of the drop operation, covering both
 * new processes and already-running apps that merely raise their window.
 *
 * @param hWinEventHook  Handle to the installed event hook (unused).
 * @param event          Event constant (always EVENT_SYSTEM_FOREGROUND here).
 * @param hwnd           Window that received the foreground.
 * @param idObject       Object identifier (unused).
 * @param idChild        Child identifier (unused).
 * @param dwEventThread  Thread that generated the event (unused).
 * @param dwmsEventTime  Timestamp of the event in milliseconds (unused).
 */
static void CALLBACK ForegroundEventHookProc(
    HWINEVENTHOOK hWinEventHook,
    DWORD         event,
    HWND          hwnd,
    LONG          idObject,
    LONG          idChild,
    DWORD         dwEventThread,
    DWORD         dwmsEventTime
) {
    // suppress unused-parameter warnings
    (void)hWinEventHook;
    (void)event;
    (void)idObject;
    (void)idChild;
    (void)dwEventThread;
    (void)dwmsEventTime;

    // only record visible top-level windows (ignore transient popups)
    if (hwnd && IsWindowVisible(hwnd)) {
        g_dropForegroundHwnd = hwnd;
    }
}


/* -------------------------------------------------------------------------- */
/* Window procedure                                                           */
/* -------------------------------------------------------------------------- */

/**
 * SendToWndProc – window procedure for the hidden owner window.
 *
 * Handles WM_INITMENUPOPUP to lazily resolve shell icons for file items
 * just before each popup/submenu is displayed, avoiding the upfront cost
 * of resolving all icons at enumeration time.
 *
 * @param hwnd    Handle to the owner window.
 * @param msg     Message identifier.
 * @param wParam  Additional message information (HMENU for WM_INITMENUPOPUP).
 * @param lParam  Additional message information.
 * @return        Result of message processing; 0 for handled messages,
 *                otherwise the result from DefWindowProcW.
 */
static LRESULT CALLBACK SendToWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    // Forward all unhandled messages to the default procedure
    if (msg != WM_INITMENUPOPUP) {
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    HMENU hMenu = (HMENU)wParam;
    int count = GetMenuItemCount(hMenu);

    // Show loading cursor while shell icons are being resolved
    HCURSOR hPrev = SetCursor(LoadCursor(NULL, IDC_APPSTARTING));

    // Lazily resolve shell icons for each undecorated file item in this popup
    for (int i = 0; i < count; i++) {
        MENUITEMINFOW mii = { sizeof(mii) };
        mii.fMask = MIIM_ID | MIIM_BITMAP | MIIM_SUBMENU;
        if (!GetMenuItemInfoW(hMenu, i, TRUE, &mii)) {
            continue;
        }

        // Skip submenus (directories), already-iconified items, and separators
        if (mii.hSubMenu || mii.hbmpItem || mii.wID == 0) {
            continue;
        }

        // wID is 1-based; map to 0-based vector index
        UINT idx = mii.wID - 1;
        if (idx < g_menuItems->count && !g_menuItems->items[idx].icon) {
            HBITMAP icon = CachedIconForItem(g_menuItems->items[idx].path);
            g_menuItems->items[idx].icon = icon;
            mii.fMask    = MIIM_BITMAP;
            mii.hbmpItem = icon;
            SetMenuItemInfoW(hMenu, i, TRUE, &mii);
        }
    }

    // Restore the cursor that was active before icon resolution
    SetCursor(hPrev);

    return 0;
}


/* -------------------------------------------------------------------------- */
/* Menu population                                                            */
/* -------------------------------------------------------------------------- */

/**
 * SkipEntry – filter out "." / ".." and hidden or system files.
 *
 * @param  findData  WIN32_FIND_DATA of current entry.
 * @return           TRUE if the entry must be ignored.
 */
static BOOL SkipEntry(const WIN32_FIND_DATAW *findData)
{
    return (findData->dwFileAttributes &
            (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM)) ||
           wcscmp(findData->cFileName, L".")  == 0 ||
           wcscmp(findData->cFileName, L"..") == 0;
}

/**
 * AddFileItem – insert a leaf item with icon into a menu and track it in a vector.
 *
 * @param parentMenu Target HMENU to receive the new item.
 * @param fileName   Null-terminated wide string of the file name (with extension).
 * @param bitmap     HBITMAP icon to display, or NULL for no icon.
 * @param commandId  Unique command identifier for the menu entry.
 * @param vec        Pointer to a MenuVector to store the path and bitmap.
 * @param path       File path (stack buffer); copied internally by VectorPush.
 * @return           void; on push failure, frees pathCopy and bitmap if set.
 */
static void AddFileItem(
    HMENU       parentMenu,
    PCWSTR      fileName,
    HBITMAP     bitmap,
    UINT        commandId,
    MenuVector  *vec,
    PCWSTR      path
) {
    // vectorPush may fail; then we must clean up our resources
    if (!VectorPush(vec, path, bitmap)) {
        if (bitmap) DeleteObject(bitmap);
        return;
    }

    // Caption without extension
    WCHAR caption[MAX_PATH];
    StringCchCopyW(caption, ARRAYSIZE(caption), fileName);
    PWSTR ext = PathFindExtensionW(caption);
    if (ext && ext > caption) {
        *ext = L'\0';
    }

    // Prepare the MENUITEMINFO structure for a bitmap + submenu entry
    MENUITEMINFOW itemInfo = { 0 };
    itemInfo.cbSize     = sizeof(itemInfo);
    itemInfo.fMask      = MIIM_ID | MIIM_STRING | MIIM_BITMAP;
    itemInfo.wID        = commandId;
    itemInfo.dwTypeData = caption;
    itemInfo.hbmpItem   = bitmap;

    InsertMenuItemW(parentMenu, commandId, FALSE, &itemInfo);
}

/**
 * AddDirectoryItem – insert a submenu entry with icon and context help ID.
 *
 * @param parentMenu    HMENU to append the new directory item.
 * @param directoryName Null-terminated wide string of the directory label.
 * @param bitmap        HBITMAP icon to display alongside the label.
 * @param subMenu       HMENU handle for the drop-down submenu.
 * @param helpId        DWORD context-help identifier for the submenu.
 * @return              void.
 */
static void AddDirectoryItem(
    HMENU       parentMenu,
    PCWSTR      directoryName,
    HBITMAP     bitmap,
    HMENU       subMenu,
    UINT        helpId
) {
    // Prepare the MENUITEMINFO structure for a bitmap + submenu entry
    MENUITEMINFOW itemInfo = { 0 };
    itemInfo.cbSize      = sizeof(itemInfo);
    itemInfo.fMask       = MIIM_SUBMENU | MIIM_STRING | MIIM_BITMAP;
    itemInfo.hSubMenu    = subMenu;                 // the submenu that drops down
    itemInfo.dwTypeData  = (LPWSTR)directoryName;   // display text
    itemInfo.hbmpItem    = bitmap;                  // icon image (32-bit ARGB)

    // Insert the menu item at the end of parentMenu
    InsertMenuItemW(
        parentMenu,
        GetMenuItemCount(parentMenu),  // position = end
        TRUE,                          // interpret wIDOrPosition as position
        &itemInfo
    );

    // Associate a help/context-ID with the submenu itself
    MENUINFO menuInfo        = { sizeof(menuInfo) };
    menuInfo.fMask           = MIM_HELPID | MIM_STYLE; // we're setting dwContextHelpID
    menuInfo.dwContextHelpID = helpId;                 // ID used to look up real path
    menuInfo.dwStyle         = MNS_AUTODISMISS
                               | MNS_NOTIFYBYPOS;

    SetMenuInfo(subMenu, &menuInfo);
}

/**
 * CompareFindData – qsort comparator for WIN32_FIND_DATAW entries.
 *
 * Sorts directories before files; within each group, sorts by file name
 * using StrCmpLogicalW (natural ordering: file1, file2, file10).
 *
 * @param a  Pointer to first WIN32_FIND_DATAW.
 * @param b  Pointer to second WIN32_FIND_DATAW.
 * @return   Negative if a < b, zero if equal, positive if a > b.
 */
static int CompareFindData(const void *a, const void *b)
{
    const WIN32_FIND_DATAW *fa = (const WIN32_FIND_DATAW *)a;
    const WIN32_FIND_DATAW *fb = (const WIN32_FIND_DATAW *)b;

    BOOL dirA = (fa->dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    BOOL dirB = (fb->dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

    // directories come first
    if (dirA != dirB) {
        return dirB - dirA;
    }

    return StrCmpLogicalW(fa->cFileName, fb->cFileName);
}

/**
 * EnumerateFolder – recursively enumerate a directory, sort the entries
 *                   alphabetically (directories first), and add them to a menu.
 *
 * Collects all valid entries into a temporary heap array, sorts with
 * StrCmpLogicalW for natural ordering, then processes them in order.
 *
 * @param menu        HMENU to which items and submenus will be added.
 * @param directory   Wide‐string path of the folder to enumerate.
 * @param nextCmdId   Pointer to the next command ID; incremented for each item.
 * @param depth       Current recursion depth; stops at MAX_DEPTH.
 * @param items       Pointer to a vector where (path, bitmap) pairs are stored.
 * @return            S_OK on success, or an HRESULT error code on failure.
 */
static HRESULT EnumerateFolder(
    HMENU       menu,
    PCWSTR      directory,
    UINT        *nextCmdId,
    UINT        depth,
    MenuVector  *items
) {
    // Stop if we've reached maximum allowed depth
    if (depth >= MAX_DEPTH) {
        return S_OK;
    }

    // Build the search pattern "directory\\*"
    WCHAR pattern[MAX_LOCAL_PATH];
    if (!PathCombineW(pattern, directory, L"*")) {
        return E_FAIL;
    }

    // Begin file enumeration
    WIN32_FIND_DATAW findData;
    HANDLE hFind = FindFirstFileExW(
        pattern,
        FindExInfoBasic,
        &findData,
        FindExSearchNameMatch,
        NULL,
        FIND_FIRST_EX_LARGE_FETCH
    );
    if (hFind == INVALID_HANDLE_VALUE) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // --- Phase 1: collect all valid entries into a temporary heap array ---
    UINT entryCount = 0;
    UINT entryCapacity = 32;
    WIN32_FIND_DATAW *entries = malloc(entryCapacity * sizeof *entries);
    if (!entries) {
        FindClose(hFind);
        return E_OUTOFMEMORY;
    }

    do {
        if (SkipEntry(&findData)) {
            continue;
        }

        // grow array if needed
        if (entryCount >= entryCapacity) {
            UINT newCap = entryCapacity * 2;
            WIN32_FIND_DATAW *tmp = realloc(entries, newCap * sizeof *tmp);
            if (!tmp) {
                break;  // process what we have so far
            }
            entries = tmp;
            entryCapacity = newCap;
        }

        entries[entryCount++] = findData;
    } while (FindNextFileW(hFind, &findData));

    FindClose(hFind);

    // --- Phase 2: sort — directories first, then alphabetical within each group ---
    if (entryCount > 1) {
        qsort(entries, entryCount, sizeof *entries, CompareFindData);
    }

    // --- Phase 3: add sorted entries to the menu and vector ---
    for (UINT i = 0; i < entryCount; ++i) {
        WIN32_FIND_DATAW *entry = &entries[i];

        // Build full child path
        WCHAR childPath[MAX_LOCAL_PATH];
        if (!PathCombineW(childPath, directory, entry->cFileName)) {
            continue;
        }

        if (entry->dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            // For subdirectories, create a new submenu
            HMENU subMenu = CreatePopupMenu();
            if (!subMenu) {
                continue;
            }

            // Retrieve icon bitmap via unified cache-aware resolver
            HBITMAP icon = CachedIconForItem(childPath);

            // Store in vector first — if this fails, nothing was added to the
            // menu yet so we can cleanly bail out without orphaning resources.
            if (!VectorPush(items, childPath, icon)) {
                if (icon) DeleteObject(icon);
                DestroyMenu(subMenu);
                continue;
            }

            // Insert directory item with icon and context-help ID
            AddDirectoryItem(menu, entry->cFileName,
                             icon, subMenu, *nextCmdId);
            (*nextCmdId)++;

            // Recurse into the subdirectory
            EnumerateFolder(subMenu, childPath, nextCmdId, depth + 1, items);
        } else {
            // Icon resolved lazily in WM_INITMENUPOPUP via CachedIconForItem
            HBITMAP icon = NULL;

            // For files, insert a regular file item
            AddFileItem(menu, entry->cFileName, icon, (*nextCmdId)++,
                        items, childPath);
        }
    }

    free(entries);

    return S_OK;
}


/* -------------------------------------------------------------------------- */
/* IDataObject builder & drop helpers                                         */
/* -------------------------------------------------------------------------- */

/**
 * PathToPIDL - Convert a filesystem path into a COM PIDL (item identifier list).
 *
 * @param hwndOwner   HWND used as the parsing context (may be NULL).
 * @param pszPath     Null-terminated wide string path to file or folder.
 * @return            Newly allocated PIDL on success (caller must free with CoTaskMemFree),
 *                    or NULL on failure.
 */
static LPITEMIDLIST PathToPIDL(HWND hwndOwner, PCWSTR pszPath)
{
    LPITEMIDLIST pidl = NULL;
    ULONG eaten = 0;
    DWORD attrs = 0;

    // Parse the display name into a PIDL
    const HRESULT hr = desktopShellFolder->lpVtbl->ParseDisplayName(
        desktopShellFolder,
        hwndOwner,
        NULL,
        (PWSTR)pszPath,
        &eaten,
        &pidl,
        &attrs
    );

    return SUCCEEDED(hr) ? pidl : NULL;
}

/**
 * GetShellInterfaceForPIDLs – bind an array of absolute PIDLs to a COM interface
 *                             obtained from their common parent shell folder.
 *
 * IMPORTANT: all PIDLs in @pidlArray MUST reside in the same parent folder.
 * The function calls SHBindToParent on each PIDL but only keeps the *last*
 * parent IShellFolder, then passes every child-relative ID to that single
 * folder's GetUIObjectOf.  If the PIDLs come from different parents, the
 * child IDs from earlier iterations will be meaningless in the context of
 * the last parent and GetUIObjectOf will fail or return a corrupt object.
 *
 * This is safe for all current callers:
 *   - IDataObject (source files): Explorer's "Send To" always selects files
 *     within a single folder view, so all paths share the same parent.
 *   - IDropTarget (destination):  always called with exactly one path.
 *
 * @param hwndOwner     HWND used for binding context (may be NULL).
 * @param pidlArray     Array of absolute PIDLs (must share same parent folder).
 * @param pidlCount     Number of entries in pidlArray[].
 * @param interfaceId   IID of the desired interface (e.g. IID_IDataObject).
 * @param ppvInterface  Receives the interface pointer; caller must Release().
 * @return              S_OK on success or an HRESULT error code.
 */
static HRESULT GetShellInterfaceForPIDLs(
    HWND           hwndOwner,
    LPITEMIDLIST   *pidlArray,
    int            pidlCount,
    REFIID         interfaceId,
    void           **ppvInterface
) {
    // Allocate array to hold child IDs relative to each parent shell folder
    LPITEMIDLIST *childIDs = calloc(pidlCount, sizeof *childIDs);
    if (!childIDs) {
        return E_OUTOFMEMORY;
    }

    *ppvInterface = NULL;
    LPSHELLFOLDER parentFolder = NULL;
    HRESULT hr = S_OK;

    // Bind each PIDL to its parent folder, saving the child-relative ID
    for (int pidlIndex = 0; pidlIndex < pidlCount; ++pidlIndex) {
        hr = SHBindToParent(
            pidlArray[pidlIndex],
            &IID_IShellFolder,
            (void**)&parentFolder,
            &childIDs[pidlIndex]
        );
        if (FAILED(hr)) {
            SAFE_RELEASE(parentFolder);
            free(childIDs);
            return hr;
        }

        // Release intermediate shell folder except for the last one
        if (pidlIndex < pidlCount - 1) {
            SAFE_RELEASE(parentFolder);
        }
    }

    // Now request the desired interface from the final parent shell folder
    hr = parentFolder->lpVtbl->GetUIObjectOf(
        parentFolder,
        hwndOwner,
        pidlCount,
        (LPCITEMIDLIST*)childIDs,
        interfaceId,
        NULL,
        ppvInterface
    );

    // Clean up
    SAFE_RELEASE(parentFolder);
    free(childIDs);

    return hr;
}

/**
 * GetShellInterfaceForPaths – convert an array of file or folder paths into
 *                             PIDLs and bind them to a requested COM interface.
 *
 * All paths must reside in the same parent directory (see the same-parent
 * constraint documented in GetShellInterfaceForPIDLs).
 *
 * @param hwndOwner     Window handle used as context for PIDL parsing.
 * @param paths         Array of wide-string file/folder paths.
 * @param pathCount     Number of entries in paths[].
 * @param interfaceId   IID of the desired COM interface
 *                      (e.g. IID_IDataObject or IID_IDropTarget).
 * @param ppvInterface  Receives the interface pointer; caller must Release().
 * @return              S_OK on success or an HRESULT error code.
 */
static HRESULT GetShellInterfaceForPaths(
    HWND    hwndOwner,
    PCWSTR  *paths,
    int     pathCount,
    REFIID  interfaceId,
    void    **ppvInterface
) {
    if (!paths || pathCount <= 0) {
        return E_INVALIDARG;
    }

    // allocate array of PIDLs
    LPITEMIDLIST *pidlArray = calloc(pathCount, sizeof *pidlArray);
    if (!pidlArray) {
        return E_OUTOFMEMORY;
    }

    *ppvInterface = NULL;

    // convert each path to a PIDL
    for (int pathIndex = 0; pathIndex < pathCount; ++pathIndex) {
        pidlArray[pathIndex] = PathToPIDL(hwndOwner, paths[pathIndex]);
        if (!pidlArray[pathIndex]) {
            // cleanup on failure
            for (int cleanupIndex = 0; cleanupIndex < pathIndex; ++cleanupIndex) {
                CoTaskMemFree(pidlArray[cleanupIndex]);
            }
            free(pidlArray);
            return E_FAIL;
        }
    }

    // bind to the requested COM interface on these PIDLs
    const HRESULT hr = GetShellInterfaceForPIDLs(
        hwndOwner,
        pidlArray,
        pathCount,
        interfaceId,
        ppvInterface
    );

    // free all PIDLs
    for (int cleanupIndex = 0; cleanupIndex < pathCount; ++cleanupIndex) {
        CoTaskMemFree(pidlArray[cleanupIndex]);
    }
    free(pidlArray);

    return hr;
}

/**
 * ExecuteDropOperation - Perform COM drag-enter, drop, and leave on the given drop target.
 *
 * @param dataObj    IDataObject with drag data.
 * @param dropTarget IDropTarget for the drop target.
 */
static void ExecuteDropOperation(IDataObject *dataObj, IDropTarget *dropTarget)
{
    // guard against null COM pointers
    if (!dataObj || !dropTarget) {
        return;
    }

    const POINTL pt = { 0 };
    DWORD effect = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;

    // drag into the target
    const HRESULT hrEnter = dropTarget->lpVtbl->DragEnter(
        dropTarget, dataObj, MK_LBUTTON, pt, &effect
    );

    // if accepted, perform drop
    if (SUCCEEDED(hrEnter) && effect) {
        dropTarget->lpVtbl->Drop(
            dropTarget, dataObj, MK_LBUTTON, pt, &effect
        );
    } else {
        dropTarget->lpVtbl->DragLeave(dropTarget);
    }
}

/**
 * ExecuteDragDrop - Perform a COM drag-and-drop of the files passed in argv[1…argc-1]
 *                   onto the target described by @entry.
 *
 * @param owner   HWND of the hidden owner window for COM calls.
 * @param entry   Pointer to the MenuEntry containing the target .path.
 * @param argc    Argument count (program name + file paths).
 * @param argv    Array of PWSTR; argv[1…] are source file paths.
 */
static void ExecuteDragDrop(HWND owner, const MenuEntry *entry, int argc, PWSTR *argv)
{
    IDataObject *pDataObj    = NULL;
    IDropTarget *pDropTarget = NULL;

    // Build IDataObject from the array of source file paths
    HRESULT hr = GetShellInterfaceForPaths(
        owner,
        argv + 1, // skip argv[0]
        argc - 1, // number of files
        &IID_IDataObject,
        (void**)&pDataObj
    );
    if (FAILED(hr) || !pDataObj) {
        return;
    }

    // Retrieve the IDropTarget for the destination folder/link
    hr = GetShellInterfaceForPaths(
        owner,
        (PCWSTR[]){ entry->path },  // address of the single target path
        1,                          // one drop target
        &IID_IDropTarget,
        (void**)&pDropTarget
    );

    if (SUCCEEDED(hr) && pDropTarget) {
        ExecuteDropOperation(pDataObj, pDropTarget);
        SAFE_RELEASE(pDropTarget);
    }

    // Clean up the data object
    SAFE_RELEASE(pDataObj);
}


/* -------------------------------------------------------------------------- */
/* Program entry                                                              */
/* -------------------------------------------------------------------------- */

/**
 * InitializeApplication – initialize OLE and standard common controls.
 *
 * @return TRUE if both OleInitialize and InitCommonControlsEx succeed; FALSE otherwise.
 */
static BOOL InitializeApplication(void)
{
    // initialize OLE for drag-and-drop COM interfaces
    if (FAILED(OleInitialize(NULL))) {
        OutputDebugStringW(L"[SendTo+] OleInitialize failed\n");
        return FALSE;
    }

    // set up common controls required for owner-drawn menus
    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_STANDARD_CLASSES };
    if (!InitCommonControlsEx(&icc)) {
        OutputDebugStringW(L"[SendTo+] InitCommonControlsEx failed\n");
        goto failed;
    }

    // complete drag and drop setup
    const HRESULT hr = SHGetDesktopFolder(&desktopShellFolder);
    if (FAILED(hr)) {
        OutputDebugStringW(L"[SendTo+] SHGetDesktopFolder failed\n");
        goto failed;
    }

    // cache setup
    hdcIconCache = CreateCompatibleDC(NULL);

    // add theme support
    OptInDarkPopupMenus();

    return TRUE;

failed:
    OleUninitialize();
    return FALSE;
}

/**
 * ParseCommandLine - Parses switches and returns a clean argv[].
 *
 * Recognised switches:
 *   /D <dir>  – override the SendTo directory.
 *   /C        – enable persistent icon cache (sendto.cache).
 *   /?  -?    – show usage and exit.
 *
 * @param  rawArgc      Argument count from CommandLineToArgvW().
 * @param  rawArgv      Argument vector from CommandLineToArgvW().
 * @param  outDir       Receives a malloc'd wide string if "/D <dir>" was supplied.
 *                      Caller must free() it when done.  May be NULL.
 * @param  outUseCache  Receives TRUE if "/C" was supplied.
 * @param  outArgc      Receives the new argument count.
 * @param  outArgv      Receives a malloc'd PWSTR[] of length outArgc:
 *                      [0] = rawArgv[0] (exe path)
 *                      [1..] = only the file arguments.
 *                      Caller must free() the array (not the strings).
 * @return              true on success (or after showing help), false on error.
 *                      If help was shown, returns false so caller can exit.
 */
static bool ParseCommandLine(
    int     rawArgc,
    PWSTR   *rawArgv,
    PWSTR   *outDir,
    bool    *outUseCache,
    int     *outArgc,
    PWSTR   **outArgv
) {
    *outDir      = NULL;
    *outUseCache = false;
    *outArgc     = 1;                   // always keep exe @ index 0
    *outArgv     = NULL;

    // allocate worst-case full array
    PWSTR *temp = malloc(rawArgc * sizeof *temp);
    if (!temp) {
        return false;
    }

    // keep program name
    temp[0] = rawArgv[0];

    // scan switches & collect files
    for(int paramIndex = 1; paramIndex < rawArgc; ++paramIndex) {
        PWSTR param = rawArgv[paramIndex];

        // help?
        if (_wcsicmp(param, L"/?")==0 || _wcsicmp(param, L"-?")==0) {
            ERR_BOX(L"Usage: SendTo+ [/D <directory>] [/C] [<file1> <file2> ...]\n\n"
                    L"  /D <dir>  Override the SendTo folder path.\n"
                    L"  /C        Enable persistent icon cache.");
            goto failed;
        }

        // override SendTo directory?
        if (_wcsicmp(param, L"/D")==0) {
            if (paramIndex + 1 < rawArgc) {
                *outDir = _wcsdup(rawArgv[++paramIndex]);
                if (!*outDir) {
                    goto failed;
                }
            } else {
                ERR_BOX(L"Error: /D requires a directory path.\n"
                        L"Usage: SendTo+ [/D <directory>] [/C] [<file1> <file2> ...]");
                goto failed;
            }

            continue;
        }

        // enable persistent icon cache?
        if (_wcsicmp(param, L"/C")==0) {
            *outUseCache = true;
            continue;
        }

        // otherwise treat as file
        temp[(*outArgc)++] = param;
    }

    // shrink-to-fit (optional)
    if (*outArgc < rawArgc) {
        PWSTR *shrunk = realloc(temp, (*outArgc)*sizeof *shrunk);
        if (shrunk) {
            temp = shrunk;
        }
    }

    *outArgv = temp;
    return true;

failed:
    // clean up /D allocation if it was set before the error
    free(*outDir);
    *outDir = NULL;

    free(temp);
    return false;
}

/**
 * ResolveSendToDirectory – build full path to "<exe folder>\\sendto" with \\?\ prefix.
 *
 * @return malloc'd wide string on success (must be freed), or NULL on failure.
 */
static PWSTR ResolveSendToDirectory(void)
{
    // get path to our own executable
    WCHAR exeFolder[MAX_PATH];
    if (!GetModuleFileNameW(NULL, exeFolder, ARRAYSIZE(exeFolder))) {
        OutputDebugStringW(L"[SendTo+] GetModuleFileNameW failed\n");
        return NULL;
    }

    PathRemoveFileSpecW(exeFolder);

    const WCHAR suffix[] = L"\\sendto";
    size_t totalLen = wcslen(exeFolder) + wcslen(suffix) + 1;

    // allocate buffer for "<exeFolder>\\sendto"
    PWSTR buffer = malloc(totalLen * sizeof(WCHAR));
    if (!buffer) {
        OutputDebugStringW(L"[SendTo+] malloc failed\n");
        return NULL;
    }

    // combine directory + suffix
    StringCchPrintfW(buffer, totalLen, L"%s%s", exeFolder, suffix);

    return buffer;
}

/**
 * ValidateSendToDirectory – verify path exists and is a directory.
 *
 * @param path wide string path to verify.
 * @return TRUE if path exists and is directory; FALSE otherwise (shows message box).
 */
static BOOL ValidateSendToDirectory(PCWSTR path)
{
    // check existence and directory attribute
    if (!path || !PathFileExistsW(path) || !PathIsDirectoryW(path)) {
        ERR_BOX(L"Cannot find 'sendto' folder next to the executable.");
        return FALSE;
    }

    return TRUE;
}

/**
 * BuildSendToMenu – create popup menu and populate it from sendto folder.
 *
 * @param sendToDir directory to enumerate.
 * @param outPopup  receives HMENU of created popup.
 * @param outItems  receives MenuVector of menu items.
 * @return TRUE on success; FALSE on failure.
 */
static BOOL BuildSendToMenu(PCWSTR sendToDir, HMENU *outPopup, MenuVector *outItems)
{
    // create empty popup
    *outPopup = CreatePopupMenu();
    *outItems = (MenuVector){ 0 };

    // pre-reserve capacity in one go to avoid repeated reallocs
    VectorEnsureCapacity(outItems, MENU_POOL_SIZE);

    // recursively fill menu and items vector
    UINT initialCmdId = 1; // start command IDs at 1
    const HRESULT hr = EnumerateFolder(
        *outPopup,
        sendToDir,
        &initialCmdId,     // pass address of a real variable
        0,
        outItems
    );

    if (FAILED(hr)) {
        ERR_BOX(L"Failed to enumerate the SendTo folder.");
        return FALSE;
    }

    // If no items were found in the SendTo folder, inform the user
    if (outItems->count == 0) {
        ERR_BOX(L"No items were found in the SendTo folder.");
        return FALSE;
    }

    return TRUE;
}

/**
 * CreateHiddenOwnerWindow – register and create a hidden window for menu ownership.
 *
 * @param hInstance application instance handle.
 * @return HWND of hidden window, or NULL on failure.
 */
static HWND CreateHiddenOwnerWindow(HINSTANCE hInstance)
{
    const wchar_t CLASS_NAME[] = L"SendToOwnerWindow";
    const WNDCLASSEXW wc = {
        .cbSize        = sizeof(WNDCLASSEXW),
        .lpfnWndProc   = SendToWndProc,
        .hInstance     = hInstance,
        .lpszClassName = CLASS_NAME
    };
    if (!RegisterClassExW(&wc)) {
        ERR_BOX(L"Call to RegisterClassExW() failed");
        return NULL;
    }

    // create invisible popup window as menu owner
    HWND hwnd = CreateWindowExW(
        0,
        CLASS_NAME,
        NULL,
        WS_POPUP,
        0, 0, 0, 0,
        HWND_DESKTOP,
        NULL,
        hInstance,
        NULL
    );
    if (!hwnd) {
        ERR_BOX(L"Call to CreateWindowExW() failed");
        return NULL;
    }

    // menu logic
    ShowWindow(hwnd, SW_HIDE);
    SetForegroundWindow(hwnd);

    return hwnd;
}

/**
 * DisplaySendToMenu – show popup at cursor and return selected command ID.
 *
 * @param popup HMENU to display.
 * @param owner HWND of hidden owner window.
 * @return chosen command ID, or 0 if none.
 */
static UINT DisplaySendToMenu(HMENU popup, HWND owner)
{
    // Resolve SetWindowCompositionAttribute (once)
    EnsureAcrylicProc();
 
    // Install event hook to apply acrylic to menu windows as they appear
    HWINEVENTHOOK acrylicHook = NULL;
    if (pSetWindowCompositionAttribute) {
        acrylicHook = SetWinEventHook(
            EVENT_OBJECT_SHOW,         /* eventMin */
            EVENT_OBJECT_SHOW,         /* eventMax */
            NULL,                      /* no DLL – in-process callback */
            AcrylicMenuEventProc,      /* callback */
            GetCurrentProcessId(),     /* only our process */
            GetCurrentThreadId(),      /* only this thread */
            WINEVENT_OUTOFCONTEXT
        );
    }
 
    // Get cursor position for menu location
    POINT cursor;
    GetCursorPos(&cursor);
 
    const UINT cmd = TrackPopupMenuEx(
        popup,
        TPM_RETURNCMD | TPM_LEFTALIGN | TPM_LEFTBUTTON,
        cursor.x, cursor.y,
        owner,
        NULL
    );
 
    // Tear down acrylic hook
    if (acrylicHook) {
        UnhookWinEvent(acrylicHook);
    }
 
    // Exit menu-mode to restore normal input
    PostMessage(owner, WM_NULL, 0, 0);
 
    return cmd;
}

/**
 * HandleOpenTarget – no-args branch: open a folder or .lnk target and bring
 *                    its window to the foreground.
 *
 * Uses ShellExecuteExW (instead of ShellExecuteW) so we can obtain the new
 * process handle, wait until its window is ready, and force it to the
 * foreground via ActivateTargetWindow.
 *
 * @param owner  HWND of the hidden owner window (used as ShellExecute context).
 * @param item   Selected MenuEntry whose .path is the file/folder to open.
 */
static void HandleOpenTarget(HWND owner, const MenuEntry *item)
{
    OutputDebugStringW(L"[SendTo+] no args: open folder/link\n");

    // Allow the new process to call SetForegroundWindow on itself if it wants to.
    AllowSetForegroundWindow(ASFW_ANY);

    SHELLEXECUTEINFOW sei  = { sizeof(sei) };
    sei.fMask  = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC;
    sei.hwnd   = owner;
    sei.lpFile = item->path;
    sei.nShow  = SW_SHOWNORMAL;

    if (ShellExecuteExW(&sei) && sei.hProcess) {
        // Wait for the new process to create its window, then
        // force it to the foreground like any normal menu action.
        ActivateTargetWindow(sei.hProcess);
        CloseHandle(sei.hProcess);
    }
}

/**
 * HandleSendFiles – drag-and-drop branch: send @argv[1…] to the selected
 *                   target and bring its window to the foreground.
 *
 * Uses a SetWinEventHook to listen for EVENT_SYSTEM_FOREGROUND events
 * during and immediately after the COM drop operation.  This is event-driven
 * (no process snapshots, no title heuristics) and covers both newly-launched
 * processes and already-running apps that simply raise their existing window.
 *
 * The hook is installed before the drop and torn down after a short
 * message-pump wait, which gives the shell enough time to launch the target
 * application and have it take the foreground.
 *
 * @param owner  HWND of the hidden owner window (used as COM context).
 * @param item   Selected MenuEntry whose .path is the drop-target folder/app.
 * @param argc   Full argument count (argv[0] = exe, argv[1…] = source files).
 * @param argv   Argument vector; argv[1…argc-1] are the files to send.
 */
static void HandleSendFiles(HWND owner, const MenuEntry *item, int argc, PWSTR *argv)
{
    OutputDebugStringW(L"[SendTo+] with args: perform drag-and-drop\n");

    // Allow the drop-target process to call SetForegroundWindow on itself
    AllowSetForegroundWindow(ASFW_ANY);

    // Reset the shared HWND that the hook callback will populate
    g_dropForegroundHwnd = NULL;

    // Install an event hook that captures any foreground activation caused
    // by the drop (works for new processes and existing windows alike)
    HWINEVENTHOOK foregroundHook = SetWinEventHook(
        EVENT_SYSTEM_FOREGROUND,    // eventMin
        EVENT_SYSTEM_FOREGROUND,    // eventMax
        NULL,                       // no DLL – callback lives in this process
        ForegroundEventHookProc,    // callback
        0,                          // all processes
        0,                          // all threads
        WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS
    );
    if (!foregroundHook) {
        OutputDebugStringW(L"[SendTo+] SetWinEventHook failed; "
                           L"target window may stay in background\n");
    }

    // Perform the actual COM drag-and-drop operation
    ExecuteDragDrop(owner, item, argc, argv);

    // Pump messages briefly so the hook callback can fire.
    // The shell may need a moment to launch the target process and have it
    // call SetForegroundWindow, so we poll up to ~2 s (40 × 50 ms).
    for (int waitCycle = 0; waitCycle < 40 && !g_dropForegroundHwnd; ++waitCycle) {
        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
            DispatchMessage(&msg);
        }
        Sleep(50);
    }

    // Tear down the hook now that we have (or timed-out waiting for) the HWND
    if (foregroundHook) {
        UnhookWinEvent(foregroundHook);
    }

    // Force the captured window to the foreground in case Windows demoted it
    ForceWindowToForeground(g_dropForegroundHwnd);
}

/**
 * SetupIconCache – apply the /C flag globally and load the cache file from disk.
 *
 * Centralises the two-step setup so RunSendTo stays at the orchestration level.
 *
 * @param useCache  TRUE if the /C flag was passed on the command line.
 */
static void SetupIconCache(bool useCache)
{
    g_useCacheFlag = useCache;
    if (g_useCacheFlag) {
        IconCacheLoad();
    }
}

/**
 * TeardownIconCache – persist any dirty cache entries to disk and free memory.
 *
 * No-op when caching is disabled (g_useCacheFlag == FALSE).
 */
static void TeardownIconCache(void)
{
    if (!g_useCacheFlag) {
        return;
    }

    IconCacheSave();
    IconCacheDestroy();
}

/**
 * ShutdownApplication – release global GDI and COM resources, uninitialise OLE.
 *
 * Must be called once at the end of RunSendTo, after all COM operations
 * (drag-and-drop, ShellExecuteEx, etc.) have completed.
 */
static void ShutdownApplication(void)
{
    // release the temporary DC used by DibFromIcon for icon rendering
    if (hdcIconCache) {
        DeleteDC(hdcIconCache);
        hdcIconCache = NULL;
    }

    SAFE_RELEASE(desktopShellFolder);
    OleUninitialize();
}

/**
 * RunSendTo – perform full SendTo+ workflow.
 *
 * Uses a single-exit cleanup pattern (goto) so that all heap allocations,
 * menu handles, and vector storage are released on every code path —
 * including early failures and user cancellation.
 *
 * @param hInstance    application instance.
 * @param argc         argument count.
 * @param argv         argument vector (wide).
 * @return             exit code (0 success, non-zero on error).
 */
static int RunSendTo(HINSTANCE hInstance, int argc, PWSTR *argv)
{
    int exitCode         = EXIT_FAILURE;
    int cleanArgc        = 0;
    PWSTR *cleanArgv     = NULL;
    PWSTR sendToDir      = NULL;
    bool useCache        = false;
    HMENU popupMenu      = NULL;
    HWND owner           = NULL;
    MenuVector menuItems = { 0 };

    if (!ParseCommandLine(argc, argv, &sendToDir, &useCache, &cleanArgc, &cleanArgv)) {
        goto cleanup;
    }

    SetupIconCache(useCache);

    if (!sendToDir) {
        sendToDir = ResolveSendToDirectory();
    }

    if (!ValidateSendToDirectory(sendToDir)) {
        goto cleanup;
    }

    // build popup menu and items
    if (!BuildSendToMenu(sendToDir, &popupMenu, &menuItems)) {
        goto cleanup;
    }

    // create hidden owner window
    owner = CreateHiddenOwnerWindow(hInstance);
    if (!owner) {
        goto cleanup;
    }

    // display menu and handle selection
    g_menuItems = &menuItems;
    UINT choice = DisplaySendToMenu(popupMenu, owner);
    if (choice) {
        MenuEntry *item = &menuItems.items[choice - 1];

        if (cleanArgc > 1) {
            // with args: perform drag-and-drop
            HandleSendFiles(owner, item, cleanArgc, cleanArgv);
        } else {
            // no args: open folder/link
            HandleOpenTarget(owner, item);
        }
    }

    exitCode = EXIT_SUCCESS;

cleanup:
    // persist icon cache to disk if it was modified
    TeardownIconCache();

    // destroy menu tree and item vector (safe even if never initialised)
    if (popupMenu) {
        DestroyMenu(popupMenu);
    }
    VectorDestroy(&menuItems);

    // free heap-allocated argument data
    free(sendToDir);
    free(cleanArgv);

    // destroy hidden owner window
    if (owner) {
        DestroyWindow(owner);
    }

    // release COM desktop folder and uninitialise OLE
    ShutdownApplication();

    return exitCode;
}

/**
 * wWinMain – main entry point for SendTo+.
 *
 * @param hInstance     application instance.
 * @param hPrevInstance unused.
 * @param lpCmdLine     unused.
 * @param nCmdShow      unused.
 * @return exit code (0 success, non-zero on error).
 */
int WINAPI wWinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ PWSTR     lpCmdLine,
    _In_ int       nCmdShow
) {
    // suppress unused-parameter warnings
    (void)hPrevInstance;
    (void)lpCmdLine;
    (void)nCmdShow;

    if (!InitializeApplication()) {
        return EXIT_FAILURE;
    }

    // better params support
    int rawArgc;
    PWSTR *rawArgv = CommandLineToArgvW(GetCommandLineW(), &rawArgc);
    if (!rawArgv) {
        ShutdownApplication();
        return EXIT_FAILURE;
    }

    int exitCode = RunSendTo(hInstance, rawArgc, rawArgv);

    // CommandLineToArgvW allocates via LocalAlloc — must be freed here
    LocalFree(rawArgv);

    return exitCode;
}
