/*
 * sendto.c – Unicode-only reworked version of SendTo+ (MSVC)
 * Copyright (c) 2025 DSR! <xchwarze@gmail.com>
 * based on https://github.com/lifenjoiner/sendto-plus
 */

#define WIN32_LEAN_AND_MEAN
#define _WIN32_WINNT 0x0601 /* Windows 7 */
#define _WIN32_IE    0x0600 /* Needed for IID_IImageList */

#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <commoncontrols.h> /* for IID_IImageList */
#include <unknwn.h>
#include <shellapi.h>
#include <strsafe.h>
#include <shobjidl.h>
#include <stdbool.h>
#include <dwmapi.h>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "dwmapi.lib")

#define MAX_DEPTH 4
#define MAX_LOCAL_PATH 32767


/* -------------------------------------------------------------------------- */
/* Utility macros                                                             */
/* -------------------------------------------------------------------------- */

#define SAFE_RELEASE(p)     do { if (p) { (p)->lpVtbl->Release(p); (p) = NULL; } } while (0)
#define RETURN_IF_FAILED(h) do { HRESULT _hr = (h); if (FAILED(_hr)) return _hr; } while (0)
#define BOOL_IF_FAILED(h)   do { if (FAILED(h)) return FALSE; } while(0)


/* -------------------------------------------------------------------------- */
/* Helpers                                                                    */
/* -------------------------------------------------------------------------- */

static LPSHELLFOLDER desktopShellFolder = NULL;

#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20  // Missing in older SDKs
#endif

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

    // Ordinals are stable since 1809
    typedef PreferredAppMode (WINAPI *SetPreferredAppMode_t)(PreferredAppMode);
    typedef BOOL (WINAPI *FlushMenuThemes_t)(void);

    HMODULE uxThemeModule = LoadLibraryW(L"uxtheme.dll");
    if (!uxThemeModule) {
        return;
    }

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
}

/**
 * AppUsesDarkTheme
 *
 * Return TRUE when the user has “Dark” selected for *Apps* in Settings.
 * Relies on uxtheme!ShouldAppsUseDarkMode (exported by ordinal 132).
 */
static inline BOOL AppUsesDarkTheme(void)
{
    // uxtheme.dll is always loaded in a GUI process, but be safe.
    HMODULE uxThemeModule = GetModuleHandleW(L"uxtheme.dll");
    if (!uxThemeModule) {
        uxThemeModule = LoadLibraryW(L"uxtheme.dll");
    }
    if (!uxThemeModule) {
        return FALSE;
    }

    typedef BOOL (WINAPI *ShouldAppsUseDarkMode_t)(void);
    ShouldAppsUseDarkMode_t ShouldAppsUseDarkMode =
        (ShouldAppsUseDarkMode_t)GetProcAddress(
            uxThemeModule,
            MAKEINTRESOURCEA(132)); // ordinal-only export

    return ShouldAppsUseDarkMode && ShouldAppsUseDarkMode();
}

/**
 * ApplyDarkThemeIfNeeded
 *
 * Toggle DWMWA_USE_IMMERSIVE_DARK_MODE so the window caption, context menus
 * and drop-shadows match the current per-app theme.
 *
 * @param hwnd  Window handle (can be your hidden owner window).
 */
static inline void ApplyDarkThemeIfNeeded(HWND hwnd)
{
    if (!hwnd) {
        return;
    }

    const BOOL enableDark = AppUsesDarkTheme();

    // Ignored by builds < 18362 — safe no-op there.
    (void)DwmSetWindowAttribute(
        hwnd,
        DWMWA_USE_IMMERSIVE_DARK_MODE,
        &enableDark,
        sizeof(enableDark));
}

/**
 * combinePath – allocate and combine two path segments.
 *
 * @param[out] outPath    Receives malloc’d wide string (must be freed by caller).
 * @param      maxChars   Maximum wchar count including terminating NUL.
 * @param      segment1   First path segment.
 * @param      segment2   Second path segment.
 * @return     S_OK on success, E_OUTOFMEMORY if allocation fails, E_FAIL if combine fails.
 */
static HRESULT combinePath(PWSTR *outPath, size_t maxChars, PCWSTR segment1, PCWSTR segment2) {
    *outPath = malloc(maxChars * sizeof **outPath);
    if (!*outPath) {
        return E_OUTOFMEMORY;
    }

    if (!PathCombineW(*outPath, segment1, segment2)) {
        free(*outPath);
        return E_FAIL;
    }

    return S_OK;
}


/* -------------------------------------------------------------------------- */
/* Dynamic array for menu items                                               */
/* -------------------------------------------------------------------------- */

/**
 * MenuEntry – one “Send To” item.
 *
 * @member path  Heap-alloc’d absolute path (owner).
 * @member icon  32-bit ARGB bitmap for the menu (may be NULL).
 */
typedef struct {
    PWSTR   path;
    HBITMAP icon;
} MenuEntry;

/**
 * MenuVector – simple grow-only array.
 *
 * @member items     Pointer to contiguous buffer (realloc’d).
 * @member count     Elements currently stored.
 * @member capacity  Allocated slots in @items.
 */
typedef struct {
    MenuEntry *items;
    UINT       count;
    UINT       capacity;
} MenuVector;

/**
 * vectorEnsureCapacity – make room for at least @need elements.
 *
 * @param vec   Target vector.
 * @param need  Desired minimum capacity.
 * @return      true on success, false on OOM.
 */
static bool vectorEnsureCapacity(MenuVector *vec, UINT need)
{
    if (need <= vec->capacity) {
        return true;
    }

    MenuEntry *tmp = realloc(vec->items, need * sizeof *tmp);
    if (!tmp) {
        return false;
    }

    // Zero the new tail so later clean-up is safe.
    ZeroMemory(tmp + vec->capacity, (need - vec->capacity) * sizeof *tmp);

    vec->items    = tmp;
    vec->capacity = need;
    return true;
}

/**
 * vectorPush – append a new entry (takes ownership of resources).
 *
 * @param vec   Vector to modify.
 * @param path  Heap path (caller must not reuse).
 * @param icon  Bitmap handle (may be NULL, ownership transferred).
 * @return      true on success, false on OOM — if false the caller still
 *              owns @path/@icon and must free/delete them.
 */
static bool vectorPush(MenuVector *vec, PWSTR path, HBITMAP icon)
{
    // If capacity growth fails, we do *not* consume the resources.
    if (!vectorEnsureCapacity(vec, vec->count + 1)) {
        return false;
    }

    // MSVC C compiler lacks compound literals → assign field-by-field.
    vec->items[vec->count].path = path;
    vec->items[vec->count].icon = icon;
    vec->count++;
    return true;
}

/**
 * vectorDestroy – free all stored paths/bitmaps and reset vector to zero.
 *
 * @param vec  Vector to wipe.
 */
static void vectorDestroy(MenuVector *vec)
{
    for (UINT i = 0; i < vec->count; ++i) {
        free(vec->items[i].path);
        if (vec->items[i].icon) DeleteObject(vec->items[i].icon);
    }
    free(vec->items);
    ZeroMemory(vec, sizeof *vec);
}


/* -------------------------------------------------------------------------- */
/* System small-icon cache                                                    */
/* -------------------------------------------------------------------------- */

// Global handle to the shared small icon list (AddRef'd)
static HIMAGELIST g_smallImageList = NULL;

/**
 * ensureSmallImageList – initialize the global small-icon image list.
 *
 * Tries SHGetImageList(SHIL_SMALL). On failure, falls back to
 * SHGetFileInfoW on the Windows directory and AddRefs it.
 */
static void ensureSmallImageList(void)
{
    if (g_smallImageList) {
        return;  // already initialized
    }

    // Try COM-based retrieval first
    HRESULT result = SHGetImageList(SHIL_SMALL, &IID_IImageList, (void**)&g_smallImageList);
    if (result == S_OK) {
        return;  // success, we own g_smallImageList
    }

    // Fallback: use Windows directory for reliable shell image list
    WCHAR sysDir[MAX_PATH];
    if (GetWindowsDirectoryW(sysDir, ARRAYSIZE(sysDir))) {
        SHFILEINFOW fileInfo = { 0 };
        HIMAGELIST fallbackList = (HIMAGELIST)SHGetFileInfoW(
            sysDir,
            FILE_ATTRIBUTE_DIRECTORY,
            &fileInfo,
            sizeof fileInfo,
            SHGFI_SYSICONINDEX | SHGFI_SMALLICON
        );
        if (fallbackList) {
            g_smallImageList = fallbackList;  // only assign if valid
        }
    }

    // AddRef so our copy stays valid after shell unloads
    if (g_smallImageList) {
        // image list is owned by shell. AddRef to keep it alive
        IUnknown *pUnk = (IUnknown*)g_smallImageList;
        if (pUnk) {
            pUnk->lpVtbl->AddRef(pUnk);
        }
    }
}

/**
 * cleanupSmallImageList – release the global small-icon image list.
 */
static void cleanupSmallImageList(void)
{
    if (!g_smallImageList) {
        return;
    }

    // Detach global reference before releasing
    IUnknown *pUnk = (IUnknown*)g_smallImageList;
    g_smallImageList = NULL;
    pUnk->lpVtbl->Release(pUnk);
}

/**
 * createDIBSection32 – allocate a top-down 32-bit DIB of given size.
 *
 * @param width   Desired bitmap width in pixels.
 * @param height  Desired bitmap height in pixels.
 * @return        New HBITMAP or NULL on failure.
 */
static HBITMAP createDIBSection32(int width, int height)
{
    // Describe a 32-bit top-down DIB
    BITMAPINFO bmi;
    ZeroMemory(&bmi, sizeof bmi);
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = width;
    bmi.bmiHeader.biHeight      = -height;  // negative => top-down orientation
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    return CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, NULL, NULL, 0);
}

/**
 * makeIconBitmap – create a 32-bit ARGB bitmap from the shared small icon list.
 *
 * @param iconIndex  Zero-based index into the small icon list.
 * @return           HBITMAP (32-bit ARGB) owned by caller, or NULL on failure.
 */
static HBITMAP makeIconBitmap(int iconIndex)
{
    ensureSmallImageList();
    if (!g_smallImageList) {
        return NULL;  // no image list available
    }

    int iconWidth, iconHeight;
    if (!ImageList_GetIconSize(g_smallImageList, &iconWidth, &iconHeight)) {
        return NULL;  // failed to retrieve dimensions
    }

    // Allocate a matching DIB
    HBITMAP bitmap = createDIBSection32(iconWidth, iconHeight);
    if (!bitmap) {
        return NULL;
    }

    // Draw the icon onto the DIB using the same DC
    HDC drawDC = CreateCompatibleDC(NULL);
    if (drawDC) {
        HGDIOBJ oldObj = SelectObject(drawDC, bitmap);
        ImageList_Draw(g_smallImageList, iconIndex, drawDC, 0, 0, ILD_TRANSPARENT);
        SelectObject(drawDC, oldObj);
        DeleteDC(drawDC);
    }

    return bitmap;
}

/**
 * dibFromIcon – convert an HICON to a 32-bit ARGB HBITMAP.
 *
 * @param iconHandle  Source HICON (ownership transferred; this function destroys it).
 * @return            HBITMAP or NULL on failure.
 */
static HBITMAP dibFromIcon(HICON iconHandle)
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

    // allocate DIB section; pointer to its bits is discarded here
    HBITMAP dibBitmap = createDIBSection32(bmpMetrics.bmWidth, bmpMetrics.bmHeight);
    if (dibBitmap) {
        // create temporary DC
        HDC drawDC = CreateCompatibleDC(NULL);
        if (drawDC) {
            HGDIOBJ oldObj = SelectObject(drawDC, dibBitmap);
            DrawIconEx(drawDC, 0, 0, iconHandle, bmpMetrics.bmWidth, bmpMetrics.bmHeight, 0, NULL, DI_NORMAL);
            SelectObject(drawDC, oldObj);
            DeleteDC(drawDC);
        }
    }

    // Cleanup original icon and bitmaps
    DeleteObject(iconInfo.hbmColor);
    DeleteObject(iconInfo.hbmMask);
    DestroyIcon(iconHandle);

    return dibBitmap;
}

/**
 * iconForPath – return a 32-bit ARGB bitmap for the given path.
 *
 * For directories: returns shell small icon.
 * For files: returns file icon (adds .lnk overlay if necessary).
 * Falls back to system list if SHGetFileInfo fails.
 */
static HBITMAP iconForPath(PCWSTR filePath)
{
    SHFILEINFOW info;
    UINT flags;

    // Directories: shell icon
    if (PathIsDirectoryW(filePath)) {
        flags = SHGFI_ICON | SHGFI_SMALLICON;
        if (SHGetFileInfoW(filePath, 0, &info, sizeof info, SHGFI_ICON | SHGFI_SMALLICON)) {
            HBITMAP result = dibFromIcon(info.hIcon);
            DestroyIcon(info.hIcon);
            if (result) {
                return result;
            }
        }

        // default for folders
        return NULL;
    }

    // Files: get icon, possibly with link overlay
    flags = SHGFI_ICON | SHGFI_SMALLICON;
    PCWSTR extension = PathFindExtensionW(filePath);
    if (extension && _wcsicmp(extension, L".lnk") == 0) {
        flags |= SHGFI_LINKOVERLAY;
    }
    
    // primary file icon (with overlay)
    if (SHGetFileInfoW(filePath, FILE_ATTRIBUTE_NORMAL, &info, sizeof info, flags)) {
        HBITMAP result = dibFromIcon(info.hIcon);
        DestroyIcon(info.hIcon);
        if (result) {
            return result;
        }
    }

    // fallback to system image list by index
    flags = SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON;
    if (SHGetFileInfoW(filePath, FILE_ATTRIBUTE_NORMAL, &info, sizeof info, flags)) {
        return makeIconBitmap(info.iIcon);
    }
    
    return NULL;
}


/* -------------------------------------------------------------------------- */
/* Menu population                                                            */
/* -------------------------------------------------------------------------- */

/**
 * skipEntry – filter out "." / ".." and hidden or system files.
 *
 * @param  fd  WIN32_FIND_DATA of current entry.
 * @return     TRUE if the entry must be ignored.
 */
static inline BOOL skipEntry(const WIN32_FIND_DATAW *findData)
{
    return (findData->dwFileAttributes &
            (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM)) ||
           wcscmp(findData->cFileName, L".")  == 0 ||
           wcscmp(findData->cFileName, L"..") == 0;
}

/**
 * addFileItem – insert a leaf item with icon into @menu and @items.
 */
static void addFileItem(
    HMENU       parentMenu,
    PCWSTR      fileName,
    HBITMAP     bitmap,
    UINT        commandId,
    MenuVector *vec,
    PWSTR       pathCopy
) {
    // vectorPush may fail; then we must clean up our resources
    if (!vectorPush(vec, pathCopy, bitmap)) {
        free(pathCopy);
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

    MENUITEMINFOW mi = { sizeof mi, MIIM_ID | MIIM_STRING | MIIM_BITMAP };
    mi.wID        = commandId;
    mi.dwTypeData = caption;
    mi.hbmpItem   = bitmap;
    InsertMenuItemW(parentMenu, commandId, FALSE, &mi);
}

/**
 * addDirectoryItem – create a submenu (recursion handled by caller).
 */
static void addDirectoryItem(
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
    MENUINFO menuInfo = { 0 };
    menuInfo.cbSize          = sizeof(menuInfo);
    menuInfo.fMask           = MIM_HELPID;          // we’re setting dwContextHelpID
    menuInfo.dwContextHelpID = helpId;              // ID used to look up real path

    SetMenuInfo(subMenu, &menuInfo);
}

/**
 * enumerateFolder – recursively enumerate a directory and add entries to a menu.
 *
 * @param menu        HMENU to which items and submenus will be added.
 * @param directory   Wide‐string path of the folder to enumerate.
 * @param nextCmdId   Pointer to the next command ID; incremented for each item.
 * @param depth       Current recursion depth; stops at MAX_DEPTH.
 * @param items       Pointer to a vector where (path, bitmap) pairs are stored.
 * @return            S_OK on success, or an HRESULT error code on failure.
 */
static HRESULT enumerateFolder(
    HMENU       menu,
    PCWSTR      directory,
    UINT       *nextCmdId,
    UINT        depth,
    MenuVector *items
) {
    // Stop if we've reached maximum allowed depth
    if (depth >= MAX_DEPTH) {
        return S_OK;
    }

    // Skip if the path isn't a directory
    if (!PathIsDirectoryW(directory)) {
        return S_OK;
    }

    // Build the search pattern "directory\\*"
    PWSTR pattern = NULL;
    HRESULT hr = combinePath(&pattern, MAX_LOCAL_PATH, directory, L"*");
    if (FAILED(hr)) {
        return hr;
    }

    // Begin file enumeration
    WIN32_FIND_DATAW findData;
    HANDLE hFind = FindFirstFileW(pattern, &findData);
    free(pattern);
    if (hFind == INVALID_HANDLE_VALUE)
        return HRESULT_FROM_WIN32(GetLastError());

    do {
        // Skip "." , "..", hidden, or system entries
        if (skipEntry(&findData))
            continue;

        // Build full child path
        PWSTR childPath = NULL;
        hr = combinePath(&childPath, MAX_LOCAL_PATH, directory, findData.cFileName);
        if (FAILED(hr)) {
            continue;
        }

        // Retrieve icon bitmap for this entry
        HBITMAP bmp = iconForPath(childPath);

        if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            // For subdirectories, create a new submenu
            HMENU subMenu = CreatePopupMenu();
           if (!subMenu) {
                // avoid inserting into NULL menu; clean up and skip
                free(childPath);
                if (bmp) {
                    DeleteObject(bmp);
                }
                continue;
            }

            // Insert directory item with icon and context-help ID
            addDirectoryItem(menu, findData.cFileName, bmp, subMenu, *nextCmdId);
            
            // Store in vector for later invocation
            vectorPush(items, childPath, bmp);
            (*nextCmdId)++;
            
            // Recurse into the subdirectory
            enumerateFolder(subMenu, childPath, nextCmdId, depth + 1, items);
        } else {
            // For files, insert a regular file item
            addFileItem(menu, findData.cFileName, bmp, (*nextCmdId)++,
                        items, childPath);
        }

    } while (FindNextFileW(hFind, &findData));

    // Clean up enumeration handle
    FindClose(hFind);
    return S_OK;
}


/* -------------------------------------------------------------------------- */
/* IDataObject builder & drop helpers                                         */
/* -------------------------------------------------------------------------- */

/**
 * PathToPIDL
 *
 * Convert a filesystem path into a COM PIDL (item identifier list).
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
    HRESULT hr = desktopShellFolder->lpVtbl->ParseDisplayName(
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
 * GetShellInterfaceForPIDLs
 *
 * Given an array of absolute PIDLs, bind to the requested COM interface
 * (IDataObject, IDropTarget, etc.) on their parent shell folder.
 *
 * @param hwndOwner     HWND used for binding context (may be NULL).
 * @param pidlArray     Array of absolute PIDLs.
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
            (LPCITEMIDLIST*)&childIDs[pidlIndex]
        );
        if (FAILED(hr)) {
            SAFE_RELEASE(parentFolder);
            free(childIDs);
            return hr;
        }

        // Release intermediate shell folder except for the last one
        if (pidlIndex < pidlCount - 1) {
            parentFolder->lpVtbl->Release(parentFolder);
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
    parentFolder->lpVtbl->Release(parentFolder);
    free(childIDs);

    return hr;
}

/**
 * GetShellInterfaceForPaths
 *
 * Convert an array of file or folder paths into PIDLs and
 * retrieve the requested COM interface from those PIDLs.
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
    HRESULT hr = GetShellInterfaceForPIDLs(
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
 * ExecuteDropOperation
 *
 * Perform COM drag-enter, drop, and leave on the given drop target.
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

    POINTL pt     = { 0 };
    DWORD  effect = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;

    // drag into the target
    HRESULT hrEnter = dropTarget->lpVtbl->DragEnter(
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
 * ExecuteDragDrop
 *
 * Perform a COM drag-and-drop of the files passed in argv[1…argc-1]
 * onto the destination specified by entry->path.
 *
 * @param owner   HWND of the hidden owner window for COM calls.
 * @param entry   Pointer to the MenuEntry containing the target .path.
 * @param argc    Argument count (program name + file paths).
 * @param argv    Array of PWSTR; argv[1…] are source file paths.
 */
static void ExecuteDragDrop(HWND owner, const MenuEntry *entry, int argc, PWSTR *argv)
{
    HRESULT      hr;
    IDataObject *pDataObj    = NULL;
    IDropTarget *pDropTarget = NULL;

    // Build IDataObject from the array of source file paths
    hr = GetShellInterfaceForPaths(
        owner,
        (PCWSTR*)(argv + 1),     // skip argv[0]
        argc - 1,                // number of files
        &IID_IDataObject,
        (void**)&pDataObj
    );
    if (FAILED(hr) || !pDataObj) {
        return;
    }

    // Retrieve the IDropTarget for the destination folder/link
    hr = GetShellInterfaceForPaths(
        owner,
        (PCWSTR*)&entry->path,   // address of the single target path
        1,                       // one drop target
        &IID_IDropTarget,
        (void**)&pDropTarget
    );

    if (SUCCEEDED(hr) && pDropTarget) {
        // Perform the actual drag-and-drop
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
        OleUninitialize();
        return FALSE;
    }

    // complete drag and drop setup
    HRESULT hr = SHGetDesktopFolder(&desktopShellFolder);
    if (FAILED(hr)) {
        OutputDebugStringW(L"[SendTo+] SHGetDesktopFolder failed\n");
        return FALSE;
    }

    // add darkmode support
    OptInDarkPopupMenus();

    return TRUE;
}

/**
 * ResolveSendToDirectory – build full path to "<exe folder>\\sendto" with \\?\ prefix.
 *
 * @return malloc’d wide string on success (must be freed), or NULL on failure.
 */
static PWSTR ResolveSendToDirectory(void)
{
    WCHAR exeFolder[MAX_PATH];
    // get path to our own executable
    GetModuleFileNameW(NULL, exeFolder, ARRAYSIZE(exeFolder));
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
        MessageBoxW(
            NULL,
            L"Cannot find 'sendto' folder next to the executable.",
            L"SendTo+ Error",
            MB_ICONERROR | MB_OK
        );
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

    // recursively fill menu and items vector
    UINT initialCmdId = 1;                     // start command IDs at 1
    HRESULT hr = enumerateFolder(
        *outPopup,
        sendToDir,
        &initialCmdId,                         // pass address of a real variable
        0,
        outItems
    );

    if (FAILED(hr)) {
        OutputDebugStringW(L"[SendTo+] enumerateFolder failed\n");
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
    WNDCLASSEXW wc = {
        .cbSize        = sizeof(WNDCLASSEXW),
        .lpfnWndProc   = DefWindowProcW,
        .hInstance     = hInstance,
        .lpszClassName = CLASS_NAME
    };
    if (!RegisterClassExW(&wc)) {
        OutputDebugStringW(L"[SendTo+] RegisterClassExW failed\n");
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
    if (hwnd) {
        ApplyDarkThemeIfNeeded(hwnd);
        ShowWindow(hwnd, SW_HIDE);
        SetForegroundWindow(hwnd);
    } else {
        OutputDebugStringW(L"[SendTo+] CreateWindowExW failed\n");
    }

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
    POINT cursor;
    // get cursor position for menu location
    GetCursorPos(&cursor);

    UINT cmd = TrackPopupMenuEx(
        popup,
        TPM_RETURNCMD | TPM_NONOTIFY | TPM_LEFTALIGN | TPM_LEFTBUTTON,
        cursor.x, cursor.y,
        owner,
        NULL
    );

    // exit menu-mode to restore normal input
    PostMessage(owner, WM_NULL, 0, 0);

    return cmd;
}

/**
 * CleanupApplication – free all resources and uninitialize subsystems.
 *
 * @param ownerWnd   hidden owner window to destroy.
 * @param popupMenu  HMENU to destroy.
 * @param sendToDir  path string to free.
 * @param items      MenuVector of menu entries to free.
 */
static void CleanupApplication(
    HWND ownerWnd,
    HMENU popupMenu,
    PWSTR sendToDir,
    MenuVector *items
) {
    // free menu item data
    vectorDestroy(items);
    DestroyMenu(popupMenu);

    // free directory path
    free(sendToDir);

    // destroy hidden window
    if (ownerWnd) {
        DestroyWindow(ownerWnd);
        UnregisterClassW(L"SendToOwnerWindow", GetModuleHandleW(NULL));
    }

    // clean COM interface
    SAFE_RELEASE(desktopShellFolder);

    // release global image list and OLE
    cleanupSmallImageList();
    OleUninitialize();
}

/**
 * wWinMain – main entry point for SendTo+ clone.
 *
 * @param hInstance     application instance.
 * @param hPrevInstance unused.
 * @param lpCmdLine     unused.
 * @param nCmdShow      unused.
 * @return exit code (0 success, non-zero on error).
 */
int WINAPI wWinMain(
    HINSTANCE hInstance,
    HINSTANCE hPrevInstance,
    PWSTR     lpCmdLine,
    int       nCmdShow
) {
    /* suppress unused-parameter warnings */
    (void)hPrevInstance;
    (void)lpCmdLine;
    (void)nCmdShow;

    WCHAR buf[256];
    StringCchPrintfW(buf, 256, L"[SendTo+] __argc=%d\n", __argc);
    OutputDebugStringW(buf);
    for (int i = 0; i < __argc; i++) {
        StringCchPrintfW(buf, 256, L"[SendTo+] __wargv[%d]=%s\n", i, __wargv[i]);
        OutputDebugStringW(buf);
    }

    int argc;
    PWSTR *argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (!argv) return 1;

    if (!InitializeApplication()) {
        return 1;
    }

    // build and verify sendto directory
    PWSTR sendToDir = ResolveSendToDirectory();
    if (!ValidateSendToDirectory(sendToDir)) {
        return 1;
    }

    // build popup menu and items
    HMENU popupMenu;
    MenuVector menuItems;
    if (!BuildSendToMenu(sendToDir, &popupMenu, &menuItems)) {
        return 1;
    }

    // create hidden owner window
    HWND owner = CreateHiddenOwnerWindow(hInstance);

    // display menu and handle selection
    UINT choice = DisplaySendToMenu(popupMenu, owner);
    if (choice) {
        MenuEntry *item = &menuItems.items[choice - 1];
        if (argc == 1) {
            // no args: open folder/link
            OutputDebugStringW(L"[SendTo+] no args: open folder/link\n");
            ShellExecuteW(NULL, NULL, item->path, NULL, NULL, SW_SHOWNORMAL);
        } else {
            // with args: perform drag-and-drop
            OutputDebugStringW(L"[SendTo+] with args: perform drag-and-drop\n");
            ExecuteDragDrop(owner, item, argc, argv);
        }
    }

    // clean up all resources
    CleanupApplication(owner, popupMenu, sendToDir, &menuItems);
    LocalFree(argv);

    return 0;
}
