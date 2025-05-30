/*
 * sendto.c – Unicode-only reworked version of SendTo+ (MSVC)
 * Copyright (c) 2025 DSR! <xchwarze@gmail.com>
 * based on https://github.com/lifenjoiner/sendto-plus
 */

#include <sdkddkver.h>
#define WIN32_LEAN_AND_MEAN
#define _WIN32_WINNT _WIN32_WINNT_WIN7
#define _WIN32_IE _WIN32_IE_IE80

#include <windows.h>
#include <stdlib.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <commoncontrols.h> /* for IID_IImageList */
#include <shellapi.h>
#include <strsafe.h>
#include <stdbool.h>
#include <dwmapi.h>

#pragma comment(lib, "comctl32.lib")   // commctrl.h – InitCommonControlsEx, ImageList_*, etc.
#pragma comment(lib, "shell32.lib")    // shlobj.h, shobjidl.h – SHGetKnownFolderPath, IShellItem, etc.
#pragma comment(lib, "shlwapi.lib")    // shlwapi.h – PathIsDirectoryW, StrCmpLogicalW, etc.
#pragma comment(lib, "dwmapi.lib")     // dwmapi.h – DwmExtendFrameIntoClientArea, etc.
#pragma comment(lib, "ole32.lib")      // COM: CoCreateInstance, etc.

#define MAX_DEPTH 5
#define MAX_LOCAL_PATH 32767


/* -------------------------------------------------------------------------- */
/* Utility macros                                                             */
/* -------------------------------------------------------------------------- */

#define SAFE_RELEASE(p)     do { if (p) { (p)->lpVtbl->Release(p); (p) = NULL; } } while (0)
#define RETURN_IF_FAILED(h) do { HRESULT _hr = (h); if (FAILED(_hr)) return _hr; } while (0)
#define BOOL_IF_FAILED(h)   do { if (FAILED(h)) return FALSE; } while(0)
#define ERR_BOX(msg)        MessageBoxW(NULL, msg, L"SendTo+", MB_OK|MB_ICONERROR)


/* -------------------------------------------------------------------------- */
/* Helpers                                                                    */
/* -------------------------------------------------------------------------- */

static HMODULE uxThemeModule = NULL;
static LPSHELLFOLDER desktopShellFolder = NULL;

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
static BOOL AppUsesDarkTheme(void)
{
    typedef BOOL (WINAPI *ShouldAppsUseDarkMode_t)(void);

    // Ordinals are stable since 1809
    ShouldAppsUseDarkMode_t ShouldAppsUseDarkMode =
        (ShouldAppsUseDarkMode_t)GetProcAddress(uxThemeModule, MAKEINTRESOURCEA(132));

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
static void ApplyDarkThemeIfNeeded(HWND hwnd)
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
 * @param      segment1   First path segment.
 * @param      segment2   Second path segment.
 * @return     S_OK on success, E_OUTOFMEMORY if allocation fails, E_FAIL if combine fails.
 */
static HRESULT CombinePath(PWSTR *outPath, PCWSTR segment1, PCWSTR segment2)
{
    *outPath = malloc(MAX_LOCAL_PATH * sizeof **outPath);
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
static bool VectorEnsureCapacity(MenuVector *vec, UINT need)
{
    if (need <= vec->capacity) {
        return TRUE;
    }

    // Amortized growth: double current capacity (or start at 64), but at least 'need'
    UINT newCap = vec->capacity ? vec->capacity * 2 : 64;
    if (newCap < need) {
        newCap = need;
    }

    MenuEntry *tmp = realloc(vec->items, newCap * sizeof *tmp);
    if (!tmp) {
        return FALSE;
    }

    // Zero the new tail so later clean-up is safe.
    ZeroMemory(tmp + vec->capacity,
               (newCap - vec->capacity) * sizeof *tmp);

    vec->items    = tmp;
    vec->capacity = newCap;

    return TRUE;
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
static bool VectorPush(MenuVector *vec, PWSTR path, HBITMAP icon)
{
    // If capacity growth fails, we do *not* consume the resources.
    if (!VectorEnsureCapacity(vec, vec->count + 1)) {
        return FALSE;
    }

    // MSVC C compiler lacks compound literals → assign field-by-field.
    vec->items[vec->count].path = path;
    vec->items[vec->count].icon = icon;
    vec->count++;

    return TRUE;
}

/**
 * vectorDestroy – free all stored paths/bitmaps and reset vector to zero.
 *
 * @param vec  Vector to wipe.
 */
static void VectorDestroy(MenuVector *vec)
{
    for (UINT i = 0; i < vec->count; ++i) {
        free(vec->items[i].path);
        if (vec->items[i].icon) DeleteObject(vec->items[i].icon);
    }
    free(vec->items);
    ZeroMemory(vec, sizeof *vec);
}


/* -------------------------------------------------------------------------- */
/* Icons tools                                                                */
/* -------------------------------------------------------------------------- */

/**
 * createDIBSection32 – allocate a top-down 32-bit DIB of given size.
 *
 * @param width   Desired bitmap width in pixels.
 * @param height  Desired bitmap height in pixels.
 * @return        New HBITMAP or NULL on failure.
 */
static HBITMAP CreateDIBSection32(int width, int height)
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
 * dibFromIcon – convert an HICON to a 32-bit ARGB HBITMAP.
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

    // allocate DIB section; pointer to its bits is discarded here
    HBITMAP dibBitmap = CreateDIBSection32(bmpMetrics.bmWidth, bmpMetrics.bmHeight);
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
static HBITMAP IconForPath(PCWSTR filePath)
{
    SHFILEINFOW info;
    UINT flags;

    // Directories: shell icon
    if (PathIsDirectoryW(filePath)) {
        flags = SHGFI_ICON | SHGFI_SMALLICON;
        if (SHGetFileInfoW(filePath, 0, &info, sizeof(info), flags)) {
            HBITMAP result = DibFromIcon(info.hIcon);
            DestroyIcon(info.hIcon);
            return result;
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
    
    if (SHGetFileInfoW(filePath, FILE_ATTRIBUTE_NORMAL, &info, sizeof(info), flags)) {
        HBITMAP result = DibFromIcon(info.hIcon);
        DestroyIcon(info.hIcon);
        return result;
    }

    // fallback: system image list (includes non-existent/virtual items)
    flags = SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON;
    if (SHGetFileInfoW(filePath, FILE_ATTRIBUTE_NORMAL, &info, sizeof(info), flags)) {
        HBITMAP result = DibFromIcon(info.hIcon);
        DestroyIcon(info.hIcon);
        return result;
    }

    return NULL;
}


/* -------------------------------------------------------------------------- */
/* Menu population                                                            */
/* -------------------------------------------------------------------------- */

/**
 * skipEntry – filter out "." / ".." and hidden or system files.
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
 * @param pathCopy   Heap-allocated copy of the file path; freed on failure.
 * @return           void; on push failure, frees pathCopy and bitmap if set.
 */
static void AddFileItem(
    HMENU       parentMenu,
    PCWSTR      fileName,
    HBITMAP     bitmap,
    UINT        commandId,
    MenuVector  *vec,
    PWSTR       path
) {
    // vectorPush may fail; then we must clean up our resources
    if (!VectorPush(vec, path, bitmap)) {
        free(path);
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
    menuInfo.fMask           = MIM_HELPID | MIM_STYLE; // we’re setting dwContextHelpID
    menuInfo.dwContextHelpID = helpId;                 // ID used to look up real path
    menuInfo.dwStyle         = MNS_AUTODISMISS
                               | MNS_NOTIFYBYPOS;

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

    // Skip if the path isn't a directory
    if (!PathIsDirectoryW(directory)) {
        return S_OK;
    }

    // Build the search pattern "directory\\*"
    PWSTR pattern = NULL;
    HRESULT hr = CombinePath(&pattern, directory, L"*");
    if (FAILED(hr)) {
        return hr;
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

    do {
        // Skip "." , "..", hidden, or system entries
        if (SkipEntry(&findData))
            continue;

        // Build full child path
        PWSTR childPath = NULL;
        hr = CombinePath(&childPath, directory, findData.cFileName);
        if (FAILED(hr)) {
            continue;
        }

        // Retrieve icon bitmap for this entry
        HBITMAP bmp = IconForPath(childPath);

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
            AddDirectoryItem(menu, findData.cFileName,
                             bmp, subMenu, *nextCmdId);
            
            // Store in vector for later invocation
            VectorPush(items, childPath, bmp);
            (*nextCmdId)++;
            
            // Recurse into the subdirectory
            EnumerateFolder(subMenu, childPath, nextCmdId, depth + 1, items);
        } else {
            // For files, insert a regular file item
            AddFileItem(menu, findData.cFileName, bmp, (*nextCmdId)++,
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
    const HRESULT hr = SHGetDesktopFolder(&desktopShellFolder);
    if (FAILED(hr)) {
        OutputDebugStringW(L"[SendTo+] SHGetDesktopFolder failed\n");
        return FALSE;
    }

    // add dark mode support
    uxThemeModule = LoadLibraryW(L"uxtheme.dll");
    if (!uxThemeModule) {
        return FALSE;
    }

    OptInDarkPopupMenus();

    return TRUE;
}

/**
 * ParseCommandLine - Parses switches and returns a clean argv[].
 *
 * @param  rawArgc      Argument count from CommandLineToArgvW().
 * @param  rawArgv      Argument vector from CommandLineToArgvW().
 * @param  outDir       Receives a malloc’d wide string if “/D <dir>” was supplied.
 *                      Caller must free() it when done.  May be NULL.
 * @param  outArgc      Receives the new argument count.
 * @param  outArgv      Receives a malloc’d PWSTR[] of length outArgc:
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
    int     *outArgc,
    PWSTR   **outArgv
) {
    *outDir  = NULL;
    *outArgc = 1;                   // always keep exe @ index 0
    *outArgv = NULL;

    // allocate worst-case full array
    PWSTR *temp = malloc(rawArgc * sizeof *temp);
    if (!temp) {
        return FALSE;
    }

    // keep program name
    temp[0] = rawArgv[0];

    // scan switches & collect files
    for(int paramIndex = 1; paramIndex < rawArgc; ++paramIndex) {
        PWSTR param = rawArgv[paramIndex];

        // help?
        if (_wcsicmp(param, L"/?")==0 || _wcsicmp(param, L"-?")==0) {
            ERR_BOX(L"Error: /D requires a directory path.\n"
                    L"Usage: SendTo+ [/D <directory>] [<file1> <file2> ...]");
            return FALSE;
        }

        // override SendTo directory?
        if (_wcsicmp(param, L"/D")==0) {
            if (paramIndex + 1 < rawArgc) {
                *outDir = _wcsdup(rawArgv[++paramIndex]);
                if (!*outDir) {
                    return FALSE;
                }
            } else {
                ERR_BOX(L"Error: /D requires a directory path.\n"
                        L"Usage: SendTo+ [/D <directory>] [<file1> <file2> ...]");
                return FALSE;
            }

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
    return TRUE;
}

/**
 * ResolveSendToDirectory – build full path to "<exe folder>\\sendto" with \\?\ prefix.
 *
 * @return malloc’d wide string on success (must be freed), or NULL on failure.
 */
static PWSTR ResolveSendToDirectory(void)
{
	// get path to our own executable
    WCHAR exeFolder[MAX_PATH];
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

    // enable bitmap arrows (chevrons) AND auto-dismiss on click-outside
    /*MENUINFO menuInfo = { sizeof(menuInfo) };
    menuInfo.fMask   = MIM_STYLE;
    menuInfo.dwStyle = MNS_CHECKORBMP     // draw ► for submenus
                       | MNS_AUTODISMISS  // close when clicking outside
                       | MNS_NOTIFYBYPOS; // get position notifications
    SetMenuInfo(*outPopup, &menuInfo);*/

    // recursively fill menu and items vector
    UINT initialCmdId = 1;                     // start command IDs at 1
    const HRESULT hr = EnumerateFolder(
        *outPopup,
        sendToDir,
        &initialCmdId,                         // pass address of a real variable
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
        // theme support setup
        //SetWindowTheme(hwnd, L"ExplorerMenu", NULL);
        ApplyDarkThemeIfNeeded(hwnd);

        // menu logic
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
    // get cursor position for menu location
    POINT cursor;
    GetCursorPos(&cursor);

    const UINT cmd = TrackPopupMenuEx(
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
 * RunSendTo – perform full SendTo+ workflow.
 *
 * @hInstance    application instance.
 * @argc         argument count.
 * @argv         argument vector (wide).
 * @return       exit code (0 success, non-zero on error).
 */
static int RunSendTo(HINSTANCE hInstance, int argc, PWSTR *argv)
{
    // build and verify sendto directory
    int cleanArgc = 0;
    PWSTR *cleanArgv = NULL;
    PWSTR sendToDir = NULL;
    if (!ParseCommandLine(argc, argv, &sendToDir, &cleanArgc, &cleanArgv)) {
        return EXIT_FAILURE;
    }

    if (!sendToDir) {
        sendToDir = ResolveSendToDirectory();
    }

    if (!ValidateSendToDirectory(sendToDir)) {
        return EXIT_FAILURE;
    }

    // build popup menu and items
    HMENU popupMenu;
    MenuVector menuItems;
    if (!BuildSendToMenu(sendToDir, &popupMenu, &menuItems)) {
        return EXIT_FAILURE;
    }

    // create hidden owner window
    HWND owner = CreateHiddenOwnerWindow(hInstance);

    // display menu and handle selection
    UINT choice = DisplaySendToMenu(popupMenu, owner);
    if (choice) {
        MenuEntry *item = &menuItems.items[choice - 1];
        if (cleanArgc > 1) {
            // with args: perform drag-and-drop
            OutputDebugStringW(L"[SendTo+] with args: perform drag-and-drop\n");
            ExecuteDragDrop(owner, item, cleanArgc, cleanArgv);
        } else {
            // no args: open folder/link
            OutputDebugStringW(L"[SendTo+] no args: open folder/link\n");
            ShellExecuteW(NULL, NULL, item->path, NULL, NULL, SW_SHOWNORMAL);
        }
    }

    // release COM desktop folder
    SAFE_RELEASE(desktopShellFolder);
    OleUninitialize();

    return EXIT_SUCCESS;
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

    if (!InitializeApplication()) {
        return EXIT_FAILURE;
    }

    // better params support
    int rawArgc;
    PWSTR *rawArgv = CommandLineToArgvW(GetCommandLineW(), &rawArgc);
    if (!rawArgv) {
        return EXIT_FAILURE;
    }

    return RunSendTo(hInstance, rawArgc, rawArgv);
}
