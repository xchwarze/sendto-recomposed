# SendTo Recomposed

A modern, Unicode-only re-implementation of the classic **SendTo+** helper for Windows.  
Compared with the original project, this version is faster, 64-bit clean and long-path-safe, shows *custom folder icons* in the popup, and carefully frees every resource.  
Licensed under **GPL v3**.

## Key Ingredients

* **Unicode-only** no ANSI/TCHAR branches – predictable builds 
* **Darkmode ready** Native dark-theme support
* **Custom folder icons** honours `desktop.ini` for nicer menu visuals 
* **Icon cache** a single `SHGetImageList` call – speedy painting 
* **Secure recursion** hidden/system items skipped, depth capped to 4 
* **Robust drag-and-drop** real `IDataObject`, always calls `DragLeave` 
* **Clean shutdown** no GDI, COM or image-list leaks 

## Building

MSVC 

```cmd
cl /EHsc /O2 /MD /DUNICODE /D_UNICODE ^
   sendto.c ^
   ole32.lib shell32.lib shlwapi.lib comctl32.lib user32.lib gdi32.lib uuid.lib
```

The output `sendto.exe` is fully 64-bit.

## Usage

1. Copy the executable anywhere on disk.
2. Run it – a **Send To** style popup appears under the cursor with every item (and sub-folder) from your personal *SendTo* directory.
3. Either click an entry to launch it, or drag files onto the menu and drop them on a target to perform the same action Explorer would.

## License

This code is released under the **GNU General Public License v3.0**.
See [LICENSE](LICENSE.txt) for full text.

## Credits

* **Original concept & code** © 2009 lifenjoiner
* **Complete refactor & modernisation** © 2025 DSR! - xchwarze@gmail.com
