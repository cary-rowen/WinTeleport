# WinTeleport

This NVDA add-on allows you to move the focused window between Windows virtual desktops using keyboard shortcuts.

## Features
- Move the focused window to the left or right virtual desktop.
- Move the focused window to a specific desktop by number.
- Optionally switch to the target desktop after moving the window.
- Report the current virtual desktop's name and number.

## Default Keyboard Shortcuts
- `NVDA+Windows+D`: Report current virtual desktop information (name, number, and total count).
- `NVDA+Control+Windows+LeftArrow`: Move the focused window to the left desktop and switch to it.
- `NVDA+Control+Windows+RightArrow`: Move the focused window to the right desktop and switch to it.
- `NVDA+Control+Windows+1` to `9`: Move the focused window to desktops 1 through 9.

*Note: If you want to move a window without switching desktops (staying on the current one), you can assign your own shortcuts in the NVDA Input Gestures dialog under the "WinTeleport" category.*

## Numeric Shortcut Limitations
- Moving directly to a specific desktop is currently limited to desktops 1 through 9. You can only assign custom key combinations that include the keys 1 through 9.

## Known Issues and Limitations
- **Compatibility**: A small number of applications with non-standard window frames (e.g., traditional/legacy versions of QQ) may not support moving between desktops via this add-on.
- **Focus Issues**: If you move the last window out of a virtual desktop without switching to the target desktop, the system focus may land on an "unknown" object.

## References

This project is inspired by and incorporates logic from the following open-source projects:
- [pyvda](https://github.com/mirober/pyvda): A Python library for Windows 10/11 virtual desktop management.
- [MoveToDesktop](https://github.com/Eun/MoveToDesktop): A utility for moving windows between virtual desktops.

## Future To-Dos (If Needed)
- **Window Pinning**: Support for "pinning" windows or applications so they remain visible on all virtual desktops.
- **Settings Panel**: Support for personalizing specific options via a configuration interface.
