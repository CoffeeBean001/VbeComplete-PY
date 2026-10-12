# VbeComplete

[![License: MIT](assets/license-mit.svg)](LICENSE)

**IntelliJ IDEA–style autocomplete for the VBA editor.**

![VbeComplete in action: typing "cus" filters the list live, and Tab accepts the highlighted candidate](assets/demo.gif)

*Read this in other languages: **English** · [简体中文](README.zh-CN.md)*

The VBE only ever completes `object.member` — it never suggests *your own* variable or procedure
names. VbeComplete fills that gap: start typing a name and a live-filtered list pops up next to the
caret; press `Tab` to accept, `↑` / `↓` to choose, `Esc` to dismiss.

It works with the VBE shared by every **Microsoft Office** host (Excel, Word, PowerPoint, Access,
Visio) and by **WPS Office** (Writer, Spreadsheets, Presentation). The current host is detected
automatically — no configuration.

## Before you start

1. **Windows** plus any of the hosts listed above.
2. **Python 3.8 or newer** (3.10 – 3.12 recommended). Use an official build from
   [python.org](https://www.python.org/downloads/) and tick *Add python.exe to PATH*.
   Python 3.13 works too — `pip` just pulls a newer `pywin32` (≥ 308).
3. **"Trust access to the VBA project object model" must be ticked** in the host you use, or the
   tool cannot read the project's code at all. *File → Options → Trust Center → Trust Center
   Settings → Macro Settings →* tick it, then **restart the application**. This is a separate
   switch from the macro security level; changing the security level alone does nothing. WPS
   exposes the same switch under its own options.
4. The first launch needs an internet connection to install `pywin32` and `pynput`; later launches
   start instantly.
5. **Restart the process after changing code** — Python reads the `.py` files at startup, so
   overwriting them without a restart keeps the old code running:
   **`stop.bat` → overwrite → `run.bat`**.

## Install and run

```bat
git clone https://github.com/CoffeeBean001/VbeComplete-V587.git
cd VbeComplete-V587
run.bat
```

`run.bat` installs the dependencies on first launch and then runs the tool silently in the
background. To watch the output instead, run `python main.py` in a console.

Then open the VBE (`Alt+F11`), click inside a code pane and type the first letters of one of your
own variable names — the list appears next to the caret.

- `stop.bat` — stop the background process
- `run_tests.bat` — run the regression suite
- `install_autostart.bat` / `uninstall_autostart.bat` — add / remove autostart on login (a shortcut
  to `pythonw main.py` in your Startup folder; the tool is focus-driven, so it attaches its hooks
  only while a VBE code pane is active)

## Usage

### Keyboard

| Key | Action |
| --- | --- |
| letter / digit / `_` | Open the candidate list and filter it live |
| `↑` / `↓` | Move the selection (wraps); the list scrolls when it overflows |
| `Shift+↑` / `Shift+↓` | Close the list and move the caret one line up / down |
| `Tab` | **Accept** the highlighted candidate (the only accept key) |
| click a candidate | Accept with the mouse |
| `Esc` / `←` / `→` / click elsewhere | Dismiss the list (the arrow keys still move the caret) |
| `Ctrl+Space` | Force the list to open |
| `Enter` *(at end of line)* | New line with smart indentation, plus an automatic block closer |
| `Shift+Enter` | Start a new line below without splitting the current one |
| `Backspace` | Delete a whitespace-only line in one stroke; delete both halves of an empty pair |

### What happens on accept

- A `Sub` gets a trailing space, a `Function` gets `()` with the caret placed inside, a `Property`
  gets nothing, and a keyword gets its own trailing space.
- Inside a `Function`, its **own** name is the return-value variable, so completing it there
  (`MyFunc = …`) inserts nothing.

### What gets suggested

- Variables, constants and procedures (`Sub` / `Function` / `Property`) from the project, plus
  modules, forms and control names.
- Without `Option Explicit`, variables that are used but never declared are suggested too.
- VBA keywords and built-in functions / data types (≈260 names) are included by default;
  **enums (`vb*` / `xl*` / `mso*`) are deliberately excluded**.
- Visibility follows VBA's own scoping rules — only names reachable from the caret are shown.

### No popup?

- **Check note 3 above** (the trust setting) first — nine times out of ten that is it, and the
  application must be restarted after ticking it.
- Look for `pythonw.exe` in Task Manager; if it is missing, run `run.bat` again.
- A list that won't close, or accepts everything, usually means stale instances: run `stop.bat`,
  then `run.bat`.
- Nothing after `.` inside a `With` block: that position is intentionally left to the VBE's own
  member list, which needs the object's type to be resolvable — compile the project
  (*Debug → Compile*) so it can be.
- **While you are declaring a new name** (`Dim …`, `Sub …`) the list stays closed, exactly like
  IDEA: you are inventing a name, so there is nothing to complete. The type position after
  `As` / `New` is unaffected.

## Running the tests

```bat
run_tests.bat
```

Runs offline against a real 14-module sample project in `tests/testdata/` — **no Office or COM
required**. A few cases read the live type library and assert Excel-specific names; when the host
detected is not Excel they are skipped automatically.

## More

- [Feature details and a point-by-point comparison with IntelliJ IDEA](docs/features.md)
- [Full Chinese documentation](README.zh-CN.md)
- [MIT License](LICENSE)
