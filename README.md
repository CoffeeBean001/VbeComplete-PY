# VbeComplete

[![License: MIT](assets/license-mit.svg)](LICENSE)

**IntelliJ IDEA–style autocomplete for the VBA editor.**

![VbeComplete in action: typing "cus" filters the list live, and Tab accepts the highlighted candidate](assets/demo.gif)

*English · [简体中文](README.zh-CN.md)*

The VBE only completes `object.member` — never *your own* variable or procedure names. This fills
that gap: type a few letters, a live-filtered list appears next to the caret, `Tab` accepts it.
Each entry carries a small **type icon** (method / property / variable / type / keyword), the same
idea as the VBE's own member list, so you can tell them apart at a glance.
Works in the VBE shared by **Microsoft Office** (Excel, Word, PowerPoint, Access, Visio) and
**WPS Office** — the host is detected automatically.

## Before you start

1. **Windows** plus any host above.
2. **Python 3.8+** (3.10 – 3.12 recommended); tick *Add python.exe to PATH* when installing.
3. **"Trust access to the VBA project object model" must be ticked**, then restart the host —
   *File → Options → Trust Center → Trust Center Settings → Macro Settings*. This is separate from
   the macro security level, and without it the tool can read no code. (WPS Office has the same
   switch under its own options.)
4. Changed the code? **`stop.bat` → overwrite → `run.bat`** — Python loads the `.py` files at startup.
5. **It restarts itself (v128).** When another automation keeps the host busy, the main thread can
   block inside a COM call — the process is still there, but the UI and the completion popup stop
   responding. After 30 seconds of that the tool exits and restarts itself; if the process is killed
   outright, the supervisor `watchdog.py` brings it back within seconds. **No more rebooting.**
   What it did is written to `watchdog.log` in the project folder; to turn it off completely set
   `AUTO_RESTART = 0` in `settings.ini`.
   ⚠️ Stop it with `stop.bat` (that stops the supervisor too). Killing `pythonw.exe` in Task Manager
   only gets it restarted a few seconds later — that is by design, not a stuck kill.

## Usage

```bat
git clone https://github.com/CoffeeBean001/VbeComplete-V587.git
cd VbeComplete-V587
run.bat
```

The first launch installs the dependencies (needs internet), then it runs silently in the
background. `stop.bat` stops it · `install_autostart.bat` / `uninstall_autostart.bat` add / remove
autostart on login · `run_tests.bat` runs the offline regression suite.

| Key | Action |
| --- | --- |
| letter / digit / `_` | Open the list and filter it live |
| `↑` / `↓` | Move the selection (wraps) |
| `Shift+↑` / `Shift+↓` | Close the list and move the caret one line |
| `Tab` | **Accept** the highlighted candidate (the only accept key) |
| click a candidate | Accept with the mouse |
| `Esc` / `←` / `→` / click elsewhere | Dismiss the list |
| `Ctrl+Space` | Force the list to open |
| `Enter` *(at end of line)* | New line with smart indentation, plus an automatic block closer |
| `Shift+Enter` | Start a new line below, leaving the current one intact |
| `Backspace` | Delete a whitespace-only line, or both halves of an empty pair |
| `Ctrl+Shift+↑` | Put the caret at the **end of the last `Dim` line** of the procedure (no line is inserted) |
| `Ctrl+Shift+↓` | Jump to the blank line above `End Sub` / `End Function` (inserting one if needed) and write it at the procedure's indent level |
| `Ctrl+Q` | Open the **snippet window** at the caret: title list on the left (fuzzy search box), code on the right; add / edit / delete, and insert the snippet at the caret |
| `Ctrl+/` | **Toggle line comment** (IDEA-style): on every line of the selection, or the caret's line when there is none — removes the comment when all of them already have one, otherwise adds it |
| `Ctrl+Shift+Enter` | Insert a new line **above** the current one (same indent; one level deeper when the current line is a block **closer** such as `End If` / `Next`) |
| `Ctrl+D` | **Duplicate the current line** below (the selected lines, when there is a selection); the caret follows to the same column |
| `Ctrl+K` | **Delete the whole line** (the selected lines, when there is a selection); the caret moves to the end of the line above |

On accept a `Sub` gets a trailing space, a `Function` gets `()` with the caret inside **and the
VBE's own parameter info pops up** (so you can see the signature while filling in the arguments),
and a keyword gets a space. Typing `(` / `"` auto-closes the pair (caret in the middle, never inside
a comment); with a Chinese IME in full-width / Chinese-punctuation mode the Chinese pair `（）` / `“”`
is inserted instead of being converted to ASCII. A `(` only auto-closes when nothing but whitespace
is to its right — **except inside a call's argument list or a procedure's parameter list, where it
always closes** (`foo(bar` + `(` -> `foo(bar())`) — and typing `"` or `(` over a selection wraps it
(`abc` -> `"abc"` / `(abc)`). **Snippet library (`Ctrl+Q`)**: a SQLite database with **two columns**
(title / code) **in the project folder** (`snippets.db`, created on first use). The window opens at
your caret; the title list is **fuzzy searchable** — a contiguous substring ranks first, but initials
work too (`gc` finds `GetCode`), and space-separated terms must all match. Inserting writes the
snippet **verbatim** at the caret (your own indentation is kept), leaves the caret after it, closes
the window and gives focus back to the VBE.

**Toggle comment (`Ctrl+/`)**: press it over a selection to comment those lines out, press it again
to bring them back verbatim; blank lines are skipped and the indentation is never swallowed. Same
rule as IDEA: it *removes* only when **every non-blank line is already commented**, otherwise it
adds. Removal strips one `' ` (quote + space) and falls back to a bare `'`, so hand-written
`'notes` uncomment correctly too. One caveat: N lines means **N undo steps** (the VBE has no
command groups), so `Ctrl+Z` needs to be pressed N times to get all the way back.

**Three line-wise editing keys (`Ctrl+Shift+Enter` / `Ctrl+D` / `Ctrl+K`)** all work the same way:
**the selected lines when there is a selection, otherwise the caret's line** — the IDEA / VS Code
convention.

- `Ctrl+Shift+Enter` inserts a line **above**. The indent matches the current line — except when that
  line is a block **closer** (`End If` / `Next i` / `End Sub`), where the new line belongs *inside*
  the block and is indented one level deeper.
- `Ctrl+D` duplicates the line **below**, with the caret following to the copy at the same column.
- `Ctrl+K` deletes the whole line, moving the caret to the **end of the line above** the deleted block.
  Deleting the entire module is allowed.

⚠️ On the key choice: `Shift+D` / `Shift+K` were the original request, but those keys are how you
**type an uppercase `D` / `K`** — VBE's auto-capitalisation only applies to keywords, not to string
contents like `"Debug"`. Hence `Ctrl+D` / `Ctrl+K`, neither of which the VBE code pane binds.

**Where the suggestions come from: `wordbank.db`, which you can edit.** VBA's keywords, built-in
functions, built-in types and common constants live in `wordbank.db` in the **project folder** (SQLite,
seeded from the list compiled into the program on first start). It is read into memory once at startup;
every keystroke afterwards only touches memory — so **editing the bank costs you no speed**, and reading
it takes about 1 ms.

Three ways to change it (restart to take effect):

- open it with any SQLite tool (DB Browser for SQLite and friends) and edit the `words` table;
- no tool needed: `py -3 wordbank.py` for a summary, `list` to see everything, `add Name` to add one,
  `disable Name` to keep but stop suggesting it, `enable` to bring it back, `remove` to delete;
- broke something? `py -3 wordbank.py reset --yes` rebuilds it (a `wordbank.db.bak-<time>` is written
  automatically first).

One row per word, primary key `(name, kind)` — names like `Date` / `String` that are **both a function
and a type** get one row each. Columns: `kind` (`keyword` / `function` / `constant` / `type`); `enabled`
(0 = keep but do not suggest); `is_builtin` (1 = shipped with the language; write 0 for your own words so
they are not subjected to the "common vocabulary" fuzzy-match tightening); `wants_space` (keywords only:
NULL = follow the default rule, 1 = always append a space, 0 = never); `note` (free text, not read by
the program).

⚠️ **Functions, variables, module and control names defined in your code are not in there** — those are
read from your live VBE project on every pass, and that is the main path. `wordbank.db` only covers the
language-provided set plus whatever you add.
⚠️ For troubleshooting: `NO_WORDBANK = 1` in `settings.ini` ignores the bank entirely and falls back to
the seed list compiled into the program.

**Turning features off: `settings.ini`** — created in the **project folder** on first start, with
**every line commented out** (so it changes nothing by itself). Delete the `;` in front of a line,
edit the value, save, restart. **Environment variables still win** over the file. To see what is
actually in effect and where each value came from, run `py -3 settings.py` in the project folder.

**What is suggested**: your variables, constants and procedures, plus modules, forms and control
names, keywords and built-in functions / types (the language-provided set now comes from
`wordbank.db`, so you can trim or extend it) — enums (`vb*` / `xl*` / `mso*`) are excluded.
While you *declare* a new name (`Dim …`, `Sub …`) no existing symbol is offered, exactly like IDEA —
but the keywords that belong at the caret are: after `Private` / `Public` you get `Sub` / `Function` /
…, inside a parameter list `ByVal` / `ByRef` / `Optional` / `ParamArray` / `As`; after `As` / `New` it
still offers types.

**No popup?** Check note 3 first — that is it nine times out of ten. Then look for `pythonw.exe` in
Task Manager; if a list won't close, run `stop.bat` then `run.bat`.

## More

[Feature details & IDEA comparison](docs/features.md) ·
[Troubleshooting 排错](docs/troubleshooting.zh-CN.md) ·
[Chinese docs](docs/features.zh-CN.md) · [MIT License](LICENSE)
