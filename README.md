# VbeComplete · Python edition

*English · [简体中文](README.zh-CN.md)*

The VBA editor (VBE) only pops up a member list after `object.` — it never completes the variable,
constant, procedure or form names **you** declared. VbeComplete fills that gap: type the first few
characters of a name (Chinese identifiers included) and a candidate list appears at the caret.
It works with **Microsoft Office** (Excel / Word / PowerPoint / Access / Visio) and **WPS**, and
detects the host from the foreground window.

![Type "cus", the list filters live, Tab completes](assets/demo.gif)

## Features

- **Live candidates** for identifiers, declarations, procedures, modules, forms, controls, keywords
  and built-in functions — filtered as you type, with a type icon on every row.
- **Smart acceptance**: `Tab` completes; `Sub` gets a trailing space, `Function` gets `()` with the
  caret inside; keywords get a space.
- **Enter auto-indent**: block openers indent one level deeper and add the matching `End If` /
  `Next` / `Loop`; `Else` / `Case` lines are re-aligned first.
- **Shift+Enter** starts a new line below without splitting the current one.
- **Backspace, two jobs**: removes a whitespace-only line in one press, or deletes an empty pair
  `()` / `""` from the middle.
- **Auto-pairing** `(` `"` — full-width `（）` / `“”` when the IME is in Chinese punctuation mode;
  typing the closing char just steps over it; a selection gets wrapped.
- **Line commands**: `Ctrl+/` comment toggle · `Ctrl+D` duplicate · `Ctrl+K` delete ·
  `Ctrl+Shift+Enter` insert above · `Ctrl+Shift+↑/↓` jump to the last declaration / closing line.
- **`Ctrl+Q` snippet library** (SQLite, two columns, fuzzy title search, insert at caret).
- **Self-healing**: a heartbeat plus `watchdog.py` restart the tool when it hangs or is killed.
- **Adaptive polling**: 50 ms while typing, 150 ms when idle, 250 ms outside the VBE.

## Pros

- Completes **your own** project symbols, not just a fixed keyword list — Chinese names work.
- One set of features across the whole Office suite and WPS.
- Survives host busy states: a hung or killed process is restarted automatically, no reboot needed.
- Comfortable cost: about **2% of one core** and **46 MB** while idle inside the VBE.
- Every shortcut obeys one rule — **if the action fails, the original key is replayed** — so keys
  never get "eaten" silently.

## Cons

- **Windows only**, and needs **Python 3.8+** installed.
- Requires *Trust access to the VBA project object model* (off by default) — without it the tool
  cannot read your code at all.
- The idle **~2% of one core is Tk's event loop floor**; it does not drop when you leave the VBE.
- No parameter-info popup, and host enum constants (`vb*` / `xl*` / `mso*`) are deliberately skipped.
- The global hook is mounted only while the VBE is focused, so completion is unavailable in other
  editors, and there is a short focus-switch delay after returning to the VBE.
- Editing the `.py` sources requires a restart (`stop.bat` → `run.bat`); there is no hot reload.
