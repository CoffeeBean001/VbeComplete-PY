# VbeComplete — feature reference

*[Back to README](../README.md) · [简体中文](features.zh-CN.md)*

> Moved out of the original `README.md` unchanged, so the front page stays short.

## Highlights

**Completion**

- **Suggestions as you type.** Type a letter, digit or `_` and the list appears; keep typing and it filters live.
- **Fuzzy and camelCase matching.** You don't have to start at the first letter — type `ds` to get
  `dataSheet` (**D**ata**S**heet). Matched characters are highlighted so you can see *why* a name matched.
- **Quality-first ranking.** Results are ordered exact › prefix › word-initial acronym › contiguous
  block › scattered, then by hit position, span and name length. The closest match is always first,
  and a plain prefix behaves exactly like a plain prefix search.
- **Scope-aware ranking** *(IDEA's "local scope first")*. A local variable or parameter outranks a
  module-level name, which outranks a `Public` name from another module.
- **Recently-used ranking** *(IDEA's "recently used")*. Names you just accepted are promoted on ties.
- **Context-aware follow-up.** Accepting a `Sub` inserts a trailing space (`DoWork `), a `Function`
  inserts `()` with the caret placed inside (`GetName()`), a `Property` inserts nothing, and a
  keyword gets its own trailing space (`public `). Inside a `Function`, its **own** name is the
  return-value variable, so completing it there — `MyFunc = …` — inserts nothing
  (no `MyFunc() = …`).
- **Quiet while you declare — except for the keywords that belong there.** At a declaration —
  `Dim …`, `Const …`, `Sub …`, `Function …`, `Property …`, `Type …` — the list never offers
  *existing symbols* (variable names, procedure names, module names): you are *inventing* a name, so
  there is nothing to complete, exactly like IDEA. Since v111 it does offer the **language keywords
  that belong at the caret**: after `Private` / `Public` you get `Sub` / `Function` / `Property` /
  `Type` / … , and inside a parameter list you get `ByVal` / `ByRef` / `Optional` / `ParamArray` /
  `As`. Positions that really are just a new name (`Dim x…`, `Private Sub F…`) stay silent. The type
  position after `As` / `New` is unaffected (there you *are* referring to an existing type, so type
  names are still offered).
- **Parameter info on accept.** Accepting a `Function` inserts `()` with the caret inside, and
  VbeComplete then presses `Ctrl+Shift+I` for you so the VBE's *own* parameter-info window appears —
  the native signature bubble that highlights the argument you are filling in and follows your
  commas. It only fires when *we* added the parentheses (typing `(` yourself already pops it). Use
  `VBECOMPLETE_NO_PARAM_INFO=1` to turn it off.
- **Unicode identifiers.** Names such as `Dim 用户名 As String` are fully supported.
- **Non-intrusive.** Suggestions appear only inside a VBE *code pane* — never in the Properties
  window, the Project explorer or the form designer.

**Editing helpers**

- Auto-paired brackets and quotes (`(`, `[`, `"`, …), with the caret placed between the two halves.
  Auto-pairing asks the IME for its conversion mode first: with a Chinese IME in **full-width /
  Chinese-punctuation** mode it inserts the Chinese pair (`（` `）`, `“` `”`) instead of silently
  converting what you typed to ASCII. `VBECOMPLETE_NO_IME_PUNCT=1` forces the half-width pair.
- **v112:** an opening bracket auto-closes **only when nothing but whitespace is to its right** —
  with code to the right, just the bracket itself is inserted (same as modern editors). And typing
  a quote over a **selection** wraps the selection instead of replacing it (`abc` -> `"abc"`);
  single-line selections only (a VBA string cannot span lines). Brackets over a selection were
  still handed to VBE in v112 — v114 changes that.
- **v113:** the rule above has one exception — **inside a call's argument list or a procedure's
  parameter list the bracket is still auto-closed** (`foo(bar(` -> `foo(bar())`,
  `Sub Foo(ByVal arr(` -> `Sub Foo(ByVal arr())`), because the code to the right of the caret
  belongs to an **outer** pair. Only the innermost unclosed `(` that sits right next to an
  identifier counts; grouping parens (`x = (a + b`), a leading `(` and already-closed pairs do not.
- **v114:** typing `(` over a **selection** now wraps it, exactly like the quote path
  (`abc` -> `(abc)`, full-width -> `（abc）`). One whitelist does both jobs:
  `VbeBackend._WRAP_OPEN` = `"` `“` `(` `（`; closing characters (`)` `）`) are deliberately
  absent — they have no "wrap" meaning, so a selection is still replaced by VBE. Still
  single-line only (a bracket pair cannot span lines in VBA without `_` continuations), and the
  wrapping itself is still `_wrap_selection` — one implementation shared with quotes, pinned by
  an AST guard (test 98.3).
- `Backspace` deletes both halves of an empty pair, and removes a whitespace-only line in one stroke.
- `Enter` at end of line indents to the right depth, aligns `Else` / `ElseIf` / `Case`, and closes a
  block you just opened (`End Sub`, `Next i`, …).
- `Shift+Enter` starts a new line below without splitting the current one — the same indentation
  logic as `Enter`, at any column.
- **`Ctrl+↑` / `Ctrl+↓` jump inside a procedure.** `Ctrl+↑` puts the caret at the **end of the
  last `Dim` line** of the current procedure — no new line is inserted. `Ctrl+↓` moves the caret to
  the line *above* `End Sub` / `End Function` / `End Property`; if that line is not blank yet, a
  blank line is inserted first, and it is written at the procedure's indent level (copied from the
  last non-blank line above), so the caret drops straight where you can keep typing; a line that
  already carries indentation is left untouched. VBE's own `Ctrl+↑` / `Ctrl+↓` do nothing in a code
  pane, so the key is handed back to the system untouched when the move cannot be made (no procedure,
  no `Dim`, or a read-only module for `Ctrl+↓`). Use `VBECOMPLETE_NO_CTRL_NAV=1` to disable, or
  `VBECOMPLETE_NO_CTRL_NAV_TAB=1` to keep the `Ctrl+↓` jump but never add the indent.
- **v115 — snippet library (`Ctrl+Q`).** A window opens at the caret: title list on the left,
  code box on the right. Clicking a title shows its code; **Insert at caret** writes it into the
  editor, closes the window and gives focus back to the VBE. The store is SQLite with **two
  columns** — `title TEXT PRIMARY KEY, code TEXT NOT NULL` — at
  **in the project folder** as `snippets.db` (created on first use; `VBECOMPLETE_SNIPPET_DB` overrides).
  ★v116: it used to live in `%APPDATA%\VbeComplete\` — the user asked for it to sit next to the code so
  it is easy to find and travels with the folder (it is in `.gitignore`, it is your data, not source).
  ★v117 layout: the right edges of the **search box, the title box and the *Insert at caret* button
  all line up with the right edge of the code box** (0 px). The code box has a vertical scrollbar to
  its right, so the title box must not span that column, and the full-width rows leave `8 + scrollbar
  width` on the right (`_scrollbar_width()` measures it, 17 is the fallback). §99.6 measures real
  widget geometry in a real Tk window, not the config.
  Titles are **fuzzy searchable**: exact title > prefix > contiguous substring > subsequence
  (`gc` finds `GetCode`), space-separated terms are ANDed. Chinese is matched by **character**
  (`数据` finds `数据处理`), not by pinyin initials — that would need a whole pinyin table.
  Insertion is line-splicing: `head + first line`, the middle lines, `last line + tail`, so it
  works mid-line too, and a selection is replaced (single- or multi-line). The snippet is written
  **verbatim** — your indentation is kept — and the caret lands right after it. Turn it off with
  `VBECOMPLETE_NO_SNIPPETS=1`.

- **v118 — the popup keeps up with your typing.** Measured with `tests/probe_perf.py` on a
  24-module / 628-line project: one full `get_identifiers()` refresh cost **224 ms**, and
  `poll_editor` threw the pool away **every 0.4 s** while you typed (any text change with
  `time.time() - _cache_time > ID_REFRESH_MIN_SEC`). So one keystroke every 400 ms had to sit
  through 224 ms of COM on the main thread — that is the "the popup got slow" feeling.
  cProfile says it is not our regexes: `PyITypeComp.Bind` 65 ms + `Invoke` 25 ms +
  `GetTypeAttr` 21 ms + `GetNames` 19 ms ≈ 160 ms of win32com **late binding**, because every
  element enumerated from `proj.VBComponents` is a fresh wrapper that has to Bind its type info
  before the first `comp.Name` / `comp.Type` / `comp.CodeModule` read.
  Fixes (three, independent): **per-module parse cache** (replay the five sets for a module whose
  line count did not change; the module under the caret is *never* cached — its results are bound
  to the caret position; `_MOD_CACHE_TTL` 30 s and "invalidate on module switch" cover renames
  that leave the line count alone), **component-list cache** (`_comp_list_cache`, reused wrappers,
  invalidated by component count or 15 s), and **`CodeModule` reuse** (`_cm_cache`).
  Full refresh: **224 ms → 27 ms**. Scoring was never the problem: 455 candidates take 0.6 ms,
  and even 6200 take 17 ms — "too many suggestions" was not the cause.
  Refresh throttling is now two-tier (`parser.line_may_declare`): 0.4 s on declaration lines
  (`Dim` / `Sub` / `As` / parameter lists — the only lines that can add a name), 2 s elsewhere.
  `VBECOMPLETE_NO_MOD_CACHE=1` restores the v117 behaviour for A/B.

- **v119 — the popup finally matches your editor font, and "recently used" survives a restart.**
  Two gaps found while measuring v118; both are silent (nothing ever errored, so nobody reported them).
  *(a)* `ui._apply_font()` computed the font size as a **float** (`pt + FONT_DELTA` → `11.0` / `8.75`),
  but Tk's `font configure -size` accepts integers only — it raises
  `TclError: expected integer but got "9.75"`, and two nested `except: pass` swallowed it. The candidate
  window therefore sat on `Consolas 10` forever: changing the VBE editor font to 14 pt or to another
  typeface had no effect. Fixed by a pure `_clamp_font_size()` that rounds (`8.75 → 9`; `int()` would
  give 8, one notch too small) and tolerates `None` / `NaN` / `±inf` / strings. One line of real change,
  but the failure mode was invisible.
  *(b)* The recently-used ranking table was a plain in-memory dict, and `stop.bat` kills the process
  with `Stop-Process -Force`, so **no exit hook ever runs** — every reboot started the ranking from zero.
  It is now persisted to `recent.json` in the project folder (same policy as `snippets.db`; override with
  `VBECOMPLETE_RECENT_FILE`), holding only *name → sequence*: a few KB. Three design points matter:
  the keystroke path only sets a dirty flag and `poll_editor` flushes **at most every 2 s** (a clean
  check costs one boolean; a write is a few KB and under 1 ms); constructing `Completer()` never reads
  the disk (the test suite builds thousands of them, and they must not depend on the user's real data);
  and writes are atomic (`tmp` + `os.replace`) with any corrupt file degrading to `{}`.
  `VBECOMPLETE_NO_RECENT_PERSIST=1` restores the v118 behaviour for A/B.

**Coexistence**

- After `object.`, VbeComplete steps aside and lets the VBE's own, type-aware member list take over.
- Global hooks are attached **only while a VBE code pane has focus**, so the tool is essentially
  invisible (and free) the rest of the time.
- **Closing other IDEs stays fast.** Reading the foreground window's title used to call
  `GetWindowText`, a *cross-thread synchronous* call: while another program is shutting down
  (PyCharm / VSCode tearing their windows down), that one line stalled our own main thread for
  1.5 s or more, which in turn stalled the loop that detaches the global hooks — so the hooks
  stayed on top of the dying process and its exit hung (a regression of the v27–v29 fix).
  The title is now read **only after** a cheap, 200 ms-capped "is this window still responding?"
  probe, and the loop itself is wrapped so that no exception can ever stop it. Measured on the
  same scenario: main-thread stall **1.484 s → 0.000 s**, hook detach **5 s+ → 0.26 s**.

## How it compares to IntelliJ IDEA

VbeComplete reproduces the *feel* of IDEA's completion; it does not try to reimplement an IDE's
semantic infrastructure.

**Implemented**

| IDEA | VbeComplete |
| --- | --- |
| Completion as you type, live filtering | ✅ |
| Fuzzy / camelCase matching with highlighted hits | ✅ |
| Quality-ranked results | ✅ |
| Local scope first | ✅ |
| Recently used promoted | ✅ |
| Auto-paired brackets and quotes | ✅ |
| Paired deletion on `Backspace` | ✅ |
| Smart indentation and automatic block closers | ✅ |
| Start a new line (`Shift+Enter`) | ✅ |
| Trailing space for keywords; `()` after functions | ✅ |
| No completion while declaring a new name | ✅ |
| Type-aware member completion | Delegated to the VBE's native member list |

**Deliberately not implemented**

Go-to-definition, Find Usages, rename refactoring, live syntax/error checking, postfix completion
and live templates all depend on a full, always-current semantic model of the project — type
inference, symbol tables, a dependency graph. The VBE exposes only the *plain text* of a
`CodeModule`, so rebuilding a compiler front-end on top of it would be costly and fragile; the
risk/benefit simply doesn't add up.

## Configuration

Behaviour is tuned through environment variables (set them before launching):

| Variable | Default | Effect |
| --- | --- | --- |
| `VBECOMPLETE_LOG` | off | `1` writes a diagnostic log to `VbeComplete.log` (`VBECOMPLETE_LOG_FILE` relocates it). |
| `VBECOMPLETE_AUTO_PAREN` | on | `0` disables inserting `()` after procedures and functions. |
| `VBECOMPLETE_AUTO_SPACE` | on | `0` disables the automatic trailing space after keywords. |
| `VBECOMPLETE_NO_PARAM_INFO` | off | `1` stops pressing `Ctrl+Shift+I` after auto-inserted `()`, i.e. no parameter-info popup. |
| `VBECOMPLETE_MEMBER_POPUP` | off | `1` restores VbeComplete's own list after `object.` (off by default). |
| `VBECOMPLETE_DECL_POPUP` | off | `1` restores suggesting inside declarations (off by default). |
| `VBECOMPLETE_NO_SCOPE_RANK` | off | `1` disables scope-aware ranking. |
| `VBECOMPLETE_NO_RECENT_RANK` | off | `1` disables recently-used ranking. |
| `VBECOMPLETE_RECENT_MAX` | 64 | How many recently-used names are remembered. |
| `VBECOMPLETE_NO_RECENT_PERSIST` | off | `1` keeps the recently-used table in memory only (`recent.json` is neither read nor written). |
| `VBECOMPLETE_RECENT_FILE` | `<project folder>\recent.json` | Where the recently-used table lives (tests and portable installs point it elsewhere). |
| `VBECOMPLETE_NO_YIELD` | off | `1` disables yielding to the VBE's native member list. |
| `VBECOMPLETE_NO_AUTOPAIR` | off | `1` disables auto-pairing. |
| `VBECOMPLETE_NO_IME_PUNCT` | off | `1` always pairs with half-width `()` / `""`, even when the IME is in Chinese-punctuation / full-width mode. |
| `VBECOMPLETE_NO_BS_DELLINE` / `VBECOMPLETE_NO_BS_DELPAIR` | off | `1` disables the two `Backspace` helpers. |
| `VBECOMPLETE_NO_ENTER_INDENT` / `VBECOMPLETE_NO_AUTOCLOSE` | off | `1` disables `Enter`-time indentation / automatic block closers. |
| `VBECOMPLETE_NO_CTRL_NAV` | off | `1` disables the `Ctrl+↑` / `Ctrl+↓` jumps; the keys stay with the VBE. |
| `VBECOMPLETE_NO_CTRL_NAV_TAB` | off | `1` keeps the `Ctrl+↓` jump but never adds the procedure-level indent to the landing line. |
| `VBECOMPLETE_NO_SNIPPETS` | off | `1` disables the `Ctrl+Q` snippet window; the key stays with the VBE. |
| `VBECOMPLETE_SNIPPET_DB` | `<project folder>\snippets.db` | Where the snippet database lives (portable installs / tests point it elsewhere). |
| `VBECOMPLETE_NO_MOD_CACHE` | off | `1` goes back to re-parsing every module on each refresh (v117 behaviour). Only useful for A/B timing; the per-module cache is what keeps a refresh at ~27 ms instead of ~224 ms. |

## Project structure

| File | Role |
| --- | --- |
| `main.py` | Entry point: global keyboard hook, event loop, wiring |
| `engine.py` | Completion engine: scoring, ranking, filtering, popup lifecycle |
| `vbe_bridge.py` | VBE COM bridge: read code, read caret, write back, caret placement |
| `parser.py` | VBA text parsing (statements, scope, comments and strings) |
| `vba_builtins.py` | VBA built-in vocabulary (functions, types, keywords, enums) |
| `ui.py` | The floating popup (tkinter) and its screen-avoidance logic |
| `snippet_ui.py` | The snippet window (`Ctrl+Q`): search, add / edit / delete, insert button |
| `snippets.py` | The snippet store (SQLite, two columns) and fuzzy title matching — no UI, no COM |
| `recent.py` | The recently-used table (`recent.json`): load, atomic save, corrupt-file tolerance — no UI, no COM (v119) |
| `log.py` | Optional diagnostic log (off by default, zero overhead) |
| `tests/` | Regression suite (`test_real_project.py`, 1062 checks) and the sample project it runs against. Three one-shot probes ship with it: `probe_ime_mode.py` prints the current IME conversion mode (how the `IME_CMODE_SYMBOL` constant below was verified), `probe_hook_gate.py` drives the real hook lifecycle against fake windows to measure how an unresponsive foreground window affects the main thread, and `probe_perf.py` times the popup path (`get_context`, a full re-parse, and the scoring curve). |

## Credits

Built by [CoffeeBean001](https://github.com/CoffeeBean001). Inspired by the completion experience
of IntelliJ IDEA.
