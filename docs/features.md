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
| `VBECOMPLETE_NO_YIELD` | off | `1` disables yielding to the VBE's native member list. |
| `VBECOMPLETE_NO_AUTOPAIR` | off | `1` disables auto-pairing. |
| `VBECOMPLETE_NO_IME_PUNCT` | off | `1` always pairs with half-width `()` / `""`, even when the IME is in Chinese-punctuation / full-width mode. |
| `VBECOMPLETE_NO_BS_DELLINE` / `VBECOMPLETE_NO_BS_DELPAIR` | off | `1` disables the two `Backspace` helpers. |
| `VBECOMPLETE_NO_ENTER_INDENT` / `VBECOMPLETE_NO_AUTOCLOSE` | off | `1` disables `Enter`-time indentation / automatic block closers. |
| `VBECOMPLETE_NO_CTRL_NAV` | off | `1` disables the `Ctrl+↑` / `Ctrl+↓` jumps; the keys stay with the VBE. |
| `VBECOMPLETE_NO_CTRL_NAV_TAB` | off | `1` keeps the `Ctrl+↓` jump but never adds the procedure-level indent to the landing line. |

## Project structure

| File | Role |
| --- | --- |
| `main.py` | Entry point: global keyboard hook, event loop, wiring |
| `engine.py` | Completion engine: scoring, ranking, filtering, popup lifecycle |
| `vbe_bridge.py` | VBE COM bridge: read code, read caret, write back, caret placement |
| `parser.py` | VBA text parsing (statements, scope, comments and strings) |
| `vba_builtins.py` | VBA built-in vocabulary (functions, types, keywords, enums) |
| `ui.py` | The floating popup (tkinter) and its screen-avoidance logic |
| `log.py` | Optional diagnostic log (off by default, zero overhead) |
| `tests/` | Regression suite (`test_real_project.py`) and the sample project it runs against. Two one-shot probes ship with it: `probe_ime_mode.py` prints the current IME conversion mode (how the `IME_CMODE_SYMBOL` constant below was verified), and `probe_hook_gate.py` drives the real hook lifecycle against fake windows to measure how an unresponsive foreground window affects the main thread. |

## Credits

Built by [CoffeeBean001](https://github.com/CoffeeBean001). Inspired by the completion experience
of IntelliJ IDEA.
