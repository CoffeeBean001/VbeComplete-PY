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
  Writing a line back also makes the VBE **reformat** it — it adds a space after each `,` and `:`
  and around `=` (`MsgBox "abc",vbExclamation:Exit Sub` reads back as
  `MsgBox "abc", vbExclamation: Exit Sub`), and it trims trailing spaces. So the caret index is
  computed on the **line as the VBE actually stored it**, never on the text we meant to write: the
  word's end is re-located inside the stored line, and only then does the trailing space get
  injected there (v126e — before that fix the index was two cells short on such a line and the caret
  landed *inside* `Sub`, with the injected space splitting the word).
- **Quiet while you declare — except for the keywords that belong there.** At a declaration —
  `Dim …`, `Const …`, `Sub …`, `Function …`, `Property …`, `Type …` — the list never offers
  *existing symbols* (variable names, procedure names, module names): you are *inventing* a name, so
  there is nothing to complete, exactly like IDEA. Since v111 it does offer the **language keywords
  that belong at the caret**: after `Private` / `Public` you get `Sub` / `Function` / `Property` /
  `Type` / … , and inside a parameter list you get `ByVal` / `ByRef` / `Optional` / `ParamArray` /
  `As`. Positions that really are just a new name (`Dim x…`, `Private Sub F…`) stay silent. The type
  position after `As` / `New` is unaffected (there you *are* referring to an existing type, so type
  names are still offered).
- **Parameter info on accept (off by default).** Accepting a `Function` inserts `()` with the caret
  inside. Older versions then tried to pop the VBE's own parameter-info window; on real machines
  that is impossible from the outside — both `Ctrl+Shift+I` and the VBE menu command
  (`CommandBars`, Id 2532) leave `PopupTipWndClass` invisible and **beep** (measured with the
  WASAPI peak meter: 0.346 vs 0.000 for a plain line write). So nothing is triggered by default
  and accepting a completion is silent. The VBE's own *Tools ▸ Options ▸ Editor ▸ Auto List
  Parameter Info* still works when you type the first argument. Opt in with
  `VBECOMPLETE_PARAM_INFO=1` (it will beep); `VBECOMPLETE_NO_PARAM_INFO=1` always wins.
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
- **`Ctrl+Shift+↑` / `Ctrl+Shift+↓` jump inside a procedure.** `Ctrl+Shift+↑` puts the caret at the **end of the
  last `Dim` line** of the current procedure — no new line is inserted. `Ctrl+Shift+↓` moves the caret to
  the line *above* `End Sub` / `End Function` / `End Property`; if that line is not blank yet, a
  blank line is inserted first, and it is written at the procedure's indent level (copied from the
  last non-blank line above), so the caret drops straight where you can keep typing; a line that
  already carries indentation is left untouched. VBE's own `Ctrl+Shift+↑` / `Ctrl+Shift+↓` do nothing in a code
  pane, so the key is handed back to the system untouched when the move cannot be made (no procedure,
  no `Dim`, or a read-only module for `Ctrl+Shift+↓`). Use `VBECOMPLETE_NO_CTRL_NAV=1` to disable, or
  `VBECOMPLETE_NO_CTRL_NAV_TAB=1` to keep the `Ctrl+Shift+↓` jump but never add the indent.
  (v126c moved these from `Ctrl+↑` / `Ctrl+↓` to `Ctrl+Shift+↑` / `Ctrl+Shift+↓`; only the modifier
  gate changed — the plan, the landing spot and both backend methods are the same code. The fallback
  `send_vk` re-sends the full combination, because dropping Shift would replay VBE's own
  *previous procedure* jump instead of the key the user pressed.)
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

- **v120 — toggle line comment (`Ctrl+/`).** Pressed over a selection it comments out every selected
  line; pressed again it brings them back verbatim. With no selection it works on the caret's line.
  Same rule as IDEA: it **removes** only when **every non-blank line is already commented**, otherwise
  it adds. Blank lines are skipped entirely (no stray `'`), the comment character lands **after the
  indentation** (`    x = 1` -> `    ' x = 1`) — at column 1 it would comment out the indentation too,
  and the VBE re-indents on the way back — and removal strips one `' ` (quote + space) first and falls
  back to a bare `'`, so a hand-written `'notes` uncomments correctly. Only `'` is recognized, never
  `Rem`. Two VBA facts to know: one edit per line means **the undo stack is N steps** for N lines (COM
  has no command groups), and `ReplaceLine` **trims trailing spaces**. The selection is restored
  afterwards so the key really is a toggle; a selection that *ends* at column 1 is read as excluding
  its last line (that is what dragging down past the newline produces), matching IDEA. Turn it off with
  `VBECOMPLETE_NO_COMMENT=1`.

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

- **v120 — one place for the 40+ switches, and an IDEA-style `Ctrl+/`.**
  *(a) `settings.ini`.* Until v119 every switch was an environment variable, and nobody remembers
  forty names or wants to edit system environment variables to turn one feature off. On first start
  the tool now writes `settings.ini` into the **project folder** with **every switch listed and every
  line commented out** — so creating the file changes nothing by itself (a test pins this down). Delete
  the leading `;`, edit, save, restart. Precedence is **environment > `settings.ini` > code default**,
  and the `VBECOMPLETE_` prefix is optional inside the file (`NO_SNIPPETS` and
  `VBECOMPLETE_NO_SNIPPETS` both work, case-insensitive). `py -3 settings.py` prints every value *and
  where it came from*; `--write` regenerates the file. Four points were easy to get wrong: the ini is
  read **once at import** (these are module-level constants and the keystroke path must never touch the
  disk); the rule is "the first layer that *parses* wins", so a garbage env value like `=maybe` is
  skipped and the ini is still consulted instead of falling straight back to the code default; an empty
  string counts as unset (`set X=` leaves an empty value on Windows); and `configparser` needs
  `interpolation=None` (else `VBE_POPUP_CLASS=A%B` blows up the whole file) plus `strict=False` **as a
  constructor argument** (`cp.strict = False` is a no-op) and `utf-8-sig` (Notepad's BOM otherwise turns
  the first key into `\ufeffAUTO_SPACE`). `settings.py` sits at the bottom of the dependency graph and
  imports nothing from the project. An old bug fell out of this: the log path was hard-coded to
  `D:\codes\VbeComplete-V587\VbeComplete.log`, so it silently stopped working on any other machine.
  *(b) `Ctrl+/` toggles line comments* over the selection, or over the caret's line when there is none,
  with IDEA's rule: it **removes** only when **every non-blank line is already commented**, otherwise it
  adds. Blank lines are skipped, the `'` lands **after the indentation** (putting it at column 1 comments
  out the indentation itself, and the VBE re-indents on the way back), and removal strips one `' ` and
  falls back to a bare `'` so hand-written `'notes` uncomments correctly. Only `'` is recognized — never
  `Rem`, where a wrong guess would eat a character of real code. Two VBA-specific hazards: `ReplaceLine`
  **trims trailing spaces**, and one edit per line means the undo stack is **N steps** (COM has no
  command groups) — both are documented rather than papered over. The selection is restored afterwards
  (`SetSelection(sl, 1, el, very large column)`, letting the VBE clamp it) so that pressing the key twice
  really is a toggle; a selection that *ends* at column 1 is treated as excluding its last line, matching
  IDEA. `VBECOMPLETE_NO_COMMENT=1` turns it off.

- **v122 — three line-wise editing keys (`Ctrl+Shift+Enter` / `Ctrl+D` / `Ctrl+K`).**
  All three follow the same rule: **the selected lines when there is a selection, otherwise the
  caret's line** (the IDEA / VS Code convention). They *do* write code, so they reuse the skeleton of
  the v94 backspace-delete-line work: the hook thread only decides, the main thread writes, and a
  failed write hands the key back to the system.

  `Ctrl+Shift+Enter` inserts a line **above**, at the current line's indent — except when that line is
  a block **closer** (`End If` / `Next i` / `End Sub`), where the new line belongs *inside* the block
  and goes one level deeper. `Ctrl+D` duplicates the line **below**, the caret following to the copy
  at the same column. `Ctrl+K` deletes the whole line, the caret moving to the **end of the line above**
  the deleted block (so you can keep typing right where that line left off; deleting the first line puts
  the caret at column 1 of the new first line).

  On the key choice: `Shift+D` / `Shift+K` were the original request, but those keys are how you type
  an uppercase `D` / `K` — and the VBE's auto-capitalisation only covers keywords, so `"Debug"` inside
  a string literal would become impossible to type. Taking an *input* capability away to gain a
  convenience is a bad trade, hence `Ctrl+D` / `Ctrl+K` (IDEA's Duplicate Line key; neither is bound in
  the VBE code pane).

  Choices worth knowing: every action writes **exactly once** (`InsertLines` / `DeleteLines`), because
  reading while writing fragments the undo stack and can leave half-written code; `new_line_above`
  only asks "is this line a closer?" and guesses nothing further (`branch_align` is designed for
  "start a new line *after* this one", not for inserting above); a new pure predicate
  **`is_block_closer`** had to be added, because `block_closer("    Next i")` returns `None` — it maps
  *header → owed closer*, so using it here silently inverts the answer; the caret only lands where
  there is no column ambiguity (column 1 when deleting, the original column when duplicating, since a
  whole-line copy puts the same character at the same display column); and deleting the entire module
  is allowed (`CountOfLines` becomes 0, so the caret placement sits in its own `try` and never turns a
  successful delete into a `False`). The hook gate adds `pending_keys > 0` to the usual checks: these
  three write text, so a half-composed IME word would leave the line text stale.
  `VBECOMPLETE_NO_INSERT_ABOVE=1` / `VBECOMPLETE_NO_DUP_LINE=1` / `VBECOMPLETE_NO_DEL_LINE=1`.

- **v123 — the vocabulary lives in `wordbank.db`, which you can edit.**
  VBA's keywords, built-in functions, built-in types and common constants used to be Python tuples in
  `vba_builtins.py`. They now live in a SQLite database in the **project folder**, seeded from those
  tuples on the first start. It is read **once at startup** and then lives in memory as plain tuples and
  frozensets, so per-keystroke cost is exactly what it was (measured: `load()` 1.0 ms,
  `load_active()` 1.3 ms, `import vba_builtins` 0.87 ms — unchanged, because the module still does no
  I/O at import time).

  This is the two-source split the user asked for, and the two halves are already independent: the
  **language-provided** set comes from the bank and is loaded at startup; **your code's** procedures,
  variables, module names and control names are read from the live VBE project on every pass
  (`_collect_identifiers`) and are **not** configurable — that is the main path and it should not be
  something a config file can break.

  Table: `words(name, kind, is_builtin, enabled, wants_space, note)` with **`PRIMARY KEY (name, kind)`** —
  `Date` and `String` are both a function and a type, so they get one row each; collapsing them into one
  row would make the `As |` position and the expression position indistinguishable. `enabled=0` keeps a
  word without suggesting it; `is_builtin=0` marks your own words so they are **not** subjected to the
  "common vocabulary" tightening; `wants_space` is keyword-only (NULL = default rule, 1 = always insert a
  space, 0 = never); `note` is free text the program never reads.

  Decisions worth knowing. `ensure()` writes the bank **only when the file does not exist** — an
  existing bank is never touched (pinned by a test that inserts a row and calls `ensure()` again).
  `load()` opens it with a `mode=ro` URI, reads, closes immediately, and holds no lock, so you can edit
  it in DB Browser while the tool runs. ⚠️ The URI is built by hand rather than with
  `urllib.request.pathname2url`: on Windows that returns `///C:/…` (three slashes already), so adding
  `file://` produced five and SQLite answered `unable to open database file`. Journal mode is
  **DELETE, not WAL** — a WAL database needs a writable `-shm` even for a read-only connection.
  Every failure mode (garbage bytes, missing `meta`, empty bank, a `schema_version` newer than the
  program, missing file) falls back to the seed list and records the reason in `WORDBANK_INFO["error"]`;
  nothing is ever raised — an optional enhancement has no business taking completion down with it.
  A newer `schema_version` means the bank is **not read at all** rather than parsed half-blindly.

  Two names are exported deliberately: `BUILTIN_*` stays the **frozen seed list** (twenty-odd tests use
  it as "the list compiled into the code", which must not go red or green because of a user's local
  config), while `ACTIVE_*` is what is actually in effect. `ACTIVE_*` equals the seed until
  `load_active()` runs, which a test pins. The v98 lazy caches in `engine` (which lower-cased the
  lists themselves) were **deleted**: a cache and "reload after the user edits the bank" are natural
  enemies, and a stale cache shows up as "I edited the bank and nothing changed".

  Maintenance without a GUI: `py -3 wordbank.py` (summary), `list [kind]`, `add Name [kind]`,
  `disable Name`, `enable Name`, `remove Name`, `reset --yes` (backs up to `wordbank.db.bak-<time>`
  first), `export` / `import` CSV. `VBECOMPLETE_NO_WORDBANK=1` ignores the bank entirely (first thing to
  try when something looks wrong); `VBECOMPLETE_WORD_BANK=<path>` moves it.

- **v126 — the list gets out of the way, and the mouse alone never opens it.**
  Two things the user asked for, both about *when* the popup should exist rather than what is in it.

  **It closes once the name is complete.** Type a name all the way out and the list used to stay on
  screen until you pressed `Esc` or clicked away, covering the lines below. Now it disappears as soon as
  the typed word **exactly equals** a candidate **and no longer candidate starts with it** — the second
  half matters: while `ws` still has `wsName` in the pool you are not done typing, and closing there
  would eat the name you were about to finish. `Ctrl+Space` is exempt (you asked for the list; it shows
  even when the name is complete). `VBECOMPLETE_HIDE_ON_COMPLETE=0` restores the old behaviour.

  **Only keystrokes trigger it.** Moving the caret with the mouse (or pasting from a context menu, or
  dragging code around) changes the text, and the poller used to treat any same-line text change as
  typing — so clicking right after a variable you had written could raise the list. The poller now asks
  a single question first: *was there a real key press recently?* (`state["last_key_ts"]`, stamped in
  the low-level keyboard hook, plus the v79 "keys pressed but not yet in the document" counter that
  covers IME composition). No key press -> hide, don't show. Synthetic keys we inject ourselves are
  explicitly excluded (`LLKHF_INJECTED`), so nothing can self-trigger.
  `VBECOMPLETE_REQUIRE_KEY_TRIGGER=0` restores the old behaviour.

- **v126c — the list looks like the VBE's own.** The candidate list is now plain **white** and its
  scrollbar is **light grey** (`COLOR_SCROLL_TRACK` / `COLOR_SCROLL_THUMB` in `ui.py`); it used to be
  cream with a beige scrollbar, which read as a foreign window next to the VBE's own member list. The
  detail strip moved from beige `#F4F3EC` to light grey `#F7F7F7` for the same reason. All of these are
  constants at the top of `ui.py`. (Same version moved the intra-procedure jumps to
  `Ctrl+Shift+↑` / `Ctrl+Shift+↓`, see above.)

- **v126d — a type icon in front of every candidate.** The list draws a small icon to the left of
  each name, the way the VBE's own member list does, so a property and a method can be told apart at a
  glance: **green arrow** for methods / functions (`Sub`, `Function`, built-in functions such as
  `MsgBox`), **purple block with a bar** for `Property` members, **blue block** for variables and
  constants, **amber hollow box** for types (class modules, `Type`, `Enum`, built-in types) and a
  **grey diamond** for language keywords. The icons are **drawn with canvas primitives**, not shipped
  bitmaps: they scale with the row height (9–16 px), so a 24 pt editor font does not leave them
  microscopic, and every shape differs in *outline* too (arrow / barred block / solid / hollow /
  diamond) — the selected row is dark blue, where colour contrast alone would not survive. The
  classification is one function, `Completer._classify_matches`, with a fixed priority: project
  procedures > project types > built-in functions > built-in types > keywords > everything else
  (`var`). A candidate with no known kind simply gets **no** icon rather than a misleading one.

- **v127 — the module outline and find-references window are gone.** `Ctrl+Shift+O` (v121's
  module outline) and `Ctrl+Shift+F` (v121's textual find references) were removed at the user's
  request — they never got used, and a navigation window nobody opens is just surface area to keep
  testing. Everything that existed only for them went with them: `jump_ui.py`, the five backend
  methods (`current_module_text` / `list_procedures_here` / `word_at_caret` / `find_references` /
  `goto_location`), `parser.list_procedures` / `parser.find_occurrences`, `tests/probe_outline.py`,
  and the `NO_OUTLINE` / `NO_USAGES` switches (`KNOWN` is back to 55). **Both keys are handed back to
  the VBE unchanged** — the hook no longer claims them, so whatever the editor does with them is what
  you get (nothing, in a code pane). If an old `settings.ini` still contains those two switches, they
  are simply unknown keys now and change no behaviour. §118 pins the deletion with source-level
  contracts (functions, constants, switches, files and docs), so a future change cannot quietly
  reintroduce half of it.

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
| List closes by itself once the name is complete | ✅ (`Esc` no longer needed) |
| Only typing opens the list, never the mouse | ✅ (clicks / menu paste never raise it) |
| Fuzzy / camelCase matching with highlighted hits | ✅ |
| Quality-ranked results | ✅ |
| Local scope first | ✅ |
| Recently used promoted | ✅ |
| Auto-paired brackets and quotes | ✅ |
| Paired deletion on `Backspace` | ✅ |
| Smart indentation and automatic block closers | ✅ |
| Start a new line (`Shift+Enter`) | ✅ |
| Toggle line comment (`Ctrl+/`) | ✅ |
| Duplicate line (`Ctrl+D`) | ✅ |
| Delete line (`Ctrl+K`) | ✅ |
| Start a new line above (`Ctrl+Shift+Enter`) | ✅ (deeper indent when above a block closer) |
| Trailing space for keywords; `()` after functions | ✅ |
| No completion while declaring a new name | ✅ |
| Type-aware member completion | Delegated to the VBE's native member list |
| Trim / extend the built-in vocabulary | ✅ the bank (`wordbank.db`) — disable any built-in keyword/function/type/constant individually, control its trailing space, or add your own host APIs. IDEA only excludes by whole package/class, which is coarser |

**Deliberately not implemented**

Go-to-definition, rename refactoring and live syntax/error checking all depend on a full,
always-current semantic model of the project — type inference, symbol tables, a dependency graph.
The VBE exposes only the *plain text* of a `CodeModule`, so rebuilding a compiler front-end on top of
it would be costly and fragile; the risk/benefit simply doesn't add up. One partial answer exists
instead of the real thing: **live templates became a snippet library** (v115, `Ctrl+Q` — you pick a
snippet, the tool never guesses an abbreviation).

## Configuration

Every switch below is looked up in three places, first hit wins: **environment variable →
`settings.ini` in the project folder → the code default**. The environment variable still beats the
file, so a one-off experiment stays a one-liner. Inside `settings.ini` the `VBECOMPLETE_` prefix is
optional and keys are case-insensitive. Run `py -3 settings.py` in the project folder to print each
effective value *and its source* (`env` / `ini` / `default`) — that is how you find out whether the
line you edited is actually the one in effect. `VBECOMPLETE_SETTINGS=<path>` points the tool at a
different ini file (portable installs and tests).

| Variable | Default | Effect |
| --- | --- | --- |
| `VBECOMPLETE_LOG` | off | `1` writes a diagnostic log to `VbeComplete.log` (`VBECOMPLETE_LOG_FILE` relocates it). |
| `VBECOMPLETE_AUTO_PAREN` | on | `0` disables inserting `()` after procedures and functions. |
| `VBECOMPLETE_AUTO_SPACE` | on | `0` disables the automatic trailing space after keywords. |
| `VBECOMPLETE_PARAM_INFO` | off | `1` makes VbeComplete try the VBE's parameter-info menu command after inserting `()`. Off by default: on real machines that command cannot resolve a signature in this context and only beeps. |
| `VBECOMPLETE_NO_PARAM_INFO` | off | `1` forces parameter info off (wins over `VBECOMPLETE_PARAM_INFO`). |
| `VBECOMPLETE_PARAM_INFO_KEY` | off | `1` falls back to synthesising `Ctrl+Shift+I` when the host has no `CommandBars` — that path beeps too, so it is off by default. |
| `VBECOMPLETE_MEMBER_POPUP` | off | `1` restores VbeComplete's own list after `object.` (off by default). |
| `VBECOMPLETE_DECL_POPUP` | off | `1` restores suggesting inside declarations (off by default). |
| `VBECOMPLETE_HIDE_ON_COMPLETE` | on | `0` keeps the list on screen after the name has been typed out completely (v126 closes it instead). |
| `VBECOMPLETE_REQUIRE_KEY_TRIGGER` | on | `0` lets any same-line text change raise the list; the default requires a recent real key press, so mouse clicks and menu pastes only close it (v126). |
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
| `VBECOMPLETE_NO_CTRL_NAV` | off | `1` disables the `Ctrl+Shift+↑` / `Ctrl+Shift+↓` jumps; the keys stay with the VBE. |
| `VBECOMPLETE_NO_CTRL_NAV_TAB` | off | `1` keeps the `Ctrl+Shift+↓` jump but never adds the procedure-level indent to the landing line. |
| `VBECOMPLETE_NO_SNIPPETS` | off | `1` disables the `Ctrl+Q` snippet window; the key stays with the VBE. |
| `VBECOMPLETE_NO_COMMENT` | off | `1` disables `Ctrl+/` comment toggling; the key stays with the VBE. |
| `VBECOMPLETE_NO_INSERT_ABOVE` | off | `1` disables `Ctrl+Shift+Enter` (insert a line above). |
| `VBECOMPLETE_NO_DUP_LINE` | off | `1` disables `Ctrl+D` (duplicate the line). |
| `VBECOMPLETE_NO_DEL_LINE` | off | `1` disables `Ctrl+K` (delete the line). |
| `VBECOMPLETE_NO_WORDBANK` | off | `1` ignores `wordbank.db` and uses the seed list compiled into the program. |
| `VBECOMPLETE_WORD_BANK` | *(project folder)* | Path to the vocabulary bank. Empty = `wordbank.db` next to the program. |
| `VBECOMPLETE_SETTINGS` | `<project folder>\settings.ini` | Which ini file to read (portable installs / tests). |
| `VBECOMPLETE_SNIPPET_DB` | `<project folder>\snippets.db` | Where the snippet database lives (portable installs / tests point it elsewhere). |
| `VBECOMPLETE_NO_MOD_CACHE` | off | `1` goes back to re-parsing every module on each refresh (v117 behaviour). Only useful for A/B timing; the per-module cache is what keeps a refresh at ~27 ms instead of ~224 ms. |

## Project structure

| File | Role |
| --- | --- |
| `main.py` | Entry point: global keyboard hook, event loop, wiring |
| `engine.py` | Completion engine: scoring, ranking, filtering, popup lifecycle |
| `vbe_bridge.py` | VBE COM bridge: read code, read caret, write back, caret placement |
| `parser.py` | VBA text parsing (statements, scope, comments and strings) |
| `vba_builtins.py` | VBA built-in vocabulary: the **seed** list (`BUILTIN_*`) plus the effective snapshot (`ACTIVE_*`) loaded from the bank |
| `wordbank.py` | The vocabulary bank (`wordbank.db`): create / read-only load / CLI. Knows about SQLite, **not** about VBA |
| `ui.py` | The floating popup (tkinter) and its screen-avoidance logic |
| `snippet_ui.py` | The snippet window (`Ctrl+Q`): search, add / edit / delete, insert button |
| `snippets.py` | The snippet store (SQLite, two columns) and fuzzy title matching — no UI, no COM |
| `recent.py` | The recently-used table (`recent.json`): load, atomic save, corrupt-file tolerance — no UI, no COM (v119) |
| `settings.py` | The configuration centre (v120): `settings.ini` parsing, the three-layer lookup, the `KNOWN` table and the commented-out template. Imports nothing from the project — `log.py` depends on it, not the other way round. `py -3 settings.py` dumps every value with its source. |
| `settings.example.ini` | The template that `settings.py` generates, checked in so you can read it without running anything (every line commented out). |
| `log.py` | Optional diagnostic log (off by default, zero overhead) |
| `tests/` | Regression suite (`test_real_project.py`, 1172 checks) and the sample project it runs against. Six one-shot probes ship with it: `probe_ime_mode.py` prints the current IME conversion mode (how the `IME_CMODE_SYMBOL` constant below was verified), `probe_hook_gate.py` drives the real hook lifecycle against fake windows to measure how an unresponsive foreground window affects the main thread, `probe_perf.py` times the popup path (`get_context`, a full re-parse, and the scoring curve), `probe_comment.py` exercises `Ctrl+/` against a real Excel + VBE instance (it appends four lines to the end of a module, runs the toggle, deletes them, and byte-compares the module with a backup taken beforehand), `probe_lineops.py` verifies v122 on the same setup (`Ctrl+Shift+Enter` above, `Ctrl+D`, `Ctrl+K`): it appends **one temporary procedure to the end of a module**, deletes it afterwards, and byte-compares the module with the backup taken beforehand. `probe_wordbank.py` verifies v123 on a real Excel + VBE: it builds a **temporary** bank (never your own `wordbank.db`), adds a word and disables a built-in function, then checks that after `load_active()` the added word is in the pool and the disabled one is gone, and that `reset_active()` restores the baseline exactly. |

## Credits

Built by [CoffeeBean001](https://github.com/CoffeeBean001). Inspired by the completion experience
of IntelliJ IDEA.
