# VbeComplete · C# 版

同一个 VBE 代码补全工具的 C# 重写版。**Python 版原样保留**（见仓库根目录），
这一份是独立的、能单独跑的实现。

```
cs\
├─ build.bat          重新编译并发布（输出 cs\dist\VbeComplete.exe）
├─ dist\VbeComplete.exe   单文件，约 270 KB
└─ src\               源码
```

## 为什么换成 C#

先说实话：**换语言本身省不了多少 CPU**。原来那 2% 里，时间几乎全花在
"跨进程调 COM"上（那是 Excel 进程在干活），C# 调 COM 一样要付这笔钱。

真正省下来的是另外两件事，正好 C# 版做得到：

| | Python 版 | C# 版 |
|---|---|---|
| 空闲时的事件循环 | Tk `mainloop` 是地板，实测占掉 97% 的空闲 CPU | WinForms 阻塞在 `GetMessage`，空闲≈0 |
| 前台不是 VBE 时 | 仍按节奏读状态 | **一个 COM 都不发**，纯 Win32 判断（0.02ms）后直接返回 |
| 轮询频率 | 固定 | 自适应：刚敲过键 50ms，停手 150ms，不在 VBE 250ms |

### 实测（本机，2026-10-20，两版串行各跑 60 秒，都不在 VBE 代码窗里）

| | CPU（占一个核） | 工作集 | 私有内存 |
|---|---|---|---|
| Python 版 | **1.33 %** | 49.2 MB | 28.7 MB |
| C# 版 | **0.13 %** | 47.7 MB | 11.7 MB |

差距全在"不在 VBE 就一个 COM 都不发"这一条上 —— Python 版每 100ms 至少要发
`ActiveCodePane` + `GetSelection`（约 0.75ms），C# 版只读两个窗口类名（0.02ms）。
**在 VBE 里打字时两版都要付 COM 的钱**，那个差距会小得多；上面这组数字代表的是
"你没在写 VBA"时的常驻开销，也就是一天里绝大部分时间。

## 快速开始

```bat
:: 1. 编译（只需要第一次，或改了源码之后）
cs\build.bat

:: 2. 启动
run-cs.bat

:: 3. 停止（也可以右键托盘图标 -> 退出）
stop-cs.bat
```

> **启动之后看不到窗口是正常的。** 这个程序没有主窗口，只在通知区域（托盘）
> 留一个图标。Win11 默认把托盘图标折叠进时钟旁边的 `^` 里 —— 第一次用时
> 请点开那个小箭头找它。启动时会弹一个气泡提示，`run-cs.bat` 也会明确打印
> `VbeComplete C# is running.`；**如果进程真的没起来，脚本会把日志尾部打出来
> 并停住窗口**，不会再一闪而过。

需要 **.NET 10 桌面运行时**（本机已装 SDK 10.0.401，所以直接能跑）。
换一台没装 .NET 的机器时，用自包含方式发布：

```bat
dotnet publish cs\src -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o cs\dist
```

## 与 Python 版共用什么

同一套配置文件，不用重新配一遍：

- `settings.ini` —— 键名、默认值、优先级完全沿用（环境变量 > ini > 默认）
- `recent.json` —— 最近使用顺序
- `wordbank.json` —— 词库（由 `py -3 wordbank.py export --out wordbank.json` 从
  SQLite 导出一次；C# 版不内嵌 SQLite 驱动，保持单文件干净）

> ⚠️ **两个版本不要同时开**。它们都装全局键盘钩子，会抢同一个按键，
> 表现是"按 Tab 补了两次"或者候选窗互相打架。

## 已实现

- 全局键盘钩子（独立线程，回调里**一次 COM 都不碰**）
- 焦点判据：焦点窗口的父链上有 `VbaWindow` 才算"在代码窗格里"
- 快照两级读取：行列没动就复用上一轮，不重读文本
- 候选池：模块内标识符 / 声明 / 过程 / 组件名 / 内建（171 函数 + 13 类型 + 79 关键字）/ 词库
- 匹配：前缀优先，其次按词边界跳步（`sb` → `selBegin`）
- 排序：作用域档 > 命中位置 > 跨度 > 最近使用 > 非内建 > 名字长度
- **打全即收**：`sub` 打全且没有更长的候选时自动收起
- 三个"不该弹"的位置：注释/字符串里、点号后（让给 VBE）、造新名字的位置
- 候选窗：白底、淡灰滚动条、五档图标（绿箭头/紫方块/蓝竖条/琥珀菱形/灰横线）、
  深蓝选中行、底部详情行、**永不抢焦点**
- `Tab` / `Enter` 接受，`↑↓` `PgUp/PgDn` 选择，`Esc` 收起，鼠标滚轮
- **补空格**：关键字后注入一个空格（不改行文本，走 `SendInput`）
- **补括号**：Function / 内建函数补 `()` 并落进括号中间；Sub 补空格；
  函数体内给自己赋值返回值时不补；`Call xxx` 里不补
- `Ctrl+Space` 手动唤出
- `Ctrl+/` 注释切换（支持选区）
- `Ctrl+Shift+Enter` 上方插行 / `Ctrl+D` 复制行 / `Ctrl+K` 删行（光标落到上一行行尾）
- `Ctrl+Shift+↑/↓` 跳到最后一条声明 / 过程收尾空白行
- 托盘图标（打开日志 / 退出）、单实例
- 自检 99 条（含 6 条启动脚本契约、5 条列语义契约）

## 还没实现（相对 Python 版）

这些 Python 版有、C# 版**暂时没有**，用之前请知道：

- 崩溃自愈 / 看门狗（Python 版的心跳 + `watchdog.py`）
- 片段库 `Ctrl+Q`
- 回车自动缩进、`End If` / `Next` 的分支对齐
- 自动配对括号与引号（敲 `(` 自动补 `)`）
- `Backspace` 删除空行 / 删除成对符号
- 让位给 VBE 自己的成员列表（Python 版那套 `yield_pending` 裁决）
- 参数信息、宿主枚举常量

## 自检与诊断

```bat
cs\dist\VbeComplete.exe --selftest    :: 99 条纯逻辑判据，结果写 selftest.log
cs\dist\VbeComplete.exe --probe       :: 能不能连上 VBE、光标在哪、模块多长
cs\dist\VbeComplete.exe --writetest   :: 写入链路验证（末尾插一行再删掉，最后逐字节校验原文）
cs\dist\VbeComplete.exe --carettest   :: 光标落点验证（5 种缩进/全角，末尾插行再删掉）
cs\dist\VbeComplete.exe --uitest      :: 把候选窗离屏渲染成 popup-preview.png
```

`--writetest` 会真的往当前模块末尾插一行 `    Dim abc` 再改成 `    Dim abcd`，
然后删掉并**逐字节比对原文**。只在你想验证写入链路时手动跑，正常启动不会碰你的代码。

`--carettest`（v131b 为"光标落点"加的）往模块末尾插 5 种行 —— 顶格 / 4 空格 /
Tab / 全角变量名 / 右侧还有同名子串 —— 各调一次补全，然后读回**真实光标列**，
跟"按真行算出来的列"对数，最后删干净并逐字节校验原文。它比纯函数用例强的地方
只有一点：**`SetSelection` 的列语义只有真 VBE 说了算**。

开日志：`set VBECOMPLETE_LOG=1` 后再启动，日志写到 `VbeComplete.log`
（C# 版的条目带 `[cs]` 前缀，和 Python 版区分得开）。

**崩溃日志 `crash.log` 不受开关控制，永远会写。** WinExe 没有控制台，异常被
CLR 吞掉时用户只看到"双击没反应"；所以入口处挂了三个兜底
（`AppDomain.UnhandledException` / `Application.ThreadException` / `Main` 自身的
try），任何一条路崩了都会把完整堆栈落进 `crash.log`。

**重复启动不会再"没反应"**：已经有一份在跑时，新进程会通过一个命名事件
叫醒那一份，让它弹气泡告诉你"已经在运行，图标在托盘里"，然后自己安静退出。

## 源码结构

| 文件 | 职责 |
|---|---|
| `Program.cs` | 入口、单实例、自检/探测模式 |
| `App.cs` | 钩子线程、自适应轮询、按键路由、行操作 |
| `Engine.cs` | 候选池、匹配、排序、补什么（括号/空格） |
| `VbeBackend.cs` | COM 读写（快照 / 全文 / 写回 / 行操作） |
| `ComCall.cs` | `IDispatch::Invoke` 直调，只为了 `GetSelection` 那四个 ByRef 出参 |
| `ComCtx.cs` | 拿到正在运行的 VBE，缓存与探活 |
| `Parser.cs` | VBA 词法与结构（过程 / 声明 / 标识符 / 类型） |
| `Rules.cs` | 纯判据：取词、替换、列换算、匹配、位置闸门 |
| `PopupForm.cs` | 候选窗（全自绘、不抢焦点） |
| `Settings.cs` / `Log.cs` / `Recent.cs` / `WordBank.cs` | 配置 / 日志 / 最近使用 / 词库 |
| `Native.cs` | Win32：钩子、按键注入、插入符矩形、窗口类名 |
| `BuiltinTables.g.cs` | 内建名字表（由 Python 版的数据源自动导出，两边逐字一致） |

### 七个踩过的坑（改代码前值得看一眼）

1. **`dynamic` 变量不能直接 `== null`** —— 会走动态绑定去找 `operator ==`，
   对 `__ComObject` 直接抛异常。全部用 `ReferenceEquals(x, null)`。
2. **`CodePane.GetSelection` 有 4 个 ByRef 出参**，`dynamic` 的后期绑定填不回来
   （报 "Error while invoking GetSelection."）；`CodePane.Selection` 属性又不在
   IDispatch 上。只能自己拼 `DISPPARAMS` 调 `IDispatch::Invoke`（见 `ComCall.cs`）。
3. **`ReplaceLine` 会重排整行**（`,` 后、`:` 后、`=` 两边补空格），所以落光标
   的位置必须按**读回来的真实行**算，不能拿写入前的下标（这是 v126e 那个
   "光标落进 Sub 中间"的根因）。
4. **`.bat` 必须是 CRLF**。LF-only 的批处理里，`goto` 标签和多行 `if (...)` 块
   会静默失效 —— 不报错、不退出，就是"那几行没执行"。表现正是"双击窗口一闪、
   程序没启动"。`.gitattributes` 已经钉住 `*.bat text eol=crlf`，别再让编辑器
   把它改回 LF。
5. **别在 `.bat` 里用 `find` 判断进程**。Git for Windows 自带一个 GNU `find`，
   它会在 PATH 上抢先，把 `/i` 当成路径然后报 `find: '/i': No such file`，
   于是存活检查永远失败。用 `findstr`（Windows 独有，没人抢）。
6. **`GetSelection` 给的是「显示列」，不是字符列**（v131b）。纯 ASCII 行上两者
   数值相同，所以这个坑只在有中文/全角（占 2 格）或 Tab（展开到 4 格）的行上
   暴露 —— 症状是"这一行根本不弹候选"（拿显示列去切字符下标，切出个空词）。
   凡是吃 `snap.CaretCol` 的判据，先过 `Rules.CharColFromDisp`；要摆光标时反过来
   用 `Rules.DispCol`。真机实测：`    学生 = Split()` 里 `(` 前是**显示列 17、
   字符列 15**。
7. **落点列就是 `DispCol(...)` 的返回值，别再 `- 1`**（v131b）。它给的已经是
   "第 N 个字符之前"的 1-based 列；减一就落到词尾那个字符之前 —— 补完括号光标
   停在 `Spli|t`，正是用户报的"落在 Split 中间"。字符列 ↔ 显示列是一对**互逆
   换算**，不是简单的"下标 + 1"。

外加一条不算坑但很容易被误判的：**这个程序启动后屏幕上不会有任何窗口。**
如果哪天又出现"启动了但好像没反应"，先 `tasklist | findstr VbeComplete`
确认进程在不在，再看托盘溢出区，别急着判定它崩了。
