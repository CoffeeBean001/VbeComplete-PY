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
| 常驻内存 | 46 MB | 约 35 MB（无 Python 解释器） |

## 快速开始

```bat
:: 1. 编译（只需要第一次，或改了源码之后）
cs\build.bat

:: 2. 启动
run-cs.bat

:: 3. 停止（也可以右键托盘图标 -> 退出）
stop-cs.bat
```

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
- 自检 85 条

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
cs\dist\VbeComplete.exe --selftest    :: 85 条纯逻辑判据，结果写 selftest.log
cs\dist\VbeComplete.exe --probe       :: 能不能连上 VBE、光标在哪、模块多长
cs\dist\VbeComplete.exe --writetest   :: 写入链路验证（末尾插一行再删掉，最后逐字节校验原文）
cs\dist\VbeComplete.exe --uitest      :: 把候选窗离屏渲染成 popup-preview.png
```

`--writetest` 会真的往当前模块末尾插一行 `    Dim abc` 再改成 `    Dim abcd`，
然后删掉并**逐字节比对原文**。只在你想验证写入链路时手动跑，正常启动不会碰你的代码。

开日志：`set VBECOMPLETE_LOG=1` 后再启动，日志写到 `VbeComplete.log`
（C# 版的条目带 `[cs]` 前缀，和 Python 版区分得开）。

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

### 三个踩过的坑（改代码前值得看一眼）

1. **`dynamic` 变量不能直接 `== null`** —— 会走动态绑定去找 `operator ==`，
   对 `__ComObject` 直接抛异常。全部用 `ReferenceEquals(x, null)`。
2. **`CodePane.GetSelection` 有 4 个 ByRef 出参**，`dynamic` 的后期绑定填不回来
   （报 "Error while invoking GetSelection."）；`CodePane.Selection` 属性又不在
   IDispatch 上。只能自己拼 `DISPPARAMS` 调 `IDispatch::Invoke`（见 `ComCall.cs`）。
3. **`ReplaceLine` 会重排整行**（`,` 后、`:` 后、`=` 两边补空格），所以落光标
   的位置必须按**读回来的真实行**算，不能拿写入前的下标（这是 v126e 那个
   "光标落进 Sub 中间"的根因）。
