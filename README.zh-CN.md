# VbeComplete — VBA 变量名/过程名自动补全工具

*[English](README.md) · **简体中文***

![VbeComplete 实际效果：输入 "cus" 实时过滤候选，Tab 确认补全](assets/demo.gif)

VBE 原生只对 `对象.成员` 弹列表，**不对你自己起的变量名/过程名弹列表**，本工具补的就是这一块：
在代码窗敲名字的前几个字符（支持中文名，如 `Dim 用户名 As String`），候选列表实时过滤，
`Tab` 确认、`↑`/`↓` 选择、`Esc` 取消。支持 **Office 全家桶**（Excel / Word / PowerPoint /
Access / Visio）与 **WPS 全家桶**，按前台窗口自动识别宿主。

## 注意事项

1. **Windows** + 上面任一宿主。
2. **Python 3.8 及以上**（推荐 3.10 ~ 3.12）；安装时勾选 *Add python.exe to PATH*。
3. **必须勾选"信任对 VBA 工程对象模型的访问"，然后重启该应用** ——
   **文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置**。这一项和"启用宏"是两回事，
   不勾选工具完全读不到代码。（WPS 在自己的选项里有同名开关。）
4. **改了代码必须重启进程**：**`stop.bat` → 覆盖文件 → `run.bat`**
   —— Python 只在启动时把 `.py` 读进内存，只覆盖文件跑的仍是旧代码。

## 安装与运行

```bat
git clone https://github.com/CoffeeBean001/VbeComplete-V587.git
cd VbeComplete-V587
run.bat
```

首次运行会装依赖（需联网），之后**后台静默运行**。`stop.bat` 结束后台进程 ·
`install_autostart.bat` / `uninstall_autostart.bat` 加 / 撤开机自启 ·
`run_tests.bat` 跑离线回归测试。

| 操作 | 作用 |
| --- | --- |
| 输入字母 / 数字 / `_` | 弹出候选列表并实时过滤 |
| `↑` / `↓` | 上下选择（循环） |
| `Shift+↑` / `Shift+↓` | 收起列表，同时把光标上/下移一行 |
| `Tab` | **确认补全**（唯一确认键） |
| 单击候选 | 用鼠标确认 |
| `Esc` / `←` / `→` / 点击别处 | 收起列表 |
| `Ctrl+Space` | 强制弹出 |
| `Enter`（行尾） | 换行 + 智能缩进；块头自动补 `End Sub` / `Next i` 等收尾 |
| `Shift+Enter` | 在当前行下方新起一行，不把本行拆开 |
| `Backspace` | 整行只有空白时一下删掉整行；光标在空括号/空引号正中间时两端同删 |

确认后：`Sub` 补一个空格，`Function` 补 `()` 并把光标放在括号里，关键字补空格。
会提示工程里的变量 / 常量 / 过程 / 模块 / 窗体 / 控件名，以及关键字与内建函数（约 260 个）；
**枚举（`vb*` / `xl*` / `mso*`）刻意不收**。**正在定义**新名字（`Dim …` / `Sub …`）时不弹，
和 IDEA 一致；`As` / `New` 后面的类型位置照常提示。

**不弹列表？** 先查第 3 条（信任设置，九成是这个）；再看任务管理器里有没有 `pythonw.exe`；
列表关不掉就 `stop.bat` 再 `run.bat`。

## 更多文档

[功能详解](docs/features.zh-CN.md) · [排错手册](docs/troubleshooting.zh-CN.md) ·
[功能对照 IDEA（英文）](docs/features.md) · [MIT License](LICENSE)
