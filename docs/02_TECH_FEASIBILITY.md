# 02 · 技术可行性研究（含 A04 真正单窗口专项）

> 版本 v0.1 · 2026-11-05 · **结论均为设计阶段判断；除“本机类型库核实”外，没有任何运行行为经过实机验证。**

## 1. 本机环境调查（2026-11-05 实测）

| 项 | 结果 | 对项目的影响 |
|---|---|---|
| 操作系统 | Windows 11 专业版 24H2，10.0.26100，x64 | 目标平台；任务栏/Alt+Tab 行为以 Win11 为准，Win10 需另测 |
| Office | Office Home 2024 Retail + Visio Pro 2024，Click-to-Run，**x64**，16.0.20527.20112 | VSTO 加载项须编译为 AnyCPU/x64；Office 2024 永久版通道，Microsoft 365 月度通道行为可能不同 |
| Word / PowerPoint | `C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE` / `POWERPNT.EXE` | 两个宿主均可测试 |
| .NET Framework | 4.8.1（Release 533320） | VSTO 目标框架 4.8 可用 |
| VSTO Runtime | v4 10.0.60910 / v4R 10.0.60828（注册于 WOW6432Node），VSTOInstaller.exe 存在于 64 位 Common Files | 运行时就绪 |
| Office PIA | GAC 中有 `Microsoft.Office.Interop.Word/PowerPoint 15.0` | 可编译；Office 2016+ 新增成员（如 `Document.AutoSaveOn`）不在 PIA 15 中，需晚绑定 |
| Visual Studio | VS 2022 Community 17.14 | **未安装“Office/SharePoint 开发”工作负载**，开发 VSTO 前需安装（PoC-0） |
| .NET SDK | 9.0.203 | 可用于 Core 层单元测试项目（多目标 net48 + net8/9） |
| Python | 3.9.13 | AI 层建议另建 3.11+ 环境 |
| WebView2 Runtime | 140.0.3485.66 | 未来 AI 面板可用 |
| 显示器 | 2048×1152（主）+ 1920×1080，系统 DPI 120（125%） | 可直接测试双屏，混合 DPI 需手动调整副屏缩放 |
| 已有 Word 加载项 | BaiduImeForWord、OneNote、Adobe PDFMaker、FpcWordAddin | 共存与 IME 测试对象；PDFMaker 等也会创建窗口/事件，需纳入兼容测试 |

## 2. 本机类型库核实结果（只证明“成员存在”）

方法：PowerShell 反射 PIA 15 + `LoadTypeLibEx` 读取 `MSWORD.OLB`/`MSPPT.OLB`。脚本留存于对话 scratch 目录。

| 成员 | 结果 | 备注 |
|---|---|---|
| `Word.Window.Hwnd / Visible / WindowState / Left / Top / Width / Height / Activate / Close / WindowNumber / Caption` | 存在 | WS-1 的全部基础 |
| `Word._Document.DocID` | 存在 | 语义/稳定性【待查证】 |
| `Word._Document.AutoSaveOn` | OLB 存在，PIA 15 **不存在** | 需 `dynamic` |
| `Word._Document.IsInAutosave / Saved / FullName / ReadOnly / Final / CoAuthoring` | 存在 | |
| `Word._Application.ShowWindowsInTaskbar` | 存在 | 2013+ 是否生效【待实机】 |
| `Word._Application.UndoRecord / KeyBindings / CustomizationContext / CompareDocuments / ProtectedViewWindows / COMAddIns / IsSandboxed` | 存在 | |
| `Word.ApplicationEvents4` 全部事件 | Startup, Quit, DocumentChange, DocumentOpen, DocumentBeforeClose, DocumentBeforePrint, DocumentBeforeSave, NewDocument, WindowActivate, WindowDeactivate, WindowSelectionChange, WindowBeforeRightClick, WindowBeforeDoubleClick, WindowSize, DocumentSync, ProtectedViewWindowOpen/BeforeEdit/BeforeClose/Size/Activate/Deactivate, XML*/MailMerge*/EPostage* | **没有** DocumentAfterClose / DocumentAfterSave / WindowOpen / WindowClose 事件 → 必须对账 |
| `Word.Documents.Open` 参数 | FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, **Visible**, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform | |
| `Word.Documents.Add` 参数 | Template, NewTemplate, DocumentType, **Visible** | |
| `Word.Range.InsertFile / FormattedText / InsertXML / ExportFragment / ImportFragment / WordOpenXML / Find / GoTo` | 存在 | C 组 |
| `PowerPoint.DocumentWindow.HWND / Activate / Close / WindowState / Left/Top/Width/Height` | 存在 | E01 |
| `PowerPoint.EApplication` | 含 **PresentationCloseFinal**、AfterNewPresentation、AfterPresentationOpen、PresentationBeforeClose、WindowActivate 等 | PowerPoint 有关闭完成事件 |

## 3. 能力分级总表

| 能力 | 级别 | 依据/说明 |
|---|---|---|
| VSTO COM 加载项在 Word 2024 x64 中加载 | 官方 | VSTO 文档；本机运行时存在；**加载本身待 PoC-0 实测** |
| CustomTaskPane 停靠顶部 | 官方 | `MsoCTPDockPosition.msoCTPDockPositionTop` |
| Word 中 CTP 需按窗口创建 | 官方 | VSTO “在多个应用程序窗口中管理自定义任务窗格”文档：Word 中 CTP 与特定文档窗口关联 |
| 顶部 CTP 的最小高度、标题栏占用 | 待实机 | PoC-1 |
| 文档/窗口事件监听 | 官方 + TLB 已核实 | 事件时序待 PoC-2 |
| `Window.Activate` 切换 | 官方 | |
| `Window.Visible=false` 隐藏文档窗口 | 官方成员 + 行为待实机 | PoC-3 核心 |
| 隐藏窗口对任务栏、Alt+Tab 的影响 | 理论可行（Windows 不为不可见顶层窗口显示任务栏按钮）/ 待实机 | PoC-3 |
| 隐藏文档在退出时保存提示可见 | 待实机（最大风险点） | PoC-3 |
| `SetParent` 承载 Word 顶层窗口 | 高风险 | Win32 文档对跨进程/跨 DPI 父子窗口有警示；Word 内部窗口结构未文档化 |
| `KeyBindings` + 宏模板扩展快捷键 | 官方 | 需宏模板 |
| 线程级键盘钩子 | 公开 Win32 API / 非 Office 扩展点 | 待实机 |
| `UndoRecord` 自定义撤销组 | 官方（Word 2010+） | |
| `ExportAsFixedFormat`、`CompareDocuments`、`Range.Find`、`InsertFile` | 官方 | |
| Office.js 实现顶部标签/窗口管理 | **不可行** | 见 §4 |
| DSOFramer/ActiveX 文档承载 | **否决** | 已不受支持 |

## 4. VSTO/COM 与 Office.js 适用性比较

| 维度 | VSTO / COM（C#，.NET Framework 4.8） | Office.js（Web 加载项） |
|---|---|---|
| 运行位置 | 进程内（WINWORD.EXE），可访问完整 Word OM 与 Win32 | 浏览器运行时（WebView2）沙箱，每文档实例 |
| 跨文档能力 | 可枚举 `Application.Documents/Windows` | 只能访问当前文档；无 Application 级文档集合 |
| 窗口事件/控制 | WindowActivate 等事件、Window.Activate/Visible | 无窗口 API |
| UI 位置 | CTP 可停靠上/下/左/右；Ribbon；自定义窗体 | 任务窗格（右侧）、对话框、Ribbon 命令；**无法停靠顶部** |
| 平台 | 仅 Windows 桌面 | Windows/Mac/Web/iPad |
| 部署 | MSI/ClickOnce，需要 VSTO 运行时 | 清单 + Web 服务，集中部署方便 |
| 前景 | 微软仍支持但不再新增功能 | 微软主推方向 |
| A 组（标签/单窗口） | **唯一可行路线** | 不可行 |
| C 组（文档处理） | 完整 | 大部分可做（Word JavaScript API），但批处理多文件不可行 |
| D 组（AI 界面） | 可用 WebView2 嵌入 Web UI | 天然 Web UI |

**结论（建议）**：宿主层采用 VSTO（A/B/C 组必须）；AI 层独立为本地服务，UI 可在 VSTO 的 CTP 中通过 WebView2 呈现，从而保留未来移植 Office.js 的可能（AI 面板的 Web 前端可复用）。**不因 AI 需求而选择 Office.js 作为宿主，也不因选择 VSTO 就断言可以实现单窗口。**

备选的 COM 加载项实现方式：非托管 C++ COM Add-in（`IDTExtensibility2` + `ICustomTaskPaneConsumer`），启动更快、无 VSTO 运行时依赖，但开发成本高。仅当 VSTO 加载时延导致 Office 禁用时作为后备（风险登记 R-12）。

## 5. A04 真正单窗口专项研究

### 5.1 问题定义

Word 2013 起每个文档拥有独立顶层框架窗口（SDI）。“单窗口”有两种可能含义，需负责人明确（决策 D-01）：

| 定义 | 说明 | 可能路线 |
|---|---|---|
| **定义 U（用户可观测单窗口）** | 用户在任何时刻只看到一个 Word 工作窗口：任务栏 1 项、Alt+Tab 1 项、窗口位置不跳动、所有功能正常；底层仍是多个 Word 顶层窗口（其余隐藏） | WS-1 |
| **定义 S（结构性单窗口）** | 底层只有一个顶层框架窗口，所有文档窗口作为其子窗口被承载 | WS-2（高风险）或 Word 原生 MDI（2013+ 已不存在） |

本 Agent 建议以**定义 U**作为 SINGLE-WINDOW-PASS 的判定口径（验收标准 SW-01~SW-14 全部基于可观测行为），理由：需求原文的目标是“用户看到一个主要 Word 工作窗口……避免出现多个独立 Word 主窗口和任务栏项”，这是可观测体验目标；定义 S 只能通过非公开窗口结构操作获得。**但这是需要负责人批准的口径决定，不是本 Agent 可以自行降级的事项。**

### 5.2 候选路线评估

| 维度 | WS-0 NativeSwitch | WS-1 可见性编排 | WS-2 Win32 宿主 |
|---|---|---|---|
| 使用的接口 | `Window.Activate` | `Window.Visible/Activate/Left/Top/Width/Height/WindowState` + Win32 `Get/SetWindowPlacement`（作用于本进程窗口） | `SetParent`、`SetWindowLongPtr(GWL_STYLE)`、消息转发 |
| 是否修改 Word 窗口结构 | 否 | 否（只改可见性/位置） | **是** |
| 非公开依赖 | 无 | 无（行为组合未承诺） | 依赖 OpusApp 窗口结构与 Word 内部焦点/owner 逻辑 |
| Hook/注入 | 无 | 无 | 可能需要子类化（进程内）|
| 任务栏 1 项 | ✗ | 预期 ✓（待实机） | ✓ |
| Alt+Tab 1 项 | ✗ | 预期 ✓（待实机） | ✓ |
| 窗口位置稳定 | ✗ | 依赖同步精度与闪烁 | ✓ |
| Ribbon/Backstage/对话框 | 原生 | 原生 | 高风险 |
| DPI 风险 | 无 | 低-中（placement 换算） | 高（子窗口不收 WM_DPICHANGED） |
| Office 更新冲突 | 无 | 低-中 | 高 |
| 退出恢复 | 无需 | `Visible=true` 全部恢复（简单） | 需完整逆操作，失败会导致窗口“丢失” |
| 卸载残留 | 无 | Ledger 清理 | 同左 + 窗口样式残留风险 |
| 可达闸口 | SWITCH-PASS | SINGLE-WINDOW-PASS（定义 U）候选 | SINGLE-WINDOW-PASS（定义 S）候选 |

### 5.3 WS-1 关键未知点（PoC-3 必须回答）

| # | 问题 | 失败后果 | 可能缓解 |
|---|---|---|---|
| Q1 | `Window.Visible=false` 后任务栏按钮与 Alt+Tab 项是否消失 | WS-1 不成立 | Win32 `ShowWindow(SW_HIDE)` 对照实验 |
| Q2 | 隐藏后 Word 是否在某些操作（激活其他窗口、Backstage、打印、保存）时自行把窗口显示出来 | 出现第二窗口 | Watchdog 修复；若频繁 → WS-1 不成立 |
| Q3 | 退出 Word 时隐藏的未保存文档的保存提示是否可见、可操作 | **数据丢失/假死风险** | `Quit` 前依次显示；`DocumentBeforeClose` 中让目标窗口可见 |
| Q4 | 位置同步精度、闪烁程度（先显后隐 vs 先隐后显） | 体验不达标 SW-14 | DWM 过渡/延迟隐藏；需负责人确认阈值 |
| Q5 | 新打开文档窗口首帧闪现位置 | 体验问题 | 在 DocumentOpen 中尽早同步；或用 `Documents.Open(Visible:=false)` 由插件打开的路径先隐藏后定位再显示 |
| Q6 | Ctrl+F6、“视图→切换窗口”、任务栏 Jump List 能否激活隐藏窗口 | 出现第二窗口或无法访问文档 | 将其视为切换请求 |
| Q7 | 隐藏窗口的 CTP、Ribbon 回调是否有异常 | 标签栏错乱 | 只在显示后刷新 |
| Q8 | 主视窗最小化/最大化/贴靠/跨屏时同步正确性 | 窗口跳动 | placement 包含显示器信息 |
| Q9 | 插件异常时恢复是否 100% 可靠 | 文档“丢失” | Recovery 进程 + Ledger |
| Q10 | `Application.ShowWindowsInTaskbar=false` 在 Word 2024 是否有任何效果 | — | 仅作信息收集 |

### 5.4 WS-2 已知/推测风险（仅供决策）

- `SetParent` 将顶层窗口变为子窗口时需自行调整 `WS_CHILD/WS_POPUP` 样式，Win32 文档明确要求调用方处理样式与 UI 状态；跨 DPI 感知上下文的父子关系受限制。
- Word 的模态对话框、Backstage、Ribbon 下拉、KeyTips 均以 Word 顶层窗口为 owner/参照，成为子窗口后行为未知。
- Office 每月更新可能改变窗口层次。
- 结论：**不进入生产路线；仅在 PoC-3 失败且负责人批准后，在隔离 VM（可快照回滚）中实验。**

### 5.5 若 WS-1 失败的决策选项（提交负责人）

1. 接受 WS-0（SWITCH-PASS）作为 v1 正式能力，明确标注“非单窗口”，A04 保持“未达标”状态继续研究；
2. 批准 WS-2 隔离实验，接受其高风险与可能的长期维护成本；
3. 调整产品定位（例如以“工作区 + 快速切换”为主卖点）。
**本 Agent 不会在未获批准时选择任何一项。**

## 6. 技术原型（PoC）计划

> 所有 PoC 在独立分支 `poc/*` 与独立解决方案中进行，不进入产品代码；在执行前需负责人批准范围。安全措施通用要求：①使用测试文档副本（`%USERPROFILE%\CrabOfficeTest\`），不打开真实工作文档；②Word 选项开启 AutoRecover 1 分钟；③实验插件仅注册在 HKCU，可一键删除注册项（附 `uninstall-poc.ps1`）；④WS-2 只在虚拟机中执行。

### PoC-0 工具链与加载验证（0.5 天）
- **目标**：确认 VSTO Word 加载项可在本机 Word 2024 x64 中构建、加载、调试。
- **步骤**：安装 VS“Office/SharePoint 开发”工作负载 → 新建 Word VSTO Add-in（.NET 4.8）→ Startup 写日志 → F5 调试 → 记录加载耗时（`HKCU\Software\Microsoft\Office\16.0\Word\AddInLoadTimes`【待查证键名】）。
- **验收**：加载成功、加载耗时 < 500ms、卸载注册项后 Word 无残留。

### PoC-1 顶部标签栏 UI（2 天）→ UI-PASS 证据
- **目标**：验证顶部 CTP 的视觉可行性。
- **步骤**：为每个 Word.Window 创建 Top CTP，承载 WinForms 自绘标签控件（假数据）→ 测量最小高度、标题栏占用 → 测试 Ribbon 折叠/展开、最大化、跨屏、DPI 100/125/150/200% → 同时做 ElementHost(WPF) 对照。
- **验收**：截图集 + 测量表；标签栏高度 ≤ 40px@100%（目标值，待确认）；无遮挡正文；记录不可控项（如 CTP 标题栏）。

### PoC-2 事件时序与身份（2 天）→ SWITCH-PASS 基础
- **目标**：建立真实的 Word 事件时序表，验证身份映射。
- **步骤**：订阅全部 ApplicationEvents4 事件写入带时间戳日志 → 执行 10 种打开路径、5 种关闭路径（含取消）、新建窗口、PV→编辑、另存为改名 → 每次事件后记录 `Windows.Count`、各 `Hwnd`、RCW 引用是否相同、`DocID` 值 → 记录取消关闭时异常的 HRESULT。
- **验收**：完整时序表（05 附录）；身份匹配方案在 200 次随机操作中 0 错配。

### PoC-3 单窗口 WS-1 可见性编排（4–5 天）→ SINGLE-WINDOW-PASS 闸口
- **目标**：回答 §5.3 的 Q1–Q10，判断 WS-1 是否能达到 SW-01~SW-14。
- **范围**：仅官方 OM + `Get/SetWindowPlacement`；不使用 Hook、注入、SetParent。
- **步骤**：
  0. Step 0：设置 `ShowWindowsInTaskbar=false` 观察 10 分钟（Q10）。
  1. 3 个测试文档，手动按钮执行 `Visible=false/true`，观察任务栏/Alt+Tab/Win+Tab（Q1）。
  2. 实现最小 `SwitchTo`（先显后隐 / 先隐后显两种顺序），60fps 录屏比较闪烁（Q4）。
  3. 在隐藏状态下：编辑可见文档 → Backstage/打印/另存为/字体对话框/宏提示（Q2、Q7）。
  4. 隐藏文档中保留未保存修改 → 退出 Word（Q3），重复 10 次。
  5. Ctrl+F6、视图→切换窗口、任务栏右键 Jump List 打开文件（Q5、Q6）。
  6. 最小化/最大化/贴靠/跨屏/混合 DPI（Q8）。
  7. 故障注入：在隐藏状态下抛未处理异常、禁用加载项、结束插件（Recovery 进程恢复）（Q9）。
- **安全措施**：启动即写 Ledger；紧急热键与 Ribbon“显示所有窗口”；外部 `show-all-word-windows.ps1`（EnumWindows + ShowWindow，仅限 WINWORD 进程）随时可执行。
- **验收**：SW-01~SW-14 逐项打分表 + 录屏证据；任何一项失败即判 WS-1 未达标，提交决策 D-01/D-02。

### PoC-4 WS-2 Win32 宿主（仅批准后，5 天，VM 内）
- **目标**：评估定义 S 的可行性与代价。
- **步骤**：自建 Frame → SetParent 一个 Word 窗口 → 检查 Ribbon/Backstage/对话框/焦点/IME/DPI → 逆操作恢复测试 → 模拟插件崩溃。
- **验收**：同 SW 矩阵 + 恢复测试；输出风险报告；默认结论为“不建议生产”，除非全部通过且负责人接受维护成本。

### PoC-5 快捷键机制（1.5 天）
- **目标**：K1（KeyBindings + 全局模板）与 K2（线程级钩子）对比。
- **步骤**：两种机制实现“下一个标签”→ 在正文/表格/页眉/IME 组字/对话框 5 场景测试 → 卸载后检查 Normal.dotm 与 KeyBindings 残留。
- **验收**：选型报告 + 冲突清单。

### PoC-6 AI 桥接（2 天，A 组完成后）
- **目标**：VSTO ↔ 本地 FastAPI/LangGraph 的结构化往返、取消、超时、令牌鉴权。
- **验收**：选中文本润色往返 P95 < 模型耗时 + 200ms；服务未启动时插件无异常。

## 7. 官方文档参考（评审前需复核；本轮联网检索不可用，链接基于知识整理）

| 主题 | 链接 |
|---|---|
| VSTO 自定义任务窗格 | https://learn.microsoft.com/visualstudio/vsto/custom-task-panes |
| 在多个窗口中管理 CTP（Word 每窗口） | https://learn.microsoft.com/visualstudio/vsto/custom-task-panes （“Manage custom task panes in multiple application windows”章节） |
| VSTO 线程支持 | https://learn.microsoft.com/visualstudio/vsto/threading-support-in-office |
| 从其他解决方案调用 VSTO 加载项代码（RequestComAddInAutomationService） | https://learn.microsoft.com/visualstudio/vsto/calling-code-in-vsto-add-ins-from-other-office-solutions |
| VSTO 加载项生命周期 / Shutdown 不保证调用 | https://learn.microsoft.com/visualstudio/vsto/events-in-office-projects |
| Word `Window` 对象（Hwnd、Visible、Activate） | https://learn.microsoft.com/office/vba/api/word.window |
| Word `Application` 事件（DocumentBeforeClose、WindowActivate 等） | https://learn.microsoft.com/office/vba/api/word.application |
| Word `Documents.Open` | https://learn.microsoft.com/office/vba/api/word.documents.open |
| Word `Document.Saved` | https://learn.microsoft.com/office/vba/api/word.document.saved |
| Word `UndoRecord` | https://learn.microsoft.com/office/vba/api/word.undorecord |
| Word `KeyBindings.Add` / `CustomizationContext` | https://learn.microsoft.com/office/vba/api/word.keybindings.add |
| Word `ProtectedViewWindow` | https://learn.microsoft.com/office/vba/api/word.protectedviewwindow |
| Word `ExportAsFixedFormat` | https://learn.microsoft.com/office/vba/api/word.document.exportasfixedformat |
| Word `Application.CompareDocuments` | https://learn.microsoft.com/office/vba/api/word.application.comparedocuments |
| Office 高 DPI 处理 | https://learn.microsoft.com/office/client-developer/ddpi/handle-high-dpi-and-dpi-scaling-in-your-office-solution |
| Win32 `SetParent` | https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-setparent |
| Win32 `SetWindowPlacement` | https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-setwindowplacement |
| Office 加载项（Office.js）概述 | https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins |
| PowerPoint `PresentationCloseFinal` | https://learn.microsoft.com/office/vba/api/powerpoint.application.presentationclosefinal |
| Office 禁用缓慢/崩溃 COM 加载项（Resiliency） | 【待查证具体页面】 |
