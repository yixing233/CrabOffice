# 01 · CrabOffice 完整功能规格（PRD + 研发规格）

> 版本：v0.1（设计评审稿） · 日期：2026-11-05 · 状态：**仅设计，未编码，未实机验证任何行为**
> 适用环境（本机调查结果，见 [00_EXECUTIVE_SUMMARY.md](00_EXECUTIVE_SUMMARY.md)）：Windows 11 专业版 24H2（10.0.26100）· Microsoft 365/Office Home 2024 Retail C2R x64 16.0.20527.20112 · .NET Framework 4.8.1 · VS 2022 Community（**未安装 Office/SharePoint 开发工作负载**）· Python 3.9.13 · 双显示器（2048×1152 主屏 + 1920×1080）· 系统 DPI 120（125%）。

## 0. 阅读约定

### 0.1 验证状态标签（所有文档通用）

| 标签 | 含义 |
|---|---|
| 【官方】 | 微软官方文档明确记载的能力（链接见 [02_TECH_FEASIBILITY.md](02_TECH_FEASIBILITY.md) §7，本轮联网检索不可用，链接需评审前复核） |
| 【TLB 已核实】 | 已在本机 `Microsoft.Office.Interop.Word 15.0` PIA 或 `MSWORD.OLB` / `MSPPT.OLB` 类型库中核实**成员存在**（2026-11-05）。只证明 API 存在，**不证明运行行为** |
| 【待实机】 | 方案合理但运行行为需 PoC 实机验证 |
| 【高风险】 | 依赖未文档化行为、Win32 窗口操纵或 Office 内部窗口结构 |
| 【待查证】 | 不确定 API/行为，禁止据此直接开发 |
| 【否决】 | 已评估不采用 |

### 0.2 三层验收闸口（不可互相冒充）

| 闸口 | 定义 | 主要相关功能 |
|---|---|---|
| **UI-PASS** | 顶部标签栏可见、样式正确、交互顺畅 | A01、A07–A10、A15 |
| **SWITCH-PASS** | 标签能准确切换 Word 文档，关闭/保存与原生兼容 | A02、A03、A05、A06、A11–A14 |
| **SINGLE-WINDOW-PASS** | 用户只看到**一个**主要可交互 Word 工作窗口；任务栏/Alt+Tab/焦点/对话框/打印/保存全部达标 | **A04**（依赖 A01–A16 全部） |

### 0.3 本文统一术语

- **DocumentRecord / WindowRecord / TabRecord**：见 [03_ARCHITECTURE_AND_INTERFACES.md](03_ARCHITECTURE_AND_INTERFACES.md) §4。一个 Document 可有多个 Window（Word “新建窗口”产生 `doc.docx:1`、`doc.docx:2`），**一个 Tab 对应一个 Window**。
- **WS-0 / WS-1 / WS-2**：窗口体验策略（`IWindowExperienceStrategy` 实现）。
  - **WS-0 NativeSwitch**：保持 Word 原生多窗口，标签只负责切换 → 最高只能达到 SWITCH-PASS。
  - **WS-1 VisibilityOrchestration（可见性编排）**：同一时刻只保留一个 Word 文档窗口可见，其他文档窗口通过官方 `Window.Visible=false` 隐藏，切换时同步位置/尺寸/状态后交换可见性 → **SINGLE-WINDOW-PASS 的首选候选，待 PoC-3 验证**。
  - **WS-2 Win32HostFrame（重设父窗口）**：自建宿主框架窗体，通过 `SetParent` 承载 Word 窗口 → 仅隔离实验，【高风险】。
- **CommandDispatcher**：所有对 Word 产生副作用的操作的唯一入口（UI、快捷键、命令面板、AI 白名单都经过它）。
- **Reconcile（对账）**：以 `Application.Documents` / `Application.Windows` / `Application.ProtectedViewWindows` 的实际枚举结果为真相源，修正 Registry。

### 0.4 功能总览与优先级

| ID | 名称 | 优先级 | 难度 | 风险 | 闸口 |
|---|---|---|---|---|---|
| A01 | 顶部标签栏 | P0 | 中 | 待实机 | UI |
| A02 | 打开文档同步生成标签 | P0 | 中 | 官方+待实机 | SWITCH |
| A03 | 标签点击切换 | P0 | 低-中 | 官方 | SWITCH |
| **A04** | **真正单窗口多文档** | **P0 闸口** | **高** | **待实机/高风险** | **SINGLE-WINDOW** |
| A05 | 新建和打开文件 | P0 | 低 | 官方 | SWITCH |
| A06 | 标签关闭与保存保护 | P0 | 中 | 官方+待实机 | SWITCH |
| A07 | 未保存标记 | P0 | 低-中 | 官方+待实机 | UI |
| A08 | 标签拖拽排序 | P1 | 中 | 低 | UI |
| A09 | 固定标签 | P1 | 低 | 低 | UI |
| A10 | 标签溢出/搜索/滚动 | P1 | 中 | 低 | UI |
| A11 | 右键菜单 | P1 | 低-中 | 官方 | SWITCH |
| A12 | 键盘快捷切换 | P1 | 中 | 待实机 | SWITCH |
| A13 | 文档和窗口身份模型 | P0（基础） | 中-高 | 待实机 | SWITCH |
| A14 | 事件同步与容错 | P0（基础） | 高 | 待实机 | SWITCH |
| A15 | DPI/多显示器/主题 | P1 | 中-高 | 待实机 | UI |
| A16 | 禁用/卸载/故障恢复 | P0（基础） | 中-高 | 待实机 | ALL |
| B01–B05 | 会话/工作区/文档管理 | P2 | 低-中 | 低 | — |
| C01–C08 | Word 效率增强 | P3 | 中-高 | 官方为主 | — |
| D01–D06 | AI 智能增强 | P4 | 高 | 模型/数据安全 | — |
| E01–E03 | PowerPoint 扩展 | P5 | 中-高 | 待实机 | — |

> 优先级只代表研发顺序，**不代表永久排除**其他模块。

---

# A 组：多标签与窗口管理（第一优先级）

## A01 顶部标签栏

**元信息**：P0 · 难度 中 · 风险【待实机】 · 闸口 UI-PASS · 依赖 A13/A14/A15

**1. 功能说明**
在每个 Word 文档窗口的功能区（Ribbon）下方、正文上方提供浏览器式横向标签条；每项显示文档图标、名称、选中态、未保存标记、固定标记、关闭按钮。解决 Word 2013+ 为 SDI（每文档一个顶层窗口）后，多文档只能靠任务栏/“切换窗口”菜单切换、效率低的问题。**本功能只负责“显示与交互”，不承担单窗口目标。**

**2. 用户流程**
1. Word 启动 → 插件加载 → 每个现有文档窗口顶部出现标签条（插件中途启用同理，A02 初始快照）。
2. 用户打开/新建文档 → 所有窗口的标签条同时新增一项（共享 `TabRegistry` 状态）。
3. 悬停标签 → 显示完整路径 Tooltip；点击 → A03；点 × → A06；右键 → A11；拖拽 → A08。
4. 用户可通过 Ribbon「CrabOffice」选项卡 →「显示/隐藏标签栏」切换显示。

**3. UI 设计**（详细见 [04_UI_SPEC.md](04_UI_SPEC.md)）
- 入口：VSTO `CustomTaskPane`，`DockPosition = msoCTPDockPositionTop`，`DockPositionRestrict = msoCTPDockPositionRestrictNoChange`（禁止用户拖到侧边）。
- 目标高度 32px@96DPI（逻辑像素，按 DPI 缩放）；实际最小高度和 CTP 标题栏占用**待 PoC-1 测量**。
- 状态：默认 / 悬停 / 选中 / 未保存 `●` / 固定（`fa-thumbtack`）/ 只读（`fa-lock`）/ 受保护视图（`fa-shield-halved`）/ 错误（`fa-triangle-exclamation`）/ 溢出（A10）。
- 异常提示：标签栏初始化失败时不弹模态框，仅在 Ribbon 上显示状态图标 + 日志；不影响 Word 编辑。

**4. 技术实现**
- 模块：`CrabOffice.Word.Host`（VSTO 外接程序项目）→ `TabStripPaneManager`；UI 控件 `CrabOffice.UI.TabStrip`（WinForms 自绘 `TabStripControl : UserControl`，首选；WPF 通过 `ElementHost` 作为备选，见 03 §3.3）。
- 步骤：
  1. `ThisAddIn_Startup`：捕获 `SynchronizationContext.Current`（WindowsForms），初始化 Core。
  2. `TabStripPaneManager.EnsurePane(Word.Window wn)`：`this.CustomTaskPanes.Add(new TabStripHost(), "CrabOffice", wn)`【官方：Word 中 CTP 与具体文档窗口关联，需每窗口创建】→ 设置 Dock/Height/Visible。
  3. 以 `WindowId` 为 Key 存储 `Dictionary<WindowId, CustomTaskPane>`；窗口关闭对账后 `CustomTaskPanes.Remove(pane)` 并 Dispose 控件。
  4. 所有 `TabStripControl` 订阅 `ITabService.TabsChanged`（UI 线程），统一重绘；控件只发出意图（`ActivateRequested(tabId)` 等），不直接调用 Word OM。
  5. 监听 `CustomTaskPane.VisibleChanged`：用户点 CTP 自带 × 关闭时，记入设置“此窗口隐藏”，并提供 Ribbon 恢复入口。
  6. `DockPositionChanged` 被触发（理论上 Restrict 后不会）→ 强制改回 Top 并记录日志。
- 主题：读取 Office 主题（注册表 `HKCU\Software\Microsoft\Office\16.0\Common\UI Theme` 值【待查证键名与取值】）+ 系统深浅色，映射到调色板；无法读取时使用中性灰配色。

**5. 数据结构**
```text
PaneBinding { windowId, CustomTaskPane pane(运行期, 不持久化), TabStripControl control, createdAt, lastHeightPx }
TabStripViewState { scrollOffset, hoverTabId?, dragState?, searchText? }   // 每个窗口独立
UiSettings { tabBarVisible, tabHeightDip=32, showFullPathTooltip, theme=Auto|Light|Dark }
```

**6. 模块接口**
- 读：`ITabService.GetTabs()`、`TabsChanged` 事件；`IThemeService.ThemeChanged`。
- 写：仅通过 `ITabService.ActivateTab/CloseTab/MoveTab/PinTab`，再由 `ICommandDispatcher` 落地。
- 被 A04 使用：WS-1 下每个窗口都有自己的 Pane，但只有可见窗口的 Pane 被用户看到。

**7. 边界场景**
- 受保护视图窗口（`ProtectedViewWindow`）：CTP 能否附着于 PV 窗口【待查证/待实机】；若不能，PV 文档在其他窗口标签条上以“受保护视图”项显示。
- 阅读模式（Read Mode）、全屏、Backstage（文件菜单）覆盖时：CTP 可能被隐藏，属 Word 行为，不强制处理。
- 嵌入式编辑（如 Excel 中编辑嵌入 Word 对象）、Outlook WordMail：不在 `Application.Windows` 内或不加载本插件，**不显示标签**。
- 窗口极窄（< 400px）：标签压缩为仅图标 + 溢出菜单。

**8. 技术风险**
- 【官方】`CustomTaskPane`、`msoCTPDockPositionTop`、每窗口创建 CTP。
- 【待实机】顶部 CTP 的最小高度、是否有标题栏/关闭按钮占位、与 Ribbon 折叠/展开时的重排、启动时闪烁。
- 【高风险，备选】若 CTP 视觉不可接受，备选为“在 Word 窗口内创建 Win32 子窗口叠加”，需单独 PoC，**默认不采用**。
- 不得宣称标签条位于 Ribbon 内部：它是独立停靠窗格，位于 Ribbon 与文档区之间。

**9. 验收标准**
- 打开 3 个文档，每个窗口顶部均显示 3 个标签；选中态与当前窗口一致；× 可见可点。
- 窗口最大化/还原/拖动跨屏、Ribbon 折叠/展开、DPI 100/125/150/200% 下标签无裁切、无遮挡正文首行。
- 关闭标签栏后 Word 编辑功能 100% 正常；重新显示后状态正确。
- 结论记为 **UI-PASS**，不得据此宣布 A04 完成。

---

## A02 打开文档同步生成标签

**元信息**：P0 · 难度 中 · 风险【官方】+【待实机】 · 闸口 SWITCH-PASS · 依赖 A13/A14

**1. 功能说明**
无论文档以何种方式进入 Word（双击资源管理器、Ctrl+O、最近文件、拖入 Word、插件打开、新建、模板新建、命令行、受保护视图打开、邮件附件），标签都自动增加；关闭后自动减少；无重复、无幽灵标签。

**2. 用户流程**
用户任意方式打开 → ≤300ms 内所有窗口标签条出现新标签并选中（若该窗口为活动窗口）→ 关闭后对账移除。

**3. UI 设计**
- 新标签插入位置：默认插入在“当前活动标签右侧”（可在设置改为“末尾”）。
- 加载中状态：文档打开过程中（大文件）显示 `fa-spinner` 动画；打开失败不生成标签。
- 异常：对账发现 Registry 与 Word 不一致时**静默修正**并写日志，不打扰用户。

**4. 技术实现**
- 事件源【TLB 已核实：`ApplicationEvents4` 含 `DocumentOpen/NewDocument/DocumentChange/DocumentBeforeClose/WindowActivate/WindowDeactivate/WindowSize/ProtectedViewWindowOpen/ProtectedViewWindowActivate/ProtectedViewWindowBeforeEdit/ProtectedViewWindowBeforeClose/Quit`；**无** DocumentAfterClose/DocumentAfterSave 事件】。
- 模块：`WordEventAdapter`（订阅 COM 事件，转换为 `HostEvent`）→ `HostEventQueue`（UI 线程合并）→ `DocumentRegistry.ApplyHostEvent()` → `TabRegistry` → `TabsChanged`。
- 步骤：
  1. 启动/插件中途启用：`InitialSnapshot()` 遍历 `Application.Windows`（每个 Window → Document）和 `Application.ProtectedViewWindows`，建立 Registry。
  2. `DocumentOpen(Doc)` / `NewDocument(Doc)`（C# 中需用 `((Word.ApplicationEvents4_Event)app).NewDocument += …` 以避开同名方法）→ 入队 `DocumentAppeared`；**不在事件处理器内做重活**，用 `SynchronizationContext.Post` 延后到事件返回后执行，以免 Word 内部状态未就绪【待实机：Doc.Windows 在 DocumentOpen 时是否已可用】。
  3. `WindowActivate(Doc, Wn)` → 若 Wn 不在 Registry（例如“新建窗口”命令、未触发 DocumentOpen 的路径）→ 补登记。
  4. 对账触发点：`WindowActivate`、`DocumentChange`、关闭命令执行后 300ms、2s 低频定时器（仅比较 Count 与 Hwnd 集合，避免全量 COM 遍历）。
  5. 去重：同一 `Word.Window` 依据 `Hwnd`【TLB 已核实 `Window.Hwnd`】+ RCW 引用比对（A13）判断是否已登记。
- 受保护视图：`ProtectedViewWindowOpen` → 登记 `WindowKind=ProtectedView`；`ProtectedViewWindowBeforeEdit` 后会产生新的普通 Document/Window，需将旧 PV 记录迁移为新记录并保持标签位置。

**5. 数据结构**
```text
HostEvent { seq(long), kind: DocumentAppeared|WindowAppeared|WindowActivated|DocumentBeforeClose|PvOpened|PvBeforeEdit|PvBeforeClose|AppQuit, hwnd?, receivedAtUtc, comRef(弱/运行期) }
ReconcileReport { added[], removed[], updated[], durationMs, trigger }
```

**6. 模块接口**
`IWordHostAdapter.EnumerateDocumentsAndWindows()` → `IDocumentRegistry.Reconcile(snapshot)` → `RegistryChanged` → `ITabService`。B01（最近关闭）、B03（会话）订阅 `DocumentRemoved`。

**7. 边界场景**
- 同一文件重复打开：Word 一般激活已有文档【待实机】，不得新增标签。
- 打开失败/取消密码框：不触发 DocumentOpen 或文档立即消失 → 对账清理。
- 隐藏打开（`Documents.Open(Visible:=false)`，第三方插件常用）：默认**不显示**为标签（Window.Visible=false 且非 CrabOffice 自己隐藏的），在设置中可选“显示后台文档”。WS-1 下必须区分“CrabOffice 隐藏”与“第三方隐藏”（WindowRecord.hiddenBy）。
- 模板（.dotx/.dotm）以编辑方式打开：正常显示，图标区分。
- 多 Word 进程（/x 启动或自动化创建的实例）：每个进程各自加载插件，Registry 进程内隔离；不跨进程合并（P3 以后再评估命名管道聚合）。

**8. 技术风险**
- 【官方】上述事件存在且有文档。
- 【待实机】各打开路径的事件序列（PoC-2 记录完整时序表）；DocumentOpen 中访问 Windows 是否安全；Startup 时已打开文档（双击启动 Word）是否触发 DocumentOpen。

**9. 验收标准**
- 10 种打开路径（见 05 矩阵 T-A02-01~10）各执行 3 次，标签数 = `Application.Windows.Count`（+PV 数），零重复。
- Word 运行中启用插件（COM 加载项对话框勾选），现有 5 个文档全部出现。
- 连续 100 次随机打开/关闭后对账报告 `added+removed=0`（即事件路径已完全同步）。

---

## A03 标签点击切换

**元信息**：P0 · 难度 低-中 · 风险【官方】 · 闸口 SWITCH-PASS · 依赖 A13/A14，被 A04 复用

**1. 功能说明**
点击标签激活目标文档窗口；保留目标文档的编辑内容、选区、视图和滚动位置（这些状态本就属于各 Word Window，激活不应改变它们）；不误触其他文档。

**2. 用户流程**
点击标签 → 目标窗口成为活动窗口（WS-0：窗口切到前台；WS-1：目标窗口原地替换可见窗口）→ 所有标签条选中态同步 → 键盘焦点位于目标文档正文。

**3. UI 设计**
- 按下即高亮（乐观更新），若 500ms 内未收到 `WindowActivate` 回执则回滚并显示短暂错误提示（标签内 `fa-triangle-exclamation` 闪烁 + Tooltip “切换失败，已重新同步”）。
- 中键点击 = 关闭（A06）；双击 = 无操作（预留）。

**4. 技术实现**
- `TabService.ActivateTab(tabId)` → `CommandDispatcher.Execute(ActivateTabCommand)` → `IWindowExperienceStrategy.SwitchTo(windowId)`：
  - WS-0：`WordHostAdapter.ActivateWindow(windowId)` → `Word.Window.Activate()`【TLB 已核实】；若窗口最小化，先 `WindowState = wdWindowStateNormal/原状态`。
  - WS-1：见 A04。
- 防循环：`SwitchInProgress` 标志 + `expectedHwnd`；收到 `WindowActivate` 且 hwnd == expectedHwnd → 视为本次切换的回执，只更新选中态，不再发起切换；超时 1s 自动清除标志。
- 焦点：Word OM 的 `Window` 没有独立的 SetFocus 方法，首先依赖 `Activate()` 自身行为；若点击后焦点仍停留在 CTP 控件内，则通过 `SynchronizationContext.Post` 延后一个消息循环，再用 Win32 `SetFocus` 把焦点交给该 Word 窗口的文档区子窗口【待实机：是否需要；子窗口类名属内部实现，只能作为兜底】。

**5. 数据结构**
`ActivationState { activeWindowId, pendingSwitch?{targetWindowId, expectedHwnd, startedAt, correlationId} }`

**6. 模块接口**
`ITabService.ActivateTab` → `ICommandDispatcher` → `IWindowExperienceStrategy.SwitchTo`；回执来自 `IDocumentRegistry.ActiveWindowChanged`。A12、C08 复用同一命令。

**7. 边界场景**
- 目标窗口已被外部关闭（标签过期）：`COMException` → 触发对账、移除标签、提示“文档已关闭”。
- 目标文档正在显示模态对话框（例如另一个窗口的“另存为”）：Word 处于模态状态时 `Activate` 可能失败【待实机】→ 标签条在模态期间显示禁用态（检测方法：`Application` 调用返回 `RPC_E_CALL_REJECTED`/超时，或 Win32 `IsWindowEnabled(mainHwnd)==false`）。
- 用户正在 IME 组字时点击标签：允许，Word 会提交/取消组字（Word 原生行为）。

**8. 技术风险**
【官方】`Window.Activate`。【待实机】从 CTP 点击后键盘焦点是否回到正文；20 个窗口时切换延迟。

**9. 验收标准**
20 个标签随机切换 500 次：选中标签 == `Application.ActiveWindow` 100%；每个文档选区/滚动位置切换前后一致（脚本读取 `Selection.Start/End` 与 `ActivePane.VerticalPercentScrolled`）；未保存内容零丢失；P95 切换延迟 < 300ms（WS-0）。

---

## A04 真正单窗口多文档体验（关键技术闸口）

**元信息**：P0 闸口 · 难度 高 · 风险 WS-1【待实机】/ WS-2【高风险】 · 闸口 **SINGLE-WINDOW-PASS** · 依赖 A01–A03、A06、A13–A16

> [!IMPORTANT]
> **A01 顶部标签栏显示成功 ≠ A03 能切换 ≠ A04 完成。** A04 的完成只能由下方 SW-01~SW-14 全部通过来证明；任何一项失败都记为“A04 未达标”，并由项目负责人决定是否接受降级，**本 Agent 不得自行把 WS-0 窗口切换器重新定义为单窗口**。

**1. 功能说明**
用户主观与客观上都只面对**一个**主要 Word 工作窗口：任务栏只有一个 Word 按钮（对应当前 Word 进程的工作窗口），Alt+Tab 只出现一个 Word 项，顶部标签在此窗口中切换文档，窗口位置/大小/最大化状态在切换中保持不变。不是侧边栏列表，也不是“多窗口切换器”。

**2. 用户流程**
1. 用户在设置中开启「单窗口模式」（默认关闭，直到 PoC-3 通过并经负责人批准）。
2. 插件进入 WS-1：以当前活动窗口为“主视窗”，其余文档窗口被隐藏（任务栏项消失）。
3. 用户打开新文档 → Word 新建的顶层窗口出现 → 插件立即把它对齐到主视窗的位置/尺寸/状态，显示它并隐藏旧窗口（视觉上为“同一窗口内换了文档”）。
4. 点击标签 → 同样的“对齐 → 显示目标 → 隐藏当前”交换。
5. 关闭当前标签 → 文档关闭后，按 MRU 顺序显示下一个文档；最后一个文档关闭时，Word 按原生行为处理（退出或留空白窗口）。
6. 关闭单窗口模式 / 插件卸载 / 插件异常 → `RestoreNativeState()`：所有被 CrabOffice 隐藏的窗口恢复可见。

**3. UI 设计**
- 设置项：「单窗口模式（实验）」开关 + 说明文字 + “遇到问题可按 Ctrl+Alt+Shift+R 恢复所有窗口”（紧急恢复热键，经 A12 机制注册；Ribbon 也提供按钮“显示所有窗口”）。
- 状态指示：标签条最左侧图标 `fa-window-maximize`（单窗口）/ `fa-clone`（原生多窗口）。
- 异常提示：进入失败 → 非模态通知“单窗口模式启动失败，已保持原生多窗口”，附“查看日志”。

**4. 技术实现**

**4.1 研究路线（按文档要求的顺序）**

| 序 | 路线 | 本 Agent 评估 | 状态 |
|---|---|---|---|
| ① | 官方对象模型能否直接提供单主窗口（MDI） | Word 2013 起为 SDI；`Application.ShowWindowsInTaskbar`【TLB 已核实成员仍存在】在 2013+ 是否仍生效**未知**，推测被忽略 | 【待实机】PoC-3 Step 0，10 分钟可证伪 |
| ② | 不破坏生命周期的窗口可见性/激活编排（**WS-1**） | 仅使用官方 `Window.Visible`、`Window.Left/Top/Width/Height/WindowState`、`Window.Activate`；所有文档、窗口仍由 Word 拥有，保存/关闭/对话框走原生路径 | 【待实机】**首选**，PoC-3 |
| ③ | Win32 承载/重设父窗口（**WS-2**） | 自建 Frame + `SetParent(OpusApp 或其子窗口)`；进程内同线程，可避免跨进程 SetParent 的问题，但 Ribbon、Backstage、模态对话框 owner、焦点/加速键、DPI、Office 更新全都是未文档化行为 | 【高风险】仅隔离 VM 实验 PoC-4，需负责人批准 |
| ✗ | DSOFramer / ActiveX 文档宿主 / WebBrowser 承载 Office | 微软已不再支持 Office ActiveX 文档承载，非 Word 主窗口体验，丢失 Ribbon 与插件生态 | 【否决】 |
| ✗ | 全局键盘/窗口 Hook、DLL 注入改写 Word 窗口过程 | 安全软件冲突、Office 更新即失效，违背项目约束 | 【否决】 |
| ✗ | 一个 Word 窗口内“换装”不同文档 | Word 对象模型中 Window 与 Document 绑定，无法让 A 窗口显示 B 文档 | 【否决】 |

**4.2 WS-1 可见性编排详细设计（候选，未验证）**

类：`VisibilityOrchestrationStrategy : IWindowExperienceStrategy`、`FramePlacementTracker`、`HiddenWindowLedger`、`SingleWindowWatchdog`。

- `Enter()`：
  1. 前置检查：无模态对话框；`Application.Windows.Count ≥ 1`；读取主视窗 placement（`GetWindowPlacement(hwnd)` 获取像素级位置 + 显示器，备用 `Window.Left/Top/Width/Height`（单位磅）+ `WindowState`）。
  2. 写 `HiddenWindowLedger`（**先落盘再隐藏**，文件 `%LOCALAPPDATA%\CrabOffice\runtime\hidden-<pid>.json`，含 hwnd、文档 FullName、时间戳）。
  3. 对非主视窗逐个 `wn.Visible = false`【TLB 已核实 `Window.Visible` 可读写；隐藏行为对任务栏/Alt+Tab 的效果待实机】。
- `SwitchTo(target)`：
  1. `Application.ScreenUpdating` 不用于窗口级（它只影响文档绘制）；采用“先显后隐”顺序减少闪烁：`ApplyPlacement(target, current.placement)` → `target.Visible = true` → `target.Activate()` → `current.Visible = false`。
  2. 位置同步优先使用 `SetWindowPlacement`（Win32 公开 API，作用于本进程窗口，【待实机】确认 Word 不会在显示后自行改回）；最大化状态下只需同步“所在显示器 + 最大化”。
  3. 更新 Ledger；更新 `ActivationState`。
- 新窗口出现（DocumentOpen/NewDocument/WindowActivate 中出现未知 hwnd）：`Post` 到下一消息循环 → 同步 placement → 隐藏旧窗口。【待实机】Word 新窗口首帧在默认位置闪现的程度。
- 关闭当前文档：A06 关闭流程完成并对账后 → 选择 MRU 下一个 → `SwitchTo`。若关闭被取消 → 不变。
- 用户通过 Word 原生 “视图 → 切换窗口” 或 Ctrl+F6 / 任务栏缩略图激活隐藏窗口：`WindowActivate` 收到一个 `Visible=false` 的窗口 → 将其视为切换请求，执行 SwitchTo（使其可见并隐藏原窗口）。【待实机】隐藏窗口是否会出现在该菜单中、能否被激活。
- `Exit()` / `RestoreNativeState()`：Ledger 中所有窗口 `Visible = true`（失败则 Win32 `ShowWindow(hwnd, SW_SHOWNA)` 兜底）→ 层叠排列 → 清空 Ledger。
- **崩溃兜底**：
  - 插件内：`AppDomain.UnhandledException`、`Application.ThreadException`、所有 COM 事件处理器 try/catch → 进入 `SafeMode` 并调用 `RestoreNativeState()`。
  - Word 进程仍在而插件被 Office 禁用（Resiliency 禁用）：外部 `CrabOffice.Recovery.exe`（独立进程、开机/插件启动时检查 Ledger）→ 通过 `EnumWindows` 找到 Ledger 中仍存活且属于 WINWORD.EXE 的 hwnd → `ShowWindow(SW_SHOW)`。仅操作 CrabOffice 自己记录的窗口。
  - Word 崩溃：窗口随进程消失，Ledger 失效；下次启动检测残留 Ledger → 清理并提示“上次异常退出，请查看 Word 文档恢复窗格”（不替代 AutoRecover）。
- `SingleWindowWatchdog`：每 1s 检查“可见的 CrabOffice 管理窗口数 == 1”且“存在至少一个可见窗口（当仍有文档时）”，否则自动修复并计数；同一会话修复 >3 次 → 自动退出单窗口模式并提示。

**4.3 WS-2 Win32 宿主框架（仅隔离实验）**
- 自建 `CrabFrameForm`（顶层，任务栏唯一项）→ `SetParent(wordWindowHwnd, frameClientHwnd)` 并移除 `WS_CAPTION/WS_THICKFRAME`、加 `WS_CHILD`。
- 已知或推测问题：Word 窗口失去顶层身份后，Backstage、Ribbon 弹出菜单、模态对话框（以 Word 顶层为 owner）、`Alt` 键提示、Win+方向键贴靠、DWM 阴影、Per-Monitor DPI 消息（`WM_DPICHANGED` 只发顶层）、拖放、IME 候选窗位置均可能异常；`SetParent` 官方文档对跨 DPI 感知上下文的父子关系有限制性说明。
- 逆操作必须完整：保存原 style/exStyle/parent/placement，退出时逐项恢复。
- 仅在 PoC-3 失败**且**负责人批准后进行。

**5. 数据结构**
```text
WindowExperienceState { mode: Native|SingleWindow|SafeMode, strategyId: "WS-0"|"WS-1"|"WS-2", primaryWindowId?, enteredAtUtc, repairCount }
FramePlacement { monitorDeviceName, showCmd, normalRectPx{l,t,r,b}, maxPosPx, dpi, wordWindowState }
HiddenWindowLedger (持久化 JSON) { schemaVersion, wordPid, sessionId, entries[{hwnd, windowId, documentFullName?, hiddenAtUtc, hiddenBy:"CrabOffice"}] }
Win32RestoreRecord (仅 WS-2) { hwnd, originalParent, originalStyle, originalExStyle, originalPlacement }
```

**6. 模块接口**
`IWindowExperienceStrategy { Enter(); SwitchTo(WindowId); OnWindowAppeared(WindowId); OnWindowRemoved(WindowId); Exit(); RestoreNativeState(RestoreReason) }`（完整签名见 03 §5）。`ITabService.ActivateTab` 永远通过策略接口，不关心是 WS-0/1/2。A16 在禁用/异常时调用 `RestoreNativeState`。

**7. 边界场景（必须逐项在 PoC-3 验证）**
| 场景 | 预期处理 |
|---|---|
| 隐藏窗口的文档有未保存内容，用户退出 Word | Word 对每个文档弹保存提示——需验证隐藏文档的提示框是否可见、是否以隐藏窗口为 owner；若不可见 → 在 `Quit`/`DocumentBeforeClose` 中先让该窗口可见 |
| 另存为/打印/查找替换等模态对话框 | 只在可见主视窗发起；对话框期间禁止切换标签 |
| 后台文档触发对话框（如宏、链接更新、修订提示） | 检测到隐藏窗口激活 → 自动切到该窗口 |
| 用户点主视窗右上角 × | 原生关闭当前文档 → 对账后显示 MRU 下一个窗口 |
| 用户最小化主视窗 | 不切换，其余窗口保持隐藏；还原时恢复 |
| 跨显示器拖动 / DPI 不同 | 每次切换前重新读取主视窗 placement；`WM_DPICHANGED` 由各窗口自行处理（隐藏窗口在显示后可能需要一次 DPI 重算）【待实机】 |
| Win+Tab / 任务视图 / 虚拟桌面 | 隐藏窗口不应出现；跨虚拟桌面切换时新显示窗口须在当前桌面【待实机】 |
| Snap/窗口贴靠布局 | 记录 placement 时包含贴靠后的 rect；贴靠组关系无法继承（已知限制） |
| 其他插件隐藏的窗口 | 不纳入编排（hiddenBy≠CrabOffice） |
| 插件被 Office 自动禁用 | Recovery 进程恢复窗口（见 4.2） |
| Word 自动化客户端（如 Python/VBA 脚本）调用 `Documents.Open(Visible:=True)` | 视为新窗口，纳入编排 |

**8. 技术风险**
- WS-1 依赖的 API 全部为官方公开成员，但“Word 是否允许长期存在多个隐藏文档窗口而不影响保存提示、打印、任务栏、Alt+Tab”的**组合行为未经文档承诺**，属【待实机】。最大风险：隐藏文档的模态提示不可见导致 Word 看似卡死；切换闪烁；Word 自身在激活时把隐藏窗口重新显示。
- WS-2 属【高风险】：依赖窗口层级内部结构，任何 Office 更新都可能破坏；不建议进入生产。
- 判定：PoC-3 前，A04 状态为“**未验证**”，不是“可行”。

**9. 验收标准（SINGLE-WINDOW-PASS，≥3 个文档，全部满足）**
| 编号 | 检查项 | 判定方法 |
|---|---|---|
| SW-01 | 任意时刻 CrabOffice 管理的可见 Word 顶层窗口 = 1 | `EnumWindows` 过滤 WINWORD 进程 + `IsWindowVisible` + 类名 `OpusApp` 脚本断言 |
| SW-02 | 任务栏仅 1 个 Word 按钮（且缩略图仅 1 个） | 人工 + 截图 |
| SW-03 | Alt+Tab、Win+Tab 仅 1 个 Word 项 | 人工 + 截图 |
| SW-04 | 切换 3 个文档时窗口位置/尺寸/最大化状态/所在显示器不变（误差 ≤ 2px） | 脚本比较 `GetWindowRect` |
| SW-05 | 切换后键盘直接输入落在目标文档 | 自动化输入 + 读取 `Range.Text` |
| SW-06 | 跨文档复制粘贴（含格式） | 人工 |
| SW-07 | Ctrl+S、另存为、AutoSave（若启用）正确作用于当前文档 | 文件时间戳 + 内容校验 |
| SW-08 | 打印/打印预览（Backstage）作用于当前文档 | 打印到 PDF 校验 |
| SW-09 | 模态对话框（另存为、字体、段落、查找替换、宏安全提示）正常显示在主视窗前且可操作 | 人工 |
| SW-10 | 关闭当前标签（含未保存→取消/不保存/保存三路径）后显示正确窗口 | 脚本 + 人工 |
| SW-11 | 退出 Word：每个隐藏未保存文档的保存提示均可见可操作，无文档丢失 | 人工 |
| SW-12 | 关闭单窗口模式/禁用插件/杀死插件线程（故障注入）后所有窗口恢复可见 | 脚本 |
| SW-13 | 双显示器 + 125%/150% 混合 DPI 下 SW-01~SW-05 通过 | 测试矩阵 |
| SW-14 | 切换闪烁：人眼可见的第二窗口残影 ≤ 1 帧级（以 60fps 录屏逐帧检查，≤ 2 帧可接受，需负责人确认阈值） | 录屏 |

任一项失败 → A04 结论为“**未达标**”，在 06 风险文档中登记并提交负责人决策（D-01）。

---

## A05 新建和打开文件

**元信息**：P0 · 难度 低 · 风险【官方】 · 闸口 SWITCH-PASS · 依赖 A02、CommandDispatcher

**1. 功能说明**：标签条右端“+”按钮快速新建空白文档或打开现有文件；后续支持把文件拖到标签区打开。文件读写完全交给 Word 原生机制。

**2. 用户流程**：单击“+” → 新建空白文档（默认）；点击“+”旁下拉箭头 → 菜单：新建空白 / 基于模板新建 / 打开… / 最近关闭（B01）/ 收藏（B05）。选择“打开…” → Word 原生“打开”对话框 → 选择文件 → A02 生成标签。

**3. UI 设计**：`fa-plus` 按钮 + `fa-caret-down` 下拉；拖放悬停时标签区出现插入指示线与“松开以打开”提示；不支持的扩展名显示 `fa-ban` 光标与提示。

**4. 技术实现**
- 新建：`NewDocumentCommand` → `Application.Documents.Add(Template: Missing, NewTemplate: false, DocumentType: wdNewBlankDocument, Visible: true)`【TLB 已核实参数】。
- 打开：`OpenDocumentCommand` → `Application.Dialogs[WdWordDialog.wdDialogFileOpen].Show()`（原生对话框，保留 Word 的 MRU、受保护视图、格式转换、密码流程）；插件内部按路径打开（B 组）用 `Documents.Open(FileName, ReadOnly, AddToRecentFiles: true, Visible: true)`。
- 拖放（P2）：`TabStripControl.AllowDrop` → `DragEnter` 过滤 `.doc/.docx/.docm/.dot/.dotx/.dotm/.rtf/.odt/.txt` → 逐个 `OpenDocumentCommand`。
- 禁止：直接以文件流写 docx；绕过 Word 打开对话框。

**5. 数据结构**：`OpenRequest { path?, readOnly, source: PlusButton|Drop|Recent|Favorite|Workspace, correlationId }`、`OpenResult { status: Opened|AlreadyOpen|Cancelled|Failed, windowId?, error? }`。

**6. 模块接口**：`ICommandDispatcher.Execute(OpenDocumentCommand)` → `IWordHostAdapter.OpenDocument(OpenRequest)`；结果通过 A02 事件落地，不直接改 TabRegistry。

**7. 边界场景**：已打开文件再次打开 → 判定 `AlreadyOpen`（按规范化路径比对 `Document.FullName`），激活已有标签；只读/被占用文件 → Word 原生提示（只读/通知/取消）；网络路径不可达 → Word 原生报错，`Failed` 记日志；OneDrive/SharePoint URL（`FullName` 为 https）→ 路径比较需特殊规范化【待实机】；受保护视图打开 → 标签显示 PV 状态。

**8. 技术风险**：全部【官方】；`Documents.Open` 对已打开文件的行为（是否弹“还原为保存版本？”取决于 `Revert` 参数）【待实机】。

**9. 验收标准**：新建 10 次标签 +10 且均为“文档N”未命名态；打开只读文件标签带 `fa-lock`；重复打开不新增标签而是激活；取消对话框无副作用；拖放 3 个文件生成 3 个标签（P2）。

---

## A06 标签关闭与保存保护

**元信息**：P0 · 难度 中 · 风险【官方】+【待实机】 · 闸口 SWITCH-PASS · 依赖 A13/A14，被 A11/B04 复用

**1. 功能说明**：单击标签 × 或中键关闭对应窗口/文档；有未保存修改时走 Word 原生“保存/不保存/取消”流程；取消 → 标签保留；保存失败 → 数据不丢失。

**2. 用户流程**：点 × → （若该文档仅此一个窗口且未保存）Word 原生保存提示 → 保存：保存后关闭，标签移除；不保存：关闭，标签移除；取消：文档、标签原样保留。

**3. UI 设计**：× 在悬停或选中时显示（固定标签不显示 ×，需先取消固定或右键关闭）；关闭进行中标签半透明 + 禁用重复点击；失败时 Tooltip 说明。

**4. 技术实现**
- `CloseTabCommand(tabId)`：
  1. 解析 `WindowRecord`；读取 `doc.Windows.Count`。
  2. 若 >1：`wn.Close()`（只关闭该窗口，文档保留）【官方：Window.Close】。
  3. 若 ==1：`doc.Close(SaveChanges: WdSaveOptions.wdPromptToSaveChanges)`。
  4. 用户取消时 Word 会抛出 `COMException`【待实机：记录确切 HRESULT，推测为“命令失败”类错误】→ 捕获为 `CloseResult.Cancelled`，不视为错误。
  5. **不在 `DocumentBeforeClose` 中删除标签**（该事件之后用户仍可能取消）；命令返回后 + `Post` 一次对账，以 `Application.Windows` 真实结果决定移除。
- 用户通过 Word 原生 × / Ctrl+W / 文件→关闭关闭：`DocumentBeforeClose` 只标记 `closingCandidate=true`；后续 `WindowActivate/DocumentChange` 或 300ms 延时对账确认。
- WS-1 下：关闭完成后由策略显示 MRU 下一个窗口（A04）。

**5. 数据结构**：`CloseRequest { tabIds[], mode: Single|Others|Right|All, respectPinned=true }`、`CloseResult { perTab[{tabId, status: Closed|Cancelled|Failed|Skipped, error?}] }`、`DocumentRecord.closingCandidateSince?`。

**6. 模块接口**：`ITabService.CloseTab` → `ICommandDispatcher` → `IWordHostAdapter.CloseTarget(windowId, CloseOptions)`；成功关闭后发布 `DocumentClosed`（B01 最近关闭订阅）。

**7. 边界场景**：同一文档多窗口；只读文档修改后关闭（Word 提示另存为）；保存时磁盘满/无权限 → Word 报错后文档仍打开，标签保留；AutoSave 打开的云文档 → 通常无提示直接关闭；文档正在后台保存（`IsInAutosave`/后台保存）→ 等待；关闭最后一个文档时 Word 可能退出 → 插件随 `Quit` 清理。

**8. 技术风险**：【官方】Close API 与 `DocumentBeforeClose(Doc, ref Cancel)`；【待实机】取消时异常类型、CTP 中触发关闭时模态对话框 owner 是否正确。

**9. 验收标准**：未保存 × → 取消：标签、内容、选区不变；不保存：标签消失且文件未改；保存：文件更新后标签消失；保存失败（只读目录）：文档与标签保留；多窗口文档关闭一个窗口后文档仍在。

---

## A07 修改未保存标记

**元信息**：P0 · 难度 低-中 · 风险【官方】+【待实机】 · 闸口 UI-PASS · 依赖 A13/A14

**1. 功能说明**：有未保存编辑时标签显示 `●`（`fa-circle` 小号），保存后消失；新建未命名文档显示斜体名称“文档1”；AutoSave 云文档显示云状态而非 ●。

**2. 用户流程**：编辑 → ≤1s 出现 ●；Ctrl+S → ≤1s 消失；撤销到保存点 → 与 Word 是否仍提示保存保持一致（以 `Saved` 为准）。

**3. UI 设计**：● 位于关闭按钮位置，悬停时变 ×（同浏览器习惯）；云 AutoSave：`fa-cloud`（已同步）/`fa-cloud-arrow-up`（同步中）。

**4. 技术实现**
- 只读 `Document.Saved`【官方】；**绝不写入** `Saved = true`。
- 刷新时机：`WindowSelectionChange`（节流 500ms，仅活动文档）、`DocumentBeforeSave` + 保存后 1s 复查（Word 无 AfterSave 事件【TLB 已核实】）、`WindowActivate`、低频轮询（活动文档 1s、非活动文档 5s，单次读取 <1ms 级【待实机】）。
- AutoSave：`Document.AutoSaveOn`【MSWORD.OLB 已核实存在，PIA 15 无此成员 → 需 `dynamic` 晚绑定并 try/catch】、`IsInAutosave`【TLB 已核实】。

**5. 数据结构**：`DocumentRecord.savedState: Saved|Dirty|Unknown`、`autoSaveOn?`、`isUntitled`（`Path == ""`）、`lastSavedProbeUtc`。

**6. 模块接口**：`IDocumentRegistry.UpdateSavedState(documentId)` → `TabsChanged`；B04、A06、B03 读取。

**7. 边界场景**：只更改了文档属性/字段更新导致 Saved=false（Word 行为一致即可）；受保护视图文档恒为 Saved；只读文档编辑后 ●；文档被脚本修改未触发选区变化 → 由轮询兜底。

**8. 技术风险**：【官方】`Saved`；【待实机】AutoSave 开启时 `Saved` 的变化频率与表现；轮询对大文档性能影响。

**9. 验收标准**：编辑/保存/另存为/撤销/重做/AutoSave 6 种操作，● 状态与“关闭时 Word 是否提示保存”100% 一致；50 文档时 CPU 占用（空闲）< 1%。

---

## A08 标签拖拽排序

**元信息**：P1 · 难度 中 · 风险 低 · 闸口 UI-PASS · 依赖 A01、B03

**1. 功能说明**：拖动标签改变顺序；不改变文件内容和 Word `Documents` 集合顺序。
**2. 用户流程**：按住标签水平拖动 ≥5px 进入拖拽 → 其他标签让位动画 → 松开落位 → 所有窗口标签条同步。Esc 取消拖拽。
**3. UI 设计**：拖动中标签半透明跟随鼠标，插入点竖线指示；固定区与非固定区之间不可互拖（拖入固定区 = 询问是否固定，P2）。
**4. 技术实现**：`TabStripControl` 处理 `MouseDown/Move/Up` + `Capture`；计算目标索引 → `ITabService.MoveTab(tabId, newIndex)` → `TabRegistry` 重排 `order`（稀疏整数或重新编号）→ `TabsChanged` → 防抖 2s 写入会话（B03）。不调用任何 Word API。
**5. 数据结构**：`TabRecord.order:int`、`DragState { tabId, startX, currentIndex, targetIndex }`。
**6. 模块接口**：`ITabService.MoveTab`；`ISessionStore.SaveSessionDebounced()`。
**7. 边界场景**：拖拽过程中有文档关闭 → 取消拖拽并重算；拖出标签条（未来“拆分到新窗口”）→ 当前版本无效果；WS-1 下拖拽不触发切换。
**8. 技术风险**：低。WinForms 自绘拖拽在高 DPI 下坐标换算需统一（A15）。
**9. 验收标准**：拖拽 20 次后顺序与预期一致；再次点击各标签激活正确文档；重启 + 恢复会话后顺序保持（以文件路径映射）。

---

## A09 固定标签

**元信息**：P1 · 难度 低 · 风险 低 · 闸口 UI-PASS · 依赖 A08、A11、B02/B03
1. **功能说明**：重要文档固定在最左侧，紧凑显示（仅图标 + 短名），不被“关闭其他/关闭右侧/关闭全部”误关。
2. **用户流程**：右键 → 固定 → 标签移至固定区；再次右键 → 取消固定。
3. **UI 设计**：固定区与普通区之间 1px 分隔线；`fa-thumbtack` 角标；固定标签无 ×。
4. **技术实现**：`ITabService.PinTab(tabId, bool)` → `TabRecord.isPinned`，固定标签 order 始终小于非固定标签；批量关闭命令在 `CommandDispatcher` 中按 `respectPinned` 过滤。持久化 key 为规范化路径（未命名文档的固定状态仅本次会话有效）。
5. **数据结构**：`TabRecord.isPinned`、`SessionTab.isPinned`。
6. **模块接口**：A11、B02、B04 读取 `isPinned`。
7. **边界场景**：固定文档被 Word 原生 × 关闭 → 正常关闭（固定只防插件批量命令）；固定的未命名文档保存后转为可持久化。
8. **技术风险**：低。
9. **验收标准**：固定 3 个后“关闭其他”只关闭非固定；重启恢复固定状态一致；取消固定后回到普通区首位。

---

## A10 标签过多、搜索和滚动

**元信息**：P1 · 难度 中 · 风险 低 · 闸口 UI-PASS · 依赖 A01、A13
1. **功能说明**：10–50 个文档时仍能快速定位：标签压缩、横向滚动、溢出菜单、模糊搜索；同名文件显示区分路径。
2. **用户流程**：标签宽度压缩到最小宽度（约 80 DIP）后出现左右滚动按钮与 `fa-chevron-down` 溢出列表；点击 `fa-magnifying-glass` 或快捷键（默认 Ctrl+Shift+A，可配置，需冲突评估）→ 弹出搜索框，输入即过滤，回车激活。
3. **UI 设计**：同名文件显示 `报告.docx — D:\项目A` 的最短可区分父目录后缀；溢出列表按标签顺序，带未保存/固定标记；活动标签被激活时自动滚动到可见。
4. **技术实现**：布局算法 `TabLayoutEngine`（宽度 = clamp(可用宽度/n, 80, 220) DIP）；只绘制可见区域（自绘控件天然虚拟化）；搜索为内存模糊匹配（标题 + 完整路径，拼音首字母匹配为 P2 可选）；同名消歧 `DisplayLabelResolver` 计算最短不同路径后缀。
5. **数据结构**：`TabRecord.displayLabel`、`disambiguationSuffix`；`SearchState { query, results[], selectedIndex }`。
6. **模块接口**：`ITabService.GetTabs()`、`SearchTabs(query)`；C08 命令面板复用搜索组件。
7. **边界场景**：未命名文档无路径 → 用“文档N”；50 个标签 + 窄窗口 → 只显示活动标签与溢出；搜索框打开时 Word 正文快捷键不受影响（失焦即关闭）。
8. **技术风险**：低；搜索弹出层需是 CTP 内控件或 owned 窗体，焦点归还 Word【待实机】。
9. **验收标准**：50 个文档时滚动、溢出、搜索均可用，重绘 < 16ms；10 个同名 `报告.docx` 全部可区分；搜索结果激活正确文档。

---

## A11 右键菜单

**元信息**：P1 · 难度 低-中 · 风险【官方】 · 闸口 SWITCH-PASS · 依赖 A06、A09、B02、CommandDispatcher
1. **功能说明**：标签右键提供：保存、另存为、关闭、关闭其他、关闭右侧、复制路径、打开所在目录、固定/取消固定、加入工作区。
2. **用户流程**：右键标签 → 菜单（不改变活动标签）→ 选择命令 → 只作用于被右键的标签（批量命令作用于计算出的集合）。
3. **UI 设计**：`ContextMenuStrip`，图标（FontAwesome）：保存 `fa-floppy-disk`、另存为 `fa-file-export`、关闭 `fa-xmark`、复制路径 `fa-copy`、打开目录 `fa-folder-open`、固定 `fa-thumbtack`、加入工作区 `fa-layer-group`。不可用项置灰并在 Tooltip 说明（如未命名文档“尚未保存到磁盘，无所在目录”）。
4. **技术实现**：每项 = 一个 `ICommand`，`CanExecute(ctx)` 统一校验（路径存在、只读、PV、未命名、模态状态）；保存 → `doc.Save()`（未命名时 Word 自动弹另存为）；另存为 → 先激活目标窗口再 `Dialogs[wdDialogFileSaveAs].Show()`（对话框需作用于该文档，**必须先激活**）；关闭其他/右侧 → 计算集合 → 逐个执行 A06 流程，遇“取消”立即停止后续并汇报；复制路径 → `Clipboard.SetText(FullName)`；打开目录 → `explorer.exe /select,"<path>"`（`Process.Start`，路径加引号并校验存在）。
5. **数据结构**：`CommandContext { targetTabIds[], invokedFrom: TabContextMenu, activeTabId }`、`BatchCommandResult`。
6. **模块接口**：`ICommandDispatcher.CanExecute/Execute`；`IWorkspaceService.AddToWorkspace`。
7. **边界场景**：右键时目标标签已失效 → 菜单关闭并对账；批量关闭中途用户取消 → 已关闭的不回滚，剩余保留；PV 文档仅允许关闭/复制路径。
8. **技术风险**：【官方】为主；另存为对话框作用对象依赖活动文档，需先激活（WS-1 下即切换）。
9. **验收标准**：对非活动标签执行每个菜单项，仅影响该标签；未命名文档“打开所在目录”置灰并有说明；“关闭其他”跳过固定标签。

---

## A12 键盘快捷切换

**元信息**：P1 · 难度 中 · 风险【待实机】 · 闸口 SWITCH-PASS · 依赖 CommandDispatcher、设置
1. **功能说明**：按标签顺序前后切换、直接跳到第 N 个标签、打开搜索；可在设置中启用/关闭和改键；不破坏 Word 已有快捷键与输入法。
2. **用户流程**：按快捷键 → 切换到相邻标签（WS-1 下原地切换）；设置页可启用/禁用每个绑定并检测冲突。
3. **UI 设计**：设置页表格（命令 / 默认键 / 当前键 / 冲突提示 `fa-triangle-exclamation`）；首次启用 Ctrl+Tab 时提示“表格中 Ctrl+Tab 用于插入制表符”。
4. **技术实现（两条候选，PoC-5 选择）**
   - **方案 K1（官方，首选）**：`Application.KeyBindings.Add(wdKeyCategoryMacro, "CrabOffice_NextTab", BuildKeyCode(...))`【TLB 已核实 KeyBindings/CustomizationContext】；先设 `Application.CustomizationContext = CrabOffice.dotm`（放在 Word 启动文件夹的全局模板）避免污染 Normal.dotm；模板内宏仅一行：`Application.COMAddIns("CrabOffice").Object.Invoke "NextTab"`，加载项通过 `RequestComAddInAutomationService` 暴露 `[ComVisible]` 接口【官方】。代价：需要分发一个带宏的模板（受信任位置/宏安全策略影响），企业禁宏环境不可用。
   - **方案 K2（备选）**：线程级键盘钩子 `SetWindowsHookEx(WH_KEYBOARD, proc, IntPtr.Zero, GetCurrentThreadId())`（仅 Word UI 线程，非全局）；需判断 IME 组字状态（`ImmGetContext/ImmGetCompositionString` 或在 `WM_IME_*` 期间不处理）和 `Selection.Information(wdWithInTable)`；属常见 VSTO 做法但非 Office 官方扩展点 → 【待实机】【中风险】。
   - 默认键位建议（待负责人确认 D-05）：下一个 `Ctrl+Alt+PageDown`、上一个 `Ctrl+Alt+PageUp`（Word 原生 Ctrl+PgDn/PgUp 为浏览对象，加 Alt 后冲突需在 PoC 检测）；Ctrl+Tab/Ctrl+Shift+Tab 作为可选项默认关闭；Ctrl+Alt+1..9 跳转第 N 个；紧急恢复 `Ctrl+Alt+Shift+R`。
5. **数据结构**：`ShortcutBinding { commandId, keys, enabled, mechanism: KeyBindings|ThreadHook, conflictWith? }`。
6. **模块接口**：`IShortcutService.Register/Unregister`；触发后 → `ICommandDispatcher.Execute(NextTabCommand)`。
7. **边界场景**：中文输入法组字中按键（百度输入法 Word 插件已安装于本机，需纳入测试）；模态对话框打开时不响应；Word 原生 Ctrl+F6（下一个窗口）在 WS-1 下会激活隐藏窗口 → 由 A04 的“隐藏窗口被激活”处理；插件卸载时必须移除 KeyBindings（K1）或卸载钩子（K2）。
8. **技术风险**：K1【官方】但依赖宏模板分发；K2【待实机】。均不使用全局 Hook。
9. **验收标准**：在正文、表格、页眉页脚、IME 组字、模态对话框 5 种状态下，快捷键行为符合预期且 Word 原有功能未受影响；禁用后 Word 快捷键恢复原状；卸载后 Normal.dotm 无残留绑定。

---

## A13 文档和窗口身份模型

**元信息**：P0 基础 · 难度 中-高 · 风险【待实机】 · 闸口 SWITCH-PASS · 被所有 A/B 功能依赖
1. **功能说明**：可靠识别同名文档、同一文件的多个窗口、未保存文档、受保护视图文档、多 Word 进程，确保命令永远作用于正确目标。
2. **用户流程**：无直接交互；表现为“永不串标签”。
3. **UI 设计**：同名消歧（A10）；同文档多窗口标签显示 `名称:2`（与 Word 标题一致，`Window.Caption`）。
4. **技术实现**
   - `DocumentId`、`WindowId`、`TabId` 均为会话内 GUID（`Guid.NewGuid()`），**互不相同、不持久化为跨会话主键**。
   - 运行期映射：`WindowRecord ↔ Word.Window RCW`、`hwnd`（`Window.Hwnd`【TLB 已核实】，仅窗口存活期有效）；`DocumentRecord ↔ Word.Document RCW`。
   - 匹配顺序：①hwnd（同进程同时刻唯一）②RCW 引用相等（.NET 对同一 COM 身份通常返回同一 RCW，Word 是否每次返回同一 IUnknown【待实机】）③`Document.DocID`（MSWORD.OLB 中存在，语义与稳定性【待查证】）④`FullName` + `Windows.Count` 兜底。
   - **禁止**写入 `Document.Variables` / 自定义属性 / 内容控件作为身份标记（会修改文档并改变 `Saved`）。
   - 跨会话（B02/B03）映射：使用 `CanonicalPath`（`Path.GetFullPath` + 大小写不敏感 + 解析映射盘/UNC；云文档保留 URL）作为“文件身份”，不是 DocumentId。
   - 多进程：Registry 只管理本进程；`WindowRecord.wordProcessId = Process.GetCurrentProcess().Id`。
5. **数据结构**：见 03 §4（DocumentRecord/WindowRecord/TabRecord 完整字段）。
6. **模块接口**：`IDocumentRegistry.TryResolve(Word.Window) → WindowId?`、`ResolveDocument(Word.Document) → DocumentId?`、`GetComWindow(WindowId)`（仅 UI 线程，返回运行期引用，不得缓存到其他对象）。
7. **边界场景**：另存为改名（FullName 变化，DocumentId 不变，标签标题更新）；PV → 编辑模式（新窗口，迁移 Tab 位置）；文档关闭后同路径重开（新 DocumentId）；“新建窗口”后关闭原窗口（Window 编号重排，`Caption` 变化）。
8. **技术风险**：【待实机】RCW 身份稳定性、hwnd 在“窗口重建”（如切换阅读模式）时是否变化。
9. **验收标准**：10 个同名不同路径文件 + 同一文件 3 个窗口 + 5 个未命名文档 + 1 个 PV 文档，执行 200 次随机切换/关闭/另存为，命令目标 0 错误。

---

## A14 事件同步与容错

**元信息**：P0 基础 · 难度 高 · 风险【待实机】 · 闸口 SWITCH-PASS · 依赖 A13
1. **功能说明**：在高频打开/激活/保存/关闭中，状态不重复注册、不死循环、不崩溃、不在后台线程碰 COM。
2. **用户流程**：无交互。
3. **UI 设计**：开发者诊断面板（隐藏，Ctrl+Alt+Shift+D 或设置中开启）显示事件流与对账统计。
4. **技术实现**
   - `WordEventAdapter`：每个 COM 事件处理器只做“捕获参数 → 生成 `HostEvent` → 入队”，全部包裹 try/catch，**绝不让异常回流到 Word**。
   - `HostEventQueue`：UI 线程 `Post` 批处理（同一消息循环内合并），去重键 `(kind, hwnd)`；单调递增 `seq`。
   - 重入保护：`RegistryMutationScope`（布尔标志 + 队列，处理期间新事件延后）；`SwitchInProgress`（A03）。
   - 线程模型：所有 Office OM 调用在 Word 主 STA 线程【官方：Office OM 非线程安全，应在主线程调用】；后台任务（SQLite、AI HTTP、文件存在性检查）用 `Task.Run`，结果 `Post` 回 UI 线程。
   - 对账 `Reconciler`：触发见 A02；比较算法 O(n)；发现差异写 `ReconcileReport` 并计数（遥测仅本地日志）。
   - COM 引用：RCW 只保存在 `ComReferenceTable`（UI 线程独占），对象失效（访问抛 `COMException`/`InvalidComObjectException`）即移除；不对 Word 对象滥用 `Marshal.ReleaseComObject`（遵循 VSTO 惯例，仅对自己循环中大量临时对象释放）。
5. **数据结构**：`HostEvent`、`ReconcileReport`、`EventStats { received, coalesced, dropped, handlerErrors }`。
6. **模块接口**：`IHostEventSink.Enqueue(HostEvent)`、`IDocumentRegistry.ApplyHostEvent/Reconcile`、`RegistryChanged` 事件（携带 diff）。
7. **边界场景**：Word 正忙（打印/保存大文档）时 COM 调用被拒（进程内一般不会出现 `RPC_E_CALL_REJECTED`，但模态状态下部分方法失败【待实机】）→ 指数退避重试 3 次；`Quit` 事件后停止所有处理；同一事件重复投递。
8. **技术风险**：【待实机】事件顺序；DocumentOpen 期间访问对象的安全性。
9. **验收标准**：脚本 5 分钟内批量打开/关闭 200 次 + 交替激活，结束时 Registry 与 Word 枚举完全一致；0 个未处理异常；Word 不出现“加载项导致缓慢/禁用”提示。

---

## A15 窗口大小、多显示器与主题

**元信息**：P1 · 难度 中-高 · 风险【待实机】 · 闸口 UI-PASS（WS-1 时影响 SINGLE-WINDOW-PASS）· 依赖 A01、A04
1. **功能说明**：100/125/150/200% DPI、双屏混合 DPI、最大化/最小化/贴靠、Office 深浅主题下标签栏与单窗口模式表现正确。
2. **用户流程**：用户拖动 Word 跨屏 / 改系统缩放 / 改 Office 主题 → 标签栏自动适配，无需重启。
3. **UI 设计**：所有尺寸以 DIP 定义；图标使用 FontAwesome 字体（矢量，随 DPI 缩放）；主题调色板：Colorful/Dark Gray/Black/White/跟随系统。
4. **技术实现**：读取 Office 的 DPI 模式（“针对最佳外观优化/兼容性优化”）【官方：Office 高 DPI 文档】；`TabStripControl` 在 `OnDpiChangedAfterParent`（.NET 4.7+ WinForms 支持，需 app.config 开启 PerMonitorV2【待实机：VSTO 宿主中 app.config 设置是否生效，因为宿主进程是 WINWORD.EXE】）或手动 `GetDpiForWindow(Handle)` 计算缩放；主题变化通过 Ribbon 回调/定时检测注册表值变化。WS-1 下 placement 以物理像素 + 显示器设备名保存。
5. **数据结构**：`DpiContext { dpi, scale, monitor }`、`ThemePalette`。
6. **模块接口**：`IThemeService`、`IDpiService`；A04 `FramePlacementTracker` 依赖 `IDpiService`。
7. **边界场景**：显示器拔出（窗口位置越界 → 移回主屏）；远程桌面会话 DPI 变化；高对比度模式（使用系统色）。
8. **技术风险**：【待实机】VSTO 中 WinForms 的 DPI 感知上下文由 Office 决定；ElementHost(WPF) 在混合 DPI 下历史问题较多 → 首选纯 WinForms 自绘。
9. **验收标准**：按 05 矩阵 DPI×显示器×窗口状态组合，标签无模糊/裁切/错位，WS-1 切换无跨屏跳动。

---

## A16 禁用、卸载与故障恢复

**元信息**：P0 基础 · 难度 中-高 · 风险【待实机】 · 闸口 ALL · 依赖 A04、A12
1. **功能说明**：关闭插件、Office 自动禁用插件、插件异常、卸载后，Word 完全恢复正常：无隐藏窗口、无残留快捷键、无无法编辑状态。
2. **用户流程**：用户在 COM 加载项中取消勾选 / 卸载程序 / 插件崩溃 → 所有文档窗口可见，Word 正常使用；下次启动若检测到上次异常，提示并提供“安全模式启动 CrabOffice（禁用单窗口与快捷键）”。
3. **UI 设计**：安全模式下标签栏显示 `fa-life-ring` 状态；通知“CrabOffice 上次异常退出，已恢复所有窗口”。
4. **技术实现**
   - `ThisAddIn_Shutdown`：`strategy.RestoreNativeState(Shutdown)` → 移除 KeyBindings/钩子 → 移除所有 CTP → 保存会话 → 释放资源。注意 Office 2013+ 快速关闭时 `Shutdown` 可能不被调用【官方：VSTO 文档说明 Office 不再保证调用 Shutdown 事件，推荐在 `DocumentBeforeClose`/`Quit` 中做清理】。
   - 全局异常：`AppDomain.CurrentDomain.UnhandledException`、`Application.ThreadException`、`TaskScheduler.UnobservedTaskException` → 记录 → `SafeMode`。
   - 启动性能：插件加载耗时须远低于 Office 自动禁用阈值（Office 会因加载缓慢禁用 COM 加载项【官方】）→ 启动只做最小初始化，其余延后。
   - 卸载：MSI/ClickOnce 卸载移除注册表 `HKCU\Software\Microsoft\Office\Word\Addins\CrabOffice.Word`、全局模板 `CrabOffice.dotm`、`%LOCALAPPDATA%\CrabOffice\runtime`；用户数据（工作区/收藏）询问是否保留。
   - `CrabOffice.Recovery.exe`：见 A04 4.2。
   - 日志：`%LOCALAPPDATA%\CrabOffice\logs\*.log`，滚动 7 天，不记录文档正文。
5. **数据结构**：`CrashMarker { pid, startedAtUtc, cleanExit=false }`（启动写入，正常退出置 true）、`HiddenWindowLedger`、`SafeModeState`。
6. **模块接口**：`ILifecycleService.Startup/Shutdown/EnterSafeMode`；所有带状态修改的模块实现 `IRestorable.RestoreNativeState()`。
7. **边界场景**：Word 被任务管理器结束；插件 DLL 被更新替换；用户手动删除模板；多进程各自的 Ledger。
8. **技术风险**：【官方】Shutdown 不保证调用、Office 禁用缓慢插件；【待实机】Recovery 进程对 Word 窗口 `ShowWindow` 的效果。
9. **验收标准**：故障注入（抛异常、`Environment.FailFast` 模拟、禁用加载项、杀插件后台线程）10 种 × 3 次，所有原隐藏窗口恢复可见、文档零丢失、卸载后注册表/Normal.dotm 无残留。

---

# B 组：会话、工作区和文档管理（第二优先级）

## B01 最近关闭

**元信息**：P2 · 难度 低 · 风险 低 · 依赖 A06、A13、存储层
1. **功能说明**：记录确认已关闭的文档，便于一键重新打开（类似浏览器 Ctrl+Shift+T）。
2. **用户流程**：“+”下拉 → 最近关闭（最多 20 条）→ 点击 → 重新打开并恢复标签位置（尽量）。
3. **UI 设计**：列表项：图标、名称、路径（灰色）、关闭时间；无持久文件的项显示 `fa-circle-exclamation` “未保存的新文档无法恢复，请查看 Word 文档恢复”且不可点。
4. **技术实现**：订阅 `DocumentClosed`（**对账确认后**，非 BeforeClose）→ 记录 `{canonicalPath, displayName, closedAtUtc, lastTabOrder, wasPinned}` → SQLite `recent_closed` 表；重开 → `OpenDocumentCommand`，打开前 `File.Exists` 检查（后台线程）。
5. **数据结构**：`RecentClosedEntry { id, canonicalPath?, displayName, closedAtUtc, tabOrder, wasPinned, recoverable }`。
6. **模块接口**：`IWorkspaceService.ListRecentClosed(limit)`、`ReopenRecentClosed(id)`。
7. **边界场景**：文件已删除/移动/无权限 → 提示并提供“从列表移除”；云文档 URL → 交给 Word 打开；同一路径重复关闭 → 合并为一条最新。
8. **技术风险**：低。
9. **验收标准**：已保存文件 100% 可重开；缺失路径提示清晰；未命名文档标注不可恢复。

## B02 工作区保存/恢复

**元信息**：P2 · 难度 中 · 风险 低 · 依赖 A08/A09/A13、存储层
1. **功能说明**：将一组文档（含顺序、固定、活动项、分组）保存为命名工作区（如“论文”“教学”），一键恢复。
2. **用户流程**：Ribbon/标签菜单 → 工作区 → 保存当前为… → 输入名称；恢复：选择工作区 → 逐项打开 → 结果面板（成功/已打开/缺失/失败）。
3. **UI 设计**：工作区面板（侧边 CTP 或对话框）：列表、重命名、删除、导出/导入 JSON；恢复进度条与失败清单。
4. **技术实现**：`WorkspaceService.SaveWorkspace(name)` 从 TabRegistry 生成 `Workspace`（未命名文档不能纳入，提示先保存）；`LoadWorkspace(id, mode: Merge|ReplaceSaved)`：逐项 `OpenDocumentCommand`（UI 线程串行，每项间 `Post` 让出消息循环），已打开的跳过；Replace 模式只关闭**已保存**且不在工作区内的文档，**绝不丢弃未保存内容**（未保存的保留并提示）。存储：SQLite（`workspaces`、`workspace_items` 表，`schema_version` 迁移）。
5. **数据结构**：`Workspace { workspaceId(GUID 持久), name, items[{canonicalPath, order, isPinned, groupId?}], activePath?, createdAtUtc, updatedAtUtc, schemaVersion }`。
6. **模块接口**：`IWorkspaceService.SaveWorkspace/LoadWorkspace/ListWorkspaces/DeleteWorkspace/AddToWorkspace`；`LoadWorkspace` 返回 `WorkspaceLoadReport`。
7. **边界场景**：文件缺失；同一文件已在其他 Word 进程打开（Word 只读提示）；工作区含 50 个文件（逐项打开耗时，允许取消）；路径为映射盘但当前未映射。
8. **技术风险**：低；批量打开触发大量事件 → 依赖 A14。
9. **验收标准**：保存→退出 Word→重启→恢复，集合/顺序/固定/活动项一致；缺失文件单独列出；未保存文档不受影响。

## B03 启动时恢复上次会话

**元信息**：P2 · 难度 低-中 · 风险 低 · 依赖 B02、A16
1. **功能说明**：可选在 Word 启动时恢复上一会话的文档集合。
2. **用户流程**：设置“启动时恢复上次会话：关闭（默认）/询问/自动”；Word 启动 → 询问通知“恢复上次 8 个文档？”→ 恢复。
3. **UI 设计**：非模态通知条（标签栏下方或 CTP 内），不阻塞 Word 启动。
4. **技术实现**：会话 = 特殊工作区 `__last_session__`；每次 RegistryChanged 防抖 5s 写入；`Quit` 时写最终版并标记 cleanExit；启动时若 Word 已因双击打开某文件，恢复时跳过已打开项；异常退出（CrashMarker）只恢复磁盘存在的路径，提示用户“未保存内容请使用 Word 文档恢复窗格”。
5. **数据结构**：`SessionSnapshot { savedAtUtc, cleanExit, items[...] }`。
6. **模块接口**：`ISessionStore.SaveSnapshot/LoadLastSnapshot`；复用 `IWorkspaceService.LoadWorkspace`。
7. **边界场景**：Word 以自动化方式启动（`/automation`、嵌入编辑）→ 不恢复（检测 `Application.Visible == false` 或无用户交互窗口）；多进程 → 只由首个进程恢复（命名 Mutex）。
8. **技术风险**：低；启动时机需在 Word 完成初始化后（`Post` + 首个 `WindowActivate`）。
9. **验收标准**：开关有效；不产生重复文档；自动化启动不恢复。

## B04 保存全部/关闭全部

**元信息**：P2 · 难度 中 · 风险 低 · 依赖 A06、A07、A09
1. **功能说明**：一键保存/关闭当前进程（或当前工作区）的所有文档，保留用户交互，逐项反馈。
2. **用户流程**：Ribbon/右键/命令面板 → 保存全部 → 对每个 Dirty 文档 `Save()`（未命名弹另存为）→ 结果汇总；关闭全部 → 逐个 A06 流程，用户在某项选“取消”→ 询问“停止剩余操作？”。
3. **UI 设计**：进度通知 + 结果面板（成功 `fa-check`、跳过 `fa-forward`、失败 `fa-xmark`、取消 `fa-ban`）。
4. **技术实现**：`SaveAllCommand`：遍历 Dirty（A07）→ 逐个 `doc.Save()`，try/catch 记录；不使用 `Documents.Save(NoPrompt:=True)`（会静默处理，不利于逐项反馈）；`CloseAllCommand`：按标签顺序（非固定优先，固定按选项）逐个关闭。WS-1 下需要用户交互的文档先切换为可见。
5. **数据结构**：`BatchCommandResult { items[{tabId, action, status, message}] , stoppedByUser }`。
6. **模块接口**：`ICommandDispatcher.Execute(SaveAllCommand/CloseAllCommand)`。
7. **边界场景**：只读文档（另存为提示）；网络中断；兼容模式文档（.doc）保存格式提示；用户在中途切换标签（批处理期间锁定标签条）。
8. **技术风险**：低。
9. **验收标准**：无任何静默丢弃；取消后剩余文档原样保留；结果报告与实际一致。

## B05 收藏和快速打开

**元信息**：P2 · 难度 低 · 风险 低 · 依赖 A05、存储层
1. **功能说明**：收藏常用文件，从“+”菜单/命令面板快速打开。
2. **用户流程**：右键标签 → 加入收藏；“+”下拉 → 收藏 → 点击打开；管理页可排序/删除/重命名显示名。
3. **UI 设计**：`fa-star`；失效项显示 `fa-link-slash` 与“文件不存在或无权限”。
4. **技术实现**：SQLite `favorites` 表；后台线程定期（打开菜单时）检查存在性；只存元数据，**不上传、不缓存正文**。
5. **数据结构**：`Favorite { id, canonicalPath, displayName, addedAtUtc, lastOpenedAtUtc, sortOrder }`。
6. **模块接口**：`IFavoritesService.Add/Remove/List/Open`。
7. **边界场景**：重复收藏合并；网络盘离线时检查超时（2s）不阻塞 UI。
8. **技术风险**：低。
9. **验收标准**：收藏、打开、失效提示正确；数据库损坏时自动备份并重建，不影响 Word。

---

# C 组：Word 效率增强（后续模块，已纳入架构）

> C 组通用约束：所有修改通过 Word 对象模型执行；每个“应用”动作包裹在 `Application.UndoRecord.StartCustomRecord("CrabOffice: <操作名>") … EndCustomRecord()`【TLB 已核实 `UndoRecord`】中，使用户可一次 Ctrl+Z 撤销；三步式“检查 → 预览 → 应用”；**禁止直接修改 DOCX ZIP 内部 XML**。长操作期间可设 `Application.ScreenUpdating=false` 并在 finally 恢复。

## C01 多项查找替换

**元信息**：P3 · 难度 中 · 风险【官方】 · 依赖 CommandDispatcher、UndoRecord
1. **功能说明**：一次导入多条查找/替换规则（含通配符、区分大小写、全字匹配、格式），预览所有命中后勾选应用。
2. **用户流程**：打开“批量替换”侧边面板 → 新建规则或导入 CSV/JSON → 选择范围（全文/选区/正文不含页眉页脚）→ 预览（列出命中上下文）→ 勾选 → 应用 → 报告；可一次撤销。
3. **UI 设计**：右侧 CTP（`msoCTPDockPositionRight`）；规则表格 + 命中列表（点击定位到文档）；规则冲突（前一条替换结果被后一条再次命中）显示 `fa-triangle-exclamation`。
4. **技术实现**：预览：对每条规则复制 `Range` → `Find.ClearFormatting()` → 设置 `Text/MatchCase/MatchWholeWord/MatchWildcards/Forward=true/Wrap=wdFindStop` → 循环 `Execute()` 收集 `(Start, End, 上下文)`；遍历 `Document.StoryRanges`（正文、脚注、页眉页脚、文本框）按用户选择；应用：按命中位置**从后往前**替换（避免偏移），使用 `range.Text = replacement` 或 `Find.Execute(Replace: wdReplaceOne)`；格式保留策略：默认继承首字符格式，可选“替换格式”。应用前重新校验每个命中范围的文本仍等于预览值（防止预览后用户编辑）。
5. **数据结构**：`ReplaceRule { id, find, replace, matchCase, wholeWord, wildcards, scope, enabled }`、`FindHit { ruleId, storyType, start, end, contextBefore, matchText, contextAfter, docVersionToken }`。
6. **模块接口**：`IDocumentEditService.PreviewFind(rules, scope)`、`ApplyReplacements(hits)`；规则集可存入 SQLite 供复用；D06 可调用（白名单）。
7. **边界场景**：修订模式（替换产生修订，提示用户）；受保护文档/只读（禁用应用）；域代码内文本；超大文档（命中 >10,000 条分页显示、可取消）。
8. **技术风险**：【官方】`Range.Find`；通配符语法为 Word 方言，需文档化。
9. **验收标准**：50 条规则、200 页文档：预览命中与 Word 原生查找计数一致；应用后无规则外变化（比较前后 `WordOpenXML` 差异仅限命中区）；一次 Ctrl+Z 完全撤销。

## C02 批量格式统一

**元信息**：P3 · 难度 中-高 · 风险【官方】 · 依赖 UndoRecord
1. **功能说明**：统一标题、正文、段落、字体、间距；优先修改样式而非逐段直接格式。
2. **用户流程**：选择格式方案（如“学位论文-某校”模板或自定义）→ 检查：列出不符合项（直接格式覆盖、未使用标题样式的伪标题）→ 预览 → 应用。
3. **UI 设计**：右侧 CTP；检查结果按类别分组，可逐条定位/忽略。
4. **技术实现**：样式层：`Document.Styles[wdStyleHeading1].Font/ParagraphFormat` 修改；直接格式检测：遍历 `Paragraphs`，比较段落 `Range.Font` 与其样式字体（返回 `wdUndefined` 表示混合）；清除直接格式：`Range.Font.Reset()` / `ParagraphFormat.Reset()`（用户确认后）；伪标题识别：加粗+字号大+独立短段落 → 只建议，不自动改。
5. **数据结构**：`FormatProfile { styles[{styleId, font, size, bold, spacingBefore, spacingAfter, lineSpacing, indent}], rules[] }`、`FormatFinding { paragraphIndex, range, issue, suggestion }`。
6. **模块接口**：`IFormatService.Check(profile)`、`Apply(findings)`。
7. **边界场景**：中英文字体分别设置（`Font.NameFarEast` 与 `Font.Name`）；表格内段落；目录/题注样式；模板链接样式（`AttachedTemplate` 更新风险）。
8. **技术风险**：【官方】API；“无关内容不变”需要差异校验工具（OOXML 比较）。
9. **验收标准**：样例论文应用后，所有标题样式符合方案；非目标段落 XML 无变化；可一次撤销。

## C03 文档合并与拆分

**元信息**：P3 · 难度 高 · 风险【官方】+【待实机】 · 依赖 A05、B02
1. **功能说明**：按顺序合并多个文档；或按标题级别/分节符拆分为多个文件。
2. **用户流程**：合并：选择文件（或工作区标签）→ 排序 → 选项（插入分节符/分页符、保留源格式/使用目标样式）→ 生成新文档（不修改源文件）；拆分：选择拆分依据 → 预览拆分点 → 选择输出目录与命名 → 生成。
3. **UI 设计**：向导对话框（WinForms）；结束后显示例外清单（页眉页脚差异、脚注编号重排、书签重名等）。
4. **技术实现**：合并：`Documents.Add()` → 逐文件 `range.InsertFile(FileName)`【TLB 已核实】并在文件间 `InsertBreak(wdSectionBreakNextPage)`；拆分：确定拆分点 Range → `range.FormattedText` 复制到新文档（`newDoc.Content.FormattedText = src.FormattedText`）或 `Range.ExportFragment/ImportFragment`【TLB 已核实】→ `SaveAs2`。复杂对象（页眉页脚、页码、样式冲突、修订、书签、脚注/尾注、交叉引用、目录域）逐项记录例外。
5. **数据结构**：`MergePlan { sources[], separator, styleMode, outputPath }`、`SplitPlan { splitBy, points[{start, title}], outputDir, namingPattern }`、`ExceptionList[]`。
6. **模块接口**：`IDocumentAssemblyService.Merge/Split`；批处理报告复用 C04 报告组件。
7. **边界场景**：源文档有密码/受限编辑；源文件已在 Word 中打开且未保存（提示使用磁盘版本或先保存）；不同纸张方向/页边距。
8. **技术风险**：页眉页脚与样式冲突行为复杂【待实机】，先以 3 类样本小范围验证。
9. **验收标准**：样本集合并后正文字数（`ComputeStatistics(wdStatisticWords)`）= 各源之和；拆分后各文件正文完整；例外清单覆盖实际差异。

## C04 批量导出 PDF

**元信息**：P3 · 难度 低-中 · 风险【官方】 · 依赖 B02、A13
1. **功能说明**：对当前标签或工作区文档批量导出 PDF，可设置命名规则与输出目录。
2. **用户流程**：选择文档 → 输出目录/命名（`{name}`、`{date}`）→ 冲突策略（跳过/覆盖/加序号）→ 开始 → 逐份报告。
3. **UI 设计**：对话框 + 进度条 + 结果表；导出完成可“打开输出目录”。
4. **技术实现**：已打开文档：`doc.ExportAsFixedFormat(OutputFileName, wdExportFormatPDF, OpenAfterExport:false, OptimizeFor: wdExportOptimizeForPrint, CreateBookmarks: wdExportCreateHeadingBookmarks, DocStructureTags:true)`【TLB 已核实】；未打开文档：`Documents.Open(path, ReadOnly:true, AddToRecentFiles:false, Visible:false)` → 导出 → `Close(wdDoNotSaveChanges)`（隐藏打开的文档不进入标签，WindowRecord 标记 `hiddenBy=CrabOfficeBatch`）。事前检查目标路径可写、文件名非法字符、长度。
5. **数据结构**：`ExportJob { items[{sourcePath|windowId, targetPath, status, error, durationMs}], options }`。
6. **模块接口**：`IExportService.ExportPdf(job)`；D06 可调用。
7. **边界场景**：文档含宏/链接更新提示（以 `DisplayAlerts = wdAlertsNone` 临时抑制，结束恢复，记录被抑制提示）；受保护视图文件（需要信任才能导出，跳过并说明）；输出目录为源目录同名 PDF 被占用。
8. **技术风险**：【官方】。
9. **验收标准**：成功/失败/跳过与报告完全一致；源文档未被修改（时间戳不变）。

## C05 章节与书签导航

**元信息**：P3 · 难度 中 · 风险【官方】 · 依赖 A13
1. **功能说明**：在侧边面板中按标题树、书签、表格、图片、批注快速跳转（增强 Word 导航窗格，增加书签/表格/批注分类与搜索）。
2. **用户流程**：打开导航面板 → 选择分类 → 点击条目 → 文档滚动到对应位置并选中。
3. **UI 设计**：右侧 CTP，Tab 页：标题 `fa-list-ol`、书签 `fa-bookmark`、表格 `fa-table`、批注 `fa-comment`；顶部搜索框。
4. **技术实现**：标题：遍历 `Paragraphs` 读取 `OutlineLevel`（或 `Document.GetCrossReferenceItems(wdRefTypeHeading)`）；书签 `Document.Bookmarks`（`ShowHidden` 控制隐藏书签）；表格 `Document.Tables`；批注 `Document.Comments`；定位 `range.Select()` + `ActiveWindow.ScrollIntoView(range)`；刷新：文档切换时 + 手动刷新 + 选区变化节流（大文档不做实时）。只读不改正文。
5. **数据结构**：`NavNode { kind, title, level, start, end, children[] }`。
6. **模块接口**：`INavigationService.BuildIndex(documentId)`；D02/D03 的引用位置跳转复用 `NavigateTo(location)`。
7. **边界场景**：非标准格式标题（无样式）→ 显示“仅识别使用标题样式或大纲级别的段落”；位置在内容变更后失效 → 用文本指纹二次查找。
8. **技术风险**：【官方】；大文档遍历性能【待实机】。
9. **验收标准**：样例文档所有标题/书签/表格/批注可定位；导航后 `Saved` 不变。

## C06 文本片段库

**元信息**：P3 · 难度 低-中 · 风险【官方】 · 依赖 存储层
1. **功能说明**：保存常用文本（纯文本或带格式）并插入光标位置。
2. **用户流程**：选中文本 → 右键/面板“保存为片段”→ 命名/分类；插入：面板双击或命令面板搜索 → 插入。
3. **UI 设计**：侧边 CTP，搜索 + 分类；每个片段标注“纯文本”`fa-font`/“富文本”`fa-paintbrush`。
4. **技术实现**：纯文本：SQLite 存 `text`，插入 `Selection.Range.Text = …` 后折叠到末尾；富文本：保存 `Range.WordOpenXML`（Flat OPC）→ 插入 `Selection.Range.InsertXML(xml)`【TLB 已核实】；备选方案为 Word 原生构建基块（`BuildingBlockEntries`，存于模板）。插入包裹 UndoRecord。
5. **数据结构**：`Snippet { id, name, category, kind: Plain|Rich, text?, flatOpcXml?, createdAtUtc, useCount }`。
6. **模块接口**：`ISnippetService.Save/Insert/Search`；C08 命令面板可插入片段。
7. **边界场景**：插入点在受保护区域/内容控件锁定；富文本带样式与目标冲突（可选“匹配目标格式”= 仅插入纯文本）。
8. **技术风险**：【官方】；InsertXML 对部分复杂对象的支持需验证。
9. **验收标准**：插入位置正确，单次撤销可回退；富文本格式保持。

## C07 文档比较

**元信息**：P3 · 难度 中 · 风险【官方】 · 依赖 A05、C05
1. **功能说明**：比较两份 Word 文档，展示新增、删除、修改（含格式）。
2. **用户流程**：选择原文档与修订文档（可从标签选择）→ 选项（比较格式/批注/页眉页脚、粒度：字符/单词）→ 生成比较结果文档（新标签）→ 侧边面板列出差异并可定位。
3. **UI 设计**：对话框 + 差异列表（`fa-plus` 新增/`fa-minus` 删除/`fa-pen` 修改）。
4. **技术实现**：`Application.CompareDocuments(OriginalDocument, RevisedDocument, Destination: wdCompareDestinationNew, Granularity: wdGranularityWordLevel, CompareFormatting, CompareHeaders, CompareComments, …)`【TLB 已核实】→ 遍历结果文档 `Revisions` 生成差异列表（`Revision.Type`、`Range`）。不以纯文本 diff 替代。
5. **数据结构**：`CompareResult { resultWindowId, diffs[{type, start, end, text, author}] }`。
6. **模块接口**：`ICompareService.Compare(originalRef, revisedRef, options)`。
7. **边界场景**：文档未保存（CompareDocuments 需要 Document 对象，可直接用已打开对象）；受保护文档；比较结果文档默认未保存。
8. **技术风险**：【官方】。
9. **验收标准**：人工构造差异样本，新增/删除/修改均能列出并定位。

## C08 快捷命令面板

**元信息**：P3 · 难度 中 · 风险 低 · 依赖 CommandDispatcher、A12
1. **功能说明**：快捷键唤出命令面板，搜索执行“保存全部、打开工作区、导出 PDF、切换到某文档”等。
2. **用户流程**：按快捷键（默认 Ctrl+Shift+P，待冲突评估）→ 浮动输入框 → 输入关键词（支持拼音首字母 P2）→ 回车执行；Esc 关闭。
3. **UI 设计**：Word 窗口中上方浮动窗体（owned by Word 窗口，`TopMost=false`），显示命令名、快捷键、不可用原因（置灰）。
4. **技术实现**：`CommandRegistry` 收集所有 `ICommand`（id、标题、关键词、`CanExecute`）；文档切换项由 TabRegistry 动态生成；执行统一走 `ICommandDispatcher.Execute`；无效输入不执行任何操作。
5. **数据结构**：`CommandDescriptor { commandId, title, keywords[], shortcut?, category }`、`PaletteState`。
6. **模块接口**：`ICommandRegistry.Search(query)`、`ICommandDispatcher`。
7. **边界场景**：模态对话框打开时不唤出；面板焦点丢失即关闭并把焦点还给 Word；IME 输入中文关键词。
8. **技术风险**：低；浮窗 DPI 与焦点【待实机】。
9. **验收标准**：所有注册命令可搜索执行；不可执行命令显示原因且执行无副作用；关闭后焦点回到正文。

---

# D 组：AI 智能与研究功能（后续模块）

> D 组通用架构（详见 03 §7）：VSTO 端 `IAiServiceGateway` ↔ 本地 `crab-ai` 服务（Python 3.11+、FastAPI、LangGraph，监听 `127.0.0.1` 随机端口 + 每次启动生成的 Bearer Token）↔ 多模型 Provider Adapter（本地模型/云端 API）。**AI 层不参与任何标签/窗口链路**；AI 服务不可用时 A/B/C 功能完全不受影响。AI 输出一律为结构化 JSON（Pydantic Schema 校验），插件端**只执行白名单命令**，所有修改需预览 + 用户确认 + 应用前原文版本校验。数据外发前显示目的地（本地/云端某 Provider）并需用户授权。
>
> 本机现状：Python 3.9.13；LangGraph 新版本对 Python 版本有最低要求【待查证具体版本】，建议研发时使用独立的 Python 3.11+ 虚拟环境或嵌入式发行版，不依赖系统 Python。

## D01 AI 润色与翻译

**元信息**：P4 · 难度 中 · 风险 模型质量/数据外发 · 依赖 AI 网关、UndoRecord
1. **功能说明**：选中中英文段落进行润色、改写、学术化、互译；预览差异后才写回。
2. **用户流程**：选中文本 → 右键菜单（Ribbon 自定义上下文菜单 `ContextMenuText`）或 AI 面板 → 选择操作与风格 → 生成 → 差异视图（删除红/新增绿）→ 接受（全部/逐句）/拒绝/重新生成。
3. **UI 设计**：右侧 AI 面板（未来 WebView2，首版 WinForms）；生成中可取消；显示“发送至：本地模型 / 云端 XX”标识 `fa-shield-halved`/`fa-cloud`。
4. **技术实现**：读取 `Selection.Range`：`Start/End/Text` + 段落边界 + `docVersionToken`（文本 SHA-256）；`SubmitTask({type:"polish", text, lang, style})` → LangGraph 图：`normalize → generate → self_check(不得新增事实/数字/引用) → schema_validate`；返回 `{segments:[{original, revised, reason}]}`；应用：重新取原 Range（`doc.Range(start,end)`），校验文本哈希一致后在 UndoRecord 中逐段替换（保留段落格式：仅替换段落内文本，不替换段落标记）；可选“以修订方式应用”（临时开启 `TrackRevisions`）。
5. **数据结构**：`AiTask { taskId, type, status, createdAtUtc, provider, input{textRef, hash}, output? }`、`RewriteSegment { original, revised, reason, confidence }`。
6. **模块接口**：`IAiServiceGateway.SubmitTask/GetTask/CancelTask`；`IDocumentEditService.ApplySegments(targetRange, segments, expectedHash)`。
7. **边界场景**：用户在生成期间修改原文 → 哈希不一致 → 拒绝应用并提示重新生成；选区跨表格/含域/脚注引用标记 → 仅处理纯文本部分并警告；超长选区分块。
8. **技术风险**：模型虚构；格式丢失（内联格式如上标/斜体在纯文本替换后丢失 → 首版提示用户，后续用 run 级映射）。
9. **验收标准**：未点击“接受”时文档 `Saved` 与内容不变；应用后可一次撤销；评测集（50 段）无新增事实/引用（人工抽检）。

## D02 全文/章节问答

**元信息**：P4 · 难度 高 · 风险 检索质量/数据范围 · 依赖 C05、AI 网关
1. **功能说明**：对当前文档（或工作区多文档）提问，回答附章节/段落引用，可点击跳转。
2. **用户流程**：AI 面板 → 选择范围（当前文档/选定章节/工作区）→ 提问 → 回答 + 引用列表 → 点击引用定位。
3. **UI 设计**：对话流；每条引用显示“第 3 章 · 段落 12”；无依据时显示“未在文档中找到依据”。
4. **技术实现**：插件端抽取：按段落导出 `{paraIndex, headingPath, text, start, end}`（UI 线程分批读取，避免卡顿）→ 发送给 AI 服务建立临时索引（内存/本地 SQLite + 向量库，如 sqlite-vec 或 FAISS【选型待定】）→ LangGraph：`retrieve(hybrid BM25+向量) → grade_relevance → answer_with_citations → verify_citations(引用文本必须是检索片段子串)`；无合格片段 → 返回 `insufficient_evidence`。
5. **数据结构**：`DocChunk { docRef, chunkId, headingPath, paraRange[start,end], text, hash }`、`Answer { text, citations[{chunkId, quote, location}], confidence, status: Answered|InsufficientEvidence }`。
6. **模块接口**：`IDocumentExtractService.ExportChunks(documentId)`、`INavigationService.NavigateTo(location)`。
7. **边界场景**：文档在索引后被修改 → 引用定位时用 quote 指纹重新查找，失败标注“位置可能已变化”；敏感文档 → 只允许本地模型。
8. **技术风险**：中文分词/检索质量；大文档抽取耗时【待实机】。
9. **验收标准**：问答评测集（≥30 题，含无答案题）：有答案题引用正确率 ≥ 90%，无答案题“不确定”识别率 ≥ 90%。

## D03 论文逻辑审查

**元信息**：P4 · 难度 高 · 风险 模型判断可靠性 · 依赖 D02、C05
1. **功能说明**：识别研究目标、方法、结果、结论之间的不一致，输出带证据和严重性的问题清单。
2. **用户流程**：AI 面板 → 论文审查 → 选择审查维度 → 运行（可后台，显示进度）→ Finding 列表 → 定位、忽略、添加批注（用户确认后用 `Comments.Add` 写入批注而非改正文）。
3. **UI 设计**：Finding 卡片：严重性（高/中/低 颜色 + `fa-circle-exclamation`）、位置、证据引用、建议。
4. **技术实现**：LangGraph 多节点：`extract_elements(按章节抽取目标/假设/方法/指标/结果/结论)` → `cross_check(目标↔结论、方法↔结果、数据一致性)` → `evidence_binding(每条 Finding 必须绑定≥1 原文引用)` → `dedupe_rank` → Schema 输出；支持断点续跑（LangGraph checkpointer，本地 SQLite）。
5. **数据结构**：`Finding { id, category, severity, description, evidence[{location, quote}], suggestion, status: Open|Ignored|Commented }`。
6. **模块接口**：`IAiServiceGateway`；`IDocumentEditService.AddComment(location, text)`。
7. **边界场景**：无标准章节结构的文档；超长论文（分章节并行，受模型配额限制）；用户取消。
8. **技术风险**：误报率；需要评测集与人工标注。
9. **验收标准**：每条 Finding 均含位置、证据、严重性、建议；证据引用 100% 可在原文找到；样本集人工评审精确率 ≥ 70%（阈值待负责人确认）。

## D04 本地文献与外部学术检索

**元信息**：P4 · 难度 高 · 风险 外部 API 可用性/版权 · 依赖 AI 服务
1. **功能说明**：导入本地 PDF/DOCX 建立文献库；检索开放学术元数据（OpenAlex、Crossref 等）；为写作提供可验证来源。
2. **用户流程**：文献库面板 → 导入文件/文件夹 → 解析与索引 → 搜索（本地 + 外部）→ 查看元数据/DOI/来源 URL → 插入引用（与 Word 引文或 Zotero 等集成为后续项）。
3. **UI 设计**：结果卡片标注“全文（本地）”/“仅元数据”/“开放获取链接”。
4. **技术实现**：Python 服务端：PDF 解析（如 PyMuPDF【许可证需评估 AGPL】或 pdfplumber）、DOCX（python-docx）；混合检索；外部：OpenAlex REST、Crossref REST（遵守礼貌访问规范、带联系邮箱、限速、缓存）；DOI 只来自 API 返回或文档原文，**禁止由模型生成**；元数据去重（DOI 优先，其次标题+年份模糊）。
5. **数据结构**：`LibraryItem { id, title, authors[], year, doi?, source, localPath?, hasFullText, url?, addedAtUtc }`。
6. **模块接口**：`IAiServiceGateway`（`library.*` 任务类型）；D02 可将文献库纳入检索范围。
7. **边界场景**：扫描版 PDF 无文本层（OCR 为后续项）；网络不可用（仅本地）；API 限流。
8. **技术风险**：第三方 API 条款与稳定性；PDF 解析库许可证。
9. **验收标准**：所有 DOI 可在 Crossref/OpenAlex 验证；缺全文条目明确标注；不编造任何字段。

## D05 中英文术语一致性

**元信息**：P4 · 难度 中-高 · 风险 低-中 · 依赖 C01、D02
1. **功能说明**：识别跨章节术语、缩写、译名不一致（如“大语言模型/大型语言模型/LLM 首次未定义”）。
2. **用户流程**：运行检查 → 术语簇列表（各变体出现次数与位置）→ 用户指定标准形式 → 生成替换建议（转交 C01 预览）→ 用户确认。
3. **UI 设计**：术语表 + 变体分布；“锁定术语”`fa-lock` 不参与建议。
4. **技术实现**：候选抽取（规则：括号缩写 `大语言模型（LLM）`、大写缩写、高频 n-gram）→ 用户词典（SQLite）→ 规则聚类（编辑距离、缩写展开）→ LLM 复核聚类（结构化输出）→ 结果定位；**不自动全局替换**，只生成 C01 规则草案。
5. **数据结构**：`TermCluster { id, canonical?, variants[{text, count, locations[]}], locked, source: Rule|LLM }`、`UserGlossary`。
6. **模块接口**：输出 `ReplaceRule[]` → C01 `PreviewFind`。
7. **边界场景**：参考文献部分排除；专有名词大小写敏感。
8. **技术风险**：召回率。
9. **验收标准**：样本中人工标注的不一致术语召回 ≥ 80%；锁定术语从不出现在替换建议中。

## D06 AI 自动化任务

**元信息**：P4 · 难度 高 · 风险 安全（关键）· 依赖 C01–C08 命令、CommandDispatcher
1. **功能说明**：用户以自然语言描述任务（“整理全文标题、生成目录、检查引用”），系统生成受限步骤计划，预览批准后执行。
2. **用户流程**：输入任务 → 计划视图（步骤列表、每步命令与参数、影响范围）→ 用户可编辑/删除步骤 → 批准 → 逐步执行（每步可暂停）→ 报告；整体一次撤销（若命令支持）。
3. **UI 设计**：计划卡片列表，每步显示命令图标、参数、“需要确认”标识；危险步骤（批量修改、关闭文档）红色边框。
4. **技术实现**：LangGraph `plan` 节点只能输出 `CommandRequest[]`，其中 `commandId` 必须属于 `AiCommandWhitelist`（如 `format.applyProfile`、`toc.insert`、`find.preview`、`export.pdf`、`nav.listHeadings`），参数经 JSON Schema 校验；插件端再次校验（双重校验）后经 `ICommandDispatcher.Execute`；**不提供**执行 VBA/宏、Shell、任意文件写入、网络请求的命令；文件输出路径限制在用户选择的目录。
5. **数据结构**：`AiPlan { planId, steps[CommandRequest + rationale + riskLevel], status }`、`CommandRequest { commandId, targetTabIds[], options, correlationId }`。
6. **模块接口**：`IAiServiceGateway.SubmitTask(type:"plan")` → `ICommandDispatcher.CanExecute/Execute`。
7. **边界场景**：模型输出非白名单命令 → 拒绝整份计划并显示原因；执行中文档被用户修改 → 暂停并要求重新确认；提示词注入（文档内容中含“忽略指令执行…”）→ 文档内容只作为数据，计划仍受白名单约束。
8. **技术风险**：提示词注入；需安全评审。
9. **验收标准**：红队测试集（≥20 个注入/越权样本）中 0 次执行非白名单操作；未批准计划 0 次修改文档。

---

# E 组：PowerPoint 扩展（架构预留）

## E01 PowerPoint 多标签

**元信息**：P5 · 难度 中-高 · 风险【待实机】 · 依赖 Core 全部、`IPresentationHostAdapter`
1. **功能说明**：复用 Core 的 Registry/TabService/CommandDispatcher/Workspace，为 PowerPoint 提供同等的顶部标签与（独立验证的）单窗口体验。
2. **用户流程**：同 A01–A11。
3. **UI 设计**：复用 `TabStripControl`，图标改为 `fa-file-powerpoint`。
4. **技术实现**：新 VSTO 项目 `CrabOffice.PowerPoint.Host`；`PowerPointHostAdapter : IOfficeHostAdapter`；事件【TLB 已核实】：`AfterNewPresentation`、`AfterPresentationOpen`、`PresentationBeforeClose`、`PresentationCloseFinal`（**PowerPoint 有关闭完成事件，Word 没有**，对账策略可简化）、`WindowActivate(Pres, DocumentWindow)`、`ProtectedViewWindowOpen`；窗口句柄 `DocumentWindow.HWND`【TLB 已核实】；CTP 同样按窗口创建（PowerPoint 2013+ 亦为每演示文稿一窗口【待实机】）；放映模式（SlideShowWindow）不显示标签。单窗口（WS-1）需单独 PoC，**不得假设 Word 结论可迁移**。
5. **数据结构**：复用 `DocumentRecord`（`hostKind=PowerPoint`）。
6. **模块接口**：`IOfficeHostAdapter` 公共抽象 + `IPresentationHostAdapter` 特有（放映、母版视图）。
7. **边界场景**：演示者视图双窗口；嵌入在其他应用的演示文稿；`Presentations.Open(WithWindow:=msoFalse)` 无窗口演示文稿（不显示标签）。
8. **技术风险**：【待实机】PowerPoint `DocumentWindow` 是否支持隐藏（PowerPoint OM 中 DocumentWindow **无** Visible 属性的可能性高【待查证】），WS-1 可能不可行，需要不同方案。
9. **验收标准**：打开/激活/关闭的演示文稿与标签一致（复用 A02/A03/A06 测试集）。

## E02 幻灯片批量操作

**元信息**：P5 · 难度 中 · 风险【官方】 · 依赖 E01
1. **功能说明**：批量复制、删除、移动幻灯片，统一字体/版式。
2. **用户流程**：选择幻灯片（缩略图多选或按范围/条件）→ 操作 → 预览影响清单 → 执行 → 报告。
3. **UI 设计**：右侧面板；预览显示受影响幻灯片编号缩略图。
4. **技术实现**：`Presentation.Slides.Range(int[])`、`SlideRange.Duplicate()`、`Slide.MoveTo(index)`、`Slide.Delete()`、`Slide.CustomLayout = …`；字体统一遍历 `Shapes` → `HasTextFrame` → `TextFrame2.TextRange.Font`；操作前记录 SlideID（`Slide.SlideID` 在演示文稿生命周期内稳定【官方】）而非索引。PowerPoint 无 Word 式 UndoRecord，撤销能力有限 → 执行前自动保存副本（用户可选）。
5. **数据结构**：`SlideBatchPlan { slideIds[], action, params, preview[] }`。
6. **模块接口**：`ISlideBatchService.Preview/Apply`。
7. **边界场景**：母版/版式占位符；组合形状；SmartArt；表格单元格文本。
8. **技术风险**：【官方】API；撤销能力弱。
9. **验收标准**：未选择幻灯片的 XML 无变化；报告与实际一致。

## E03 Word 转 PPT

**元信息**：P5 · 难度 中-高 · 风险 中 · 依赖 C05、E01、（可选）D 组
1. **功能说明**：从 Word 文档结构生成演示大纲，用户编辑确认后生成 PPT。
2. **用户流程**：在 Word 中选择“生成演示文稿”→ 大纲编辑器（标题→幻灯片、要点→项目符号；可选 AI 摘要）→ 选择模板 → 生成 → 在 PowerPoint 中打开继续编辑。
3. **UI 设计**：大纲树编辑器，每个节点显示来源位置（可回跳 Word）。
4. **技术实现**：Word 端 C05 抽取标题树与段落；（可选）D 组摘要要点（标注 AI 生成）；PowerPoint 端通过 COM 自动化（`new PowerPoint.Application()` 或 `Marshal.GetActiveObject`）→ `Presentations.Add()` → `Slides.AddSlide(index, CustomLayout)` → 填充占位符；每张幻灯片备注写入来源位置，保证可追踪。
5. **数据结构**：`OutlineNode { title, bullets[], sourceLocation, aiGenerated }`、`DeckPlan { template, nodes[] }`。
6. **模块接口**：`IDocumentExtractService`、`IDeckBuilder.Build(plan)`。
7. **边界场景**：PowerPoint 未安装/不同版本；模板缺少所需版式；超长段落自动拆分。
8. **技术风险**：跨应用 COM 自动化稳定性【待实机】。
9. **验收标准**：标题层级与 Word 一致；每张幻灯片可追溯到来源段落；AI 生成内容有明确标识。
