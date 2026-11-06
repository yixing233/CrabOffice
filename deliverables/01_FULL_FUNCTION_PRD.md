# 01_FULL_FUNCTION_PRD — 逐项功能规格（CrabOffice v1.0 设计稿，未批准编码）

> 通用说明：每项含 ①功能说明 ②用户流程 ③UI 设计 ④技术实现 ⑤数据结构 ⑥模块接口 ⑦边界场景 ⑧技术风险与难度 ⑨验收标准。风险标注沿用 02 号文档体系。【官方支持】/【待实机验证】/【待查证】/【理论可行·高风险】。通用模块（EventNormalizer、CommandDispatcher、Registry、策略层）见 03 号文档，各条目引用不重复展开。通用边界（多窗口、异常退出、COM 主线程铁律）见 03/05 号文档。

---

# A 组：标签与窗口管理（P0）

## A01 顶部标签栏
① 在 Word 功能区下方、正文上方显示浏览器式标签条，显示文档名/图标/选中态/关闭入口/状态标识，不遮挡正文，主题自适应。
② 用户打开/切换文档→标签条实时反映；点击标签即切换（衔接 A03）。
③ 入口：随插件加载常驻。状态见 04 号文档状态矩阵。异常：窗格创建失败→toast 报错+日志，Word 正常可用。
④ 技术实现：ThisAddIn 启动时遍历 Application.Windows，对每个 Word.Window 经 CustomTaskPanes.Add(control, title, window) 创建专属顶置窗格（DockPosition=msoCTPDockPositionTop，Height≈40）；内容为 ElementHost 包裹的 WPF TabBarControl；WindowActivate 时确保目标窗口拥有窗格。所有窗格共享同一 TabService 快照，UI 仅做渲染。【官方支持】Top 与 Ribbon 实际相对位置/最小高度【待实机验证】(PoC-TP-1)。步骤：1 启动枚举→2 建 pane→3 绑定快照→4 事件回流渲染。
⑤ TabBarViewModel：tabs[](TabRecord)、activeTabId、theme、overflowItems。
⑥ 依赖 IDocumentRegistry.GetSnapshot()；消费领域事件 WindowActivated/DocOpened/DocClosedConfirmed/DocDirty。
⑦ DPI 变更 relayout；主题切换重读；窗格宿主窗口销毁时清理 pane 引用；Word 无文档时显示空态条。
⑧ 难度：中。风险：布局细节待实机。
⑨ T01+T23 通过=UI-PASS。

## A02 打开文档同步生成标签
① 任何路径打开/新建文档均自动增减标签，无重复无幽灵。
② 双击文件/Ctrl+O/最近列表/插件"+"→标签出现；关闭文档→标签消失。
③ 无独立 UI；异常：事件丢失由对账兜底。
④ DocumentOpen、NewDocument、WindowActivate 事件→EventNormalizer→Registry 唯一写入口；启动时 EnumerateDocumentsAndWindows 快照初始化；Reconcile 在激活/保存/定时(30s)触发。【官方支持】
⑤ DocumentRegistry 内部字典 DocumentId→DocumentRecord + 路径签名索引。
⑥ IWordHostAdapter.HostEvent→IDocumentRegistry.ApplyHostEvent→UI 广播。
⑦ 插件中途加载（快照）；同名文件（DocumentId 区分）；同文件重复打开（Word 原生聚焦已有窗口，快照确认）；受保护视图/云占位文件打开事件时序差异【待实机验证】。
⑧ 难度：中。风险：受保护视图等时序待实机。
⑨ T02、T03、T22 通过。

## A03 标签点击切换
① 点击标签准确切换到目标文档，保持选区/视图/滚动，不误操作其他文档。
② 点击→Word 激活对应窗口→选中态随 WindowActivate 回流确认更新。
③ Active 态仅在回流确认后设置（UI 不先斩后奏）。异常：窗口已销毁→toast"文档已关闭"并触发对账。
④ TabService.ActivateTab(tabId)→查 WindowRecord→IWindowExperienceStrategy.SwitchTo→Word.Window.Activate()【官方支持】；suppress 标记（correlationId）防"点击→事件→重复切换"循环，2s 超时自动清除。
⑤ 映射 tabId→windowId→(弱引用 Window)。
⑥ ITabService↔IWindowExperienceStrategy↔IWordHostAdapter；回流经 Registry→UI。
⑦ 目标窗口最小化时恢复 Normal 再激活；窗口句柄失效→对账重建；点击速率高→去抖合并。
⑧ 难度：低。风险：官方 API，低。
⑨ T04、T05 通过（20 标签循环切换一致、状态保持）。

## A04 真正单窗口多文档体验（关键技术闸口）
① 用户只见一个主要可交互 Word 工作窗口；标签在该窗口内切换文档；任务栏与 Alt+Tab 仅一条目；无多个独立 Word 主窗口。**明确不是侧边栏列表、不是多窗口切换器、不是顶部标签栏本身。**
② 用户点击标签→仅当前主窗口内容切换为目标文档，任务栏/Alt+Tab 不出现第二条目；打开新文档不再弹出新主窗口；关闭文档主窗口保留（直至最后一个）。
③ 单窗口模式为设置项（默认开，用户可关回 S0）；托盘/标签栏角标显示当前模式名，如实标注实现等级；模态对话框弹出时策略暂停周转。
④ 策略矩阵与判定见 02 号文档第 4 节：
   - S0 多窗口切换：全官方 API，为保底与开关回退。
   - S1 Minimize 策略：WindowState=Minimize【官方支持】，但不达 SINGLE-WINDOW-PASS，仅作 D1 决策后的**如实命名降级项**。
   - S2 HideTurnover 策略（唯一候选）：user32 ShowWindow(SW_HIDE/SW_SHOW)+SetForegroundWindow【理论可行·高风险】。SwitchTo 时序：显示目标→激活目标→稳定后隐藏旧窗口；Enter 时隐藏除首个外全部；Exit/RestoreNativeState 恢复全部可见。需先监测模态对话框（EnumWindows 检测 owned dialog）暂停周转。Word 是否自行重显、焦点序列等**必须 PoC-SW-1 实测**，未过闸口不得宣称实现。
   - S3 SetParent/注入：默认否决（02 号文档）。
   严禁把 A01 顶置窗格成功、A03 切换成功当作 A04 完成——三者验收层不同（UI-PASS/SWITCH-PASS/SINGLE-WINDOW-PASS）。
⑤ ExperienceModeSetting；策略内部 hiddenWindowIds 集合（崩溃自愈依据）。
⑥ IWindowExperienceStrategy（Enter/SwitchTo/Exit/RestoreNativeState）；崩溃后下次启动扫描可见性并恢复（A16 联动）。
⑦ 未保存文档：切换不触碰内容，保存流程交 Word 原生；多 Word 进程：策略仅管本进程，检测到第二进程→禁用单窗口模式并提示；异常退出：T25 自愈；Office 更新后：回归 T16–T19。
⑧ 难度：高。风险：S2 非公开行为组合，成败未知；失败则上报决策 D1，不得悄悄降级需求定义。
⑨ T16–T19+T21+T25 全绿=SINGLE-WINDOW-PASS；任一红灯即判失败并出报告。

## A05 新建和打开文件
① 标签栏"+"一键新建/打开。
② 左键"+"=新空白文档；右键/长按=菜单（新建/打开文件/打开工作区）。
③ 同 04 号文档。异常：打开取消（用户点 Esc）无副作用；只读文件标签带锁标。
④ CommandDispatcher→IWordHostAdapter.NewDocument()→Application.Documents.Add()；Open→Documents.Open(FullName, ReadOnly:=detected)【官方支持】；文件对话框由 Word 原生 FileDialog 承担，不自制，不绕过 Word 直接流写 docx。
⑤ 无新增结构（复用 Registry）。
⑥ ICommandDispatcher→IWordHostAdapter；结果回流生成 DocOpened。
⑦ 打开已打开文件→聚焦已有标签（去重）；打开失败（损坏/权限）→Word 原生错误框，插件无标签增量；拖拽文件到标签区列为后续项（本期不承诺）。
⑧ 难度：低。官方支持。
⑨ T02 已覆盖；补充用例：新建、打开、只读、重复打开行为各自明确。

## A06 标签关闭与保存保护
① 点 × 关闭文档；未保存走 Word 原生保存/不保存/取消流程；取消则标签保留。
② 点 ×→Word 保存框→用户选择→结果反映到标签。
③ ×仅悬停/激活态显示；固定标签无 ×。异常：保存失败→标签 Error 态+toast，不静默关闭。
④ CloseTab→CommandDispatcher（校验）→Word 原生关闭流程（Document.Close 或 Window.Close，prompt 由 Word 弹出）【官方支持】。**DocumentBeforeClose 只表示"准备关闭"，绝不能据此删标签**；仅以关闭后的 DocClosedConfirmed（对账确认）为准。区分同文档多窗口：关一个窗口≠关文档，以 WindowRecord 粒度处理。
⑤ RecentClosedService 登记（联动 B01）。
⑥ ITabService.CloseTab→ICommandDispatcher→IWordHostAdapter.CloseTarget。
⑦ 用户取消→无状态变化；SaveAs 对话框取消→文档保留标签保留；最后窗口关闭→Word 视宿主设置退出，插件按 HostShutdown 清理；崩溃在关闭流程中→对账修复。
⑧ 难度：中。风险：BeforeClose 语义与多窗口边界需实机确认细节。
⑨ T06、T07、T08 通过；取消场景文档与标签完整保留。

## A07 修改未保存标记
① 有未保存编辑时标签显示 ●；保存后消失；未命名文档状态明确。
② 编辑→●出现；Ctrl+S→消失；另存为→消失且标题更新。
③ ●主色圆点，NeverSaved 常显；只读文档不显示 ●。异常：状态失联→以 Document.Saved 实读为准。
④ 数据源 Document.Saved【官方支持】；触发时机：WindowSelectionChange（有界节流≤500ms）、DocSaved 事件、激活时、定时器兜底；**绝不写 Saved=true**；AutoSave/OneDrive 云状态与本地 Dirty 是不同语义，UI 文案区分（云文档显示云同步态而非 ●）。
⑤ DocumentRecord.SavedState 枚举（Saved/Dirty/NeverSaved）。
⑥ Registry.ApplyHostEvent(DocDirty/DocSaved)→UI。
⑦ 撤销回未编辑点：以 Document.Saved 实读为准（Word 会自行恢复 Saved）；批量替换等程序化修改同样触发刷新。
⑧ 难度：低。官方支持；节流策略需实机调参。
⑨ T09：编辑/保存/撤销/另存后状态与 Word 保存提示逻辑一致。

## A08 标签拖拽排序
① 拖动重排标签，不改 Word 真实文档顺序。
② 按住标签水平拖→插入指示线→释放落位。
③ DragOver 竖线指示；拖拽中禁用点击切换。异常：拖出窗格边界→取消归位。
④ 纯 UI 层维护 TabOrder（列表重排），**不依赖 Documents 集合顺序**；释放后写会话配置（SQLite tab_order 表）。
⑤ TabRecord.Order；会话持久化行（workspaceId/order 数组）。
⑥ ITabService.MoveTab→Data 层；UI 事件回流仅更新渲染。
⑦ 固定标签不可拖入常规区（分区墙）；重启后顺序恢复；会话内新开文档默认排尾。
⑧ 难度：低。官方支持（纯自管数据）。
⑨ T10：拖拽后激活与重启恢复顺序正确，文档内容零变化。

## A09 固定标签
① 重要文档固定于左侧分区，"关闭其他"不误关。
② 右键→固定/取消固定；图钉图标替代 ×。
③ 固定区与常规区有分隔线；固定标签不可拖出分区。异常：固定上限提示（建议 ≤8，可配）。
④ TabState.isPinned 持久化；批量命令（关闭其他/关闭右侧）执行前按 isPinned 过滤；CommandDispatcher.CanExecute 强制校验。
⑤ TabRecord.IsPinned。
⑥ ITabService.PinTab；CommandDispatcher 过滤逻辑共用。
⑦ 会话恢复保持固定态；被固定文档被外部删除→恢复时提示；关闭固定文档需显式点图钉取消后或菜单确认。
⑧ 难度：低。官方支持。
⑨ T11：固定、取消固定、关闭其他过滤、会话恢复一致。

## A10 标签过多、搜索和滚动
① 10–50 文档可快速定位：横向滚动/溢出菜单/模糊搜索；同名文件路径区分。
② 滚轮/拖滚条横移；点 🔍 输入即过滤（标题+完整路径模糊匹配）；同名项副行显示父目录；激活标签自动滚入可视区。
③ 溢出菜单(▼)列出全部；搜索浮层支持 Esc 关闭。异常：搜索无结果→空态文案。
④ WPF 虚拟化 ItemsControl（VirtualizingStackPanel）+ ScrollToHorizontalOffset；搜索为内存快照过滤（≤200 标签无需索引）。
⑤ 快照只读；无新增持久化。
⑥ ITabService.GetTabs→ViewModel；激活事件触发 BringIntoView。
⑦ 50 标签性能（帧率无卡顿）；超长路径截断中段；DPI 缩放下滚动条命中区。
⑧ 难度：中。官方支持（WPF 能力）。
⑨ T12、T13：50 标签可操作；同名文件搜索正确区分。

## A11 右键菜单
① 标签右键：保存/另存为/关闭/关闭其他/关闭右侧/复制路径/打开所在目录/固定/加入工作区。
② 右键弹出上下文菜单，仅作用于该标签对应文档。
③ 未命名文档"打开所在目录"禁用+tooltip"文档尚未保存到磁盘"；"复制路径"同样禁用；只读文档"保存"禁用。异常：批量关闭中途中断→已完成项保留，逐项反馈。
④ 全部菜单项经 CommandDispatcher.Execute（逐项 CanExecute：路径/权限/未保存校验）→Word 官方 API（Save/SaveAs2/Close/File 资源管理器 Process.Start("explorer.exe", /select))【官方支持】；批量关闭逐项走 Word 保存流程，不跳过任何提示。
⑤ CommandRequest{commandId, targetTabIds[], options, correlationId}。
⑥ ICommandDispatcher→IWordHostAdapter；"加入工作区"调 IWorkspaceService。
⑦ 未保存批量关闭逐项弹框，用户 Esc 中断→剩余项保持打开；固定项被"关闭其他"过滤；路径含特殊字符的 explorer 调用转义。
⑧ 难度：中。官方支持。
⑨ T14：12 项菜单逐一验证仅作用于目标标签；T11 固定过滤。

## A12 键盘快捷切换
① Ctrl+Tab / Ctrl+Shift+Tab 按标签顺序切换；设置可开关；不破坏 Word 已有快捷键。
② 按键循环切换；设置页有开关与备选键位。
③ 开关默认开；冲突时设置页提示。异常：绑定失败→toast+日志并回退。
④ 首选 Word 原生扩展：Application.KeyBindings.Add(WdKey, "宏/内置命令")【官方支持机制】——但 VSTO 回调需宏 shim 或绑定到内置命令再拦截【待实机验证 PoC-KB-1：Ctrl+Tab 能否重绑、IME 兼容】。**系统级全局键盘钩子不作为默认方案**，仅在证实不可行且用户明确同意后作为可选回退。
⑤ 设置项：shortcuts.enabled、shortcuts.binding。
⑥ 快捷键处理→ITabService.ActivateTab(下一个)。
⑦ 模态对话框打开时不响应；IME 组合窗口激活时按键直透【重点实测】；设置关闭后移除绑定（逆操作登记）。
⑧ 难度：中。风险：KeyBindings 对该组合键支持度待查证，钩子回退为高风险备选。
⑨ T15：正常编辑+中文 IME+对话框场景无冲突。

## A13 文档和窗口身份模型
① 同名文档、同文件多窗口、未保存文档、多 Word 进程均正确识别不串换。
② 透明能力，用户体验体现为一切定位准确。
③ 异常：身份冲突检测到→对账+日志。
④ DocumentId/WindowId/TabId 分离建模（03 号文档）；会话 GUID 绑定弱引用；Hwnd 仅窗口生命期内有效【Word.Window.Hwnd 待查证，备选 Caption+EnumWindows 匹配】；路径仅元数据非主键；三注册表 DocumentRegistry/WindowRegistry/TabRegistry 唯一写入口+Reconcile。
⑤ 见 03 号文档领域模型。
⑥ Registry 为所有功能的地基。
⑦ 另存为→路径回填、DocumentId 不变；"新建窗口"（同文档第二窗口）→新增 WindowRecord 共享 DocumentId，双标签或单标签多窗口标记【交互细节随 PoC-ID-1 定】；关闭重开→新 GUID+路径关联最近关闭；第二 Word 进程→策略层禁用单窗口并提示。
⑧ 难度：中。风险：Hwnd 获取待查证。
⑨ PoC-ID-1 + T04 + T13：同名/多窗口/关重开不串换。

## A14 事件同步与容错
① 高频事件下不重复注册、不死循环、不崩溃。
② 透明能力。
③ 异常：对账发现漂移→自动修复+日志计数。
④ EventNormalizer：事件→内部领域事件→Registry 唯一写入口→UI 广播；去重（事件签名）、重入保护（suppress+correlationId，2s 超时）、主线程调度（禁止后台线程触 COM）、浅对账每事件+全量 Reconcile 定时/关键节点。
⑤ ReconcileReport{added, removed, corrected}。
⑥ 贯穿所有模块的基础设施。
⑦ 连续批量开/关（T22）；Word 窗口交替激活；事件在模态期间堆积→模态结束后统一消化。
⑧ 难度：中。官方支持。
⑨ T22：风暴后 Registry 与 Word 实况零漂移，无崩溃。

## A15 窗口大小、多显示器与主题
① 各 DPI/双屏/最大化最小化/深浅主题显示正确。
② 透明能力。
③ 异常：黑屏/错位→自动 relayout+日志。
④ 顶置窗格自适应布局（WPF 布局天然 DPI 弹性）；PerMonitorV2 声明；Office 主题注册表读取【官方行为·键位待确认】+高对比度系统色；Win32 路线（若 S2 采用）专项 DPI 风险评估。
⑤ 主题缓存+变更监听。
⑥ UI 层与策略层联动。
⑦ 100→200% 实时切换；跨屏混合缩放；最大化/还原往返。
⑧ 难度：中。风险：S2 下跨 DPI 隐藏窗口恢复位置需实测。
⑨ T23、T24：指定矩阵无遮挡/焦点丢失/黑屏。

## A16 禁用、卸载与故障恢复
① 关闭/异常后 Word 完全恢复正常，无隐藏窗口、不阻碍编辑。
② 用户禁用插件→Word 行为与未装插件时一致。
③ 异常：toast+日志 id；连续 3 次启动失败→自动安全模式（仅 Registry 快照+对账，禁用策略与 UI 花活）。
④ 生命周期：OnShutdown 清 pane、解绑事件、恢复策略态；全局异常处理器（AppDomain+Dispatcher）→日志→隔离降级；**所有窗口控制操作必须登记逆操作**（RestoreNativeState：枚举本插件隐藏/最小化的窗口全部恢复可见）；启动自愈：扫描隐藏窗口集并恢复。
⑤ hiddenWindowIds 持久化（自愈依据）；日志文件轮转。
⑥ 策略层↔HostAdapter 生命周期钩子。
⑦ 任务管理器杀 Word（T25）；强杀插件进程（T20）；卸载后残留检查（EnumWindows 扫描）。
⑧ 难度：中。官方支持为主。
⑨ T20、T21、T25：全部原始窗口恢复可见、文档不丢失。

---

# B 组：会话、工作区与文档管理（P1）

## B01 最近关闭
① 列出最近关闭文档，点击重开。
② "+"菜单→最近关闭→列表→点击→Documents.Open。
③ 列表项含名称+路径+时间；失效项置灰提示"路径不存在"。未保存且无持久文件的项显示"（未保存，不可恢复）"，不得宣称可恢复。
④ 确认关闭（DocClosedConfirmed+有 CanonicalPath）后登记 {path, displayName, closedAt} 至 SQLite。
⑤ RecentClosedItem 表（上限 50，LRU）。
⑥ Registry→RecentClosedService→UI 菜单。
⑦ 文件被删/移动→打开报错并提示；无权限→权限错误文案；未命名未保存→不登记。
⑧ 难度：低。官方支持。
⑨ T27：已保存文件可恢复；丢失/无权限提示清楚。

## B02 工作区保存/恢复
① 按"论文/教学/项目"保存文档集合+标签顺序+固定态，一键恢复。
② "+"菜单→工作区→保存当前/打开/管理；恢复时逐项打开并显示进度与失败清单。
③ 恢复完成 toast"成功 N 项，失败 M 项"；失败清单含原因（缺失/权限/损坏）。
④ SQLite workspaces+workspace_items（versioned schema：路径、tabOrder、pinned、activeTabId、schemaVersion）；恢复=逐项 Documents.Open（绝不覆盖未保存内容——恢复前不关闭现有文档）。
⑤ Workspace 领域对象（03 号文档）。
⑥ IWorkspaceService↔ICommandDispatcher（打开走统一命令）。
⑦ 缺失文件→失败清单不中断；重复打开→去重聚焦；恢复期间用户取消→已完成项保留；未保存现有文档→不触碰。
⑧ 难度：中。官方支持。
⑨ T26：重启后可恢复；缺失项单独提示；不弹重复文档。

## B03 启动时恢复上次会话
① 可选"启动恢复上一工作区"，可关闭。
② Word 启动→插件加载→按开关恢复上次清单。
③ 开关在设置页；恢复进度 toast。异常：异常退出后仅恢复持久化存在的路径。
④ 正常关闭（HostShutdown）时保存会话清单；启动时 HostReady 后逐项打开；**不能替代 Word AutoRecover**（两者独立，文案说明）。
⑤ settings.json: session.restoreOnStart；会话清单表。
⑥ WorkspaceService 复用恢复逻辑。
⑦ 开关关→不恢复；上次异常退出→清单可能不完整，只恢复已持久化项；重复恢复→去重。
⑧ 难度：低。官方支持。
⑨ 开关有效、不弹重复文档（T26 变体）。

## B04 保存全部/关闭全部
① 一键处理工作区全部文档。
② 右键/菜单→保存全部/关闭全部→逐项处理→汇总报告。
③ 进行中逐项进度；取消/失败逐项反馈，无静默丢弃。
④ 逐文档 Save/Close（保留 Word 保存提示交互）【官方支持】；任一项用户取消→该项记"已取消"，继续或不继续按设置；权限不足/另存失败→逐项失败记录。
⑤ CommandResult{perItem results}。
⑥ CommandDispatcher 批量执行（复用 A11 管道）。
⑦ 只读文档→失败项；关闭全部含固定项→需确认框；中途 Word 弹模态→等待用户。
⑧ 难度：低-中。官方支持。
⑨ T28：不存在静默丢弃；取消后剩余文档状态正确反馈。

## B05 收藏和快速打开
① 收藏常用文件，从标签菜单快速打开。
② 右键→加入收藏；"+"菜单→收藏列表→点击打开。
③ 失效收藏置灰+tooltip"路径不存在"，不自动删除；不上传文件正文（纯本地）。
④ SQLite favorites{path, title, addedAt}；打开前 File.Exists+权限检测。
⑤ Favorites 表。
⑥ FavoritesService→CommandDispatcher.OpenDocument。
⑦ 路径失效→错误提示；文件移动→提示并可选更新路径；重名显示路径区分。
⑧ 难度：低。官方支持。
⑨ 失效收藏有清晰错误提示；正常收藏一键打开。

---

# C 组：Word 效率增强（P2）

## C01 多项查找替换
① 导入多条查找/替换规则，预览命中，选择性应用。
② 面板导入规则（CSV/手动）→预览命中列表（文档+位置上下文）→勾选→应用→报告。
③ 预览列表可勾选；应用中进度可取消。异常：正则非法→行内报错不执行。
④ Range.Find.Execute 定位【官方支持】；Find.Text 上限 255 字符→长规则切片处理【官方行为，需验证边界】；应用经 Application.UndoRecord.BeginCustomRecord 合并为单一撤销单元【官方支持】；**禁止直接改 DOCX zip 内部字符串**。
⑤ ReplacementRule{id, find, replace, matchCase, wildcards, scope, enabled}；HitRecord{rangeStart/end, context}。
⑥ 面板→CommandDispatcher（word.findReplace）。
⑦ 修订模式开启→提示且默认暂停应用；受保护文档→禁用；应用中用户编辑→ratchet 失败中止；取消→已应用项仍可整体撤销一次。
⑧ 难度：中。官方支持；长文本/通配符边界待验证。
⑨ T30：500 命中预览正确、应用正确、整体撤销一次回退、格式不变。

## C02 批量格式统一
① 统一标题/正文/段落/字体/间距。
② 选择方案（内置或自定义映射）→只检查→预览差异→应用。
③ 三步流程（检查/预览/应用）；差异列表可逐项展开。
④ 优先操作 Word Styles（修改样式定义）而非逐段刷格式【官方支持】；选中 Range 的 ParagraphFormat/Font 用于局部统一；应用前快照基线用于差异对比。
⑤ FormatScheme{id, styleMappings, directFormattingRules}。
⑥ 面板→CommandDispatcher（word.formatUnify）。
⑦ 文档含未知样式→映射到"正文"并列出；只读→禁用；应用后无无关内容变化（验收核心）。
⑧ 难度：中。官方支持。
⑨ 修改前后仅目标属性变化（对比快照），无关内容零变化。

## C03 文档合并与拆分
① 多文件按顺序合并；或按章节拆分为多文件。
② 合并：选文件+顺序+分节选项→预览清单→生成新文档；拆分：按 Heading1/节→预览切分点→生成多文件。
③ 进度与例外清单（图片/脚注/书签等处理情况）。异常：源文件打开中→提示先关闭或只读并入。
④ 合并：Selection.InsertFile / Range.FormattedText 拷贝【官方支持】，分节符 InsertBreak(wdSectionBreakNextPage)；拆分：按 Range 复制到新 Document 并 SaveAs2。样式冲突策略：目标样式优先/重命名映射，用户可选。修订先接受或拒绝（用户选）。页眉页脚/页码/书签/脚注逐节处理并在例外清单报告。
⑤ MergePlan{items[], sectionOptions}；SplitPlan{splitPoints[]}。
⑥ CommandDispatcher（word.merge/word.split）。
⑦ 复杂对象（跨文档链接图、OLE、目录域）可能失真→例外清单如实列出，不静默丢弃；目标盘空间检查。
⑧ 难度：高。官方支持但组合行为复杂，先小样验证。
⑨ 无关键正文丢失；复杂对象存在例外清单；抽验页码/脚注完整性。

## C04 批量导出 PDF
① 工作区文档批量导出 PDF，可选命名规则与目录。
② 面板选文档集+命名规则→冲突预检→导出→逐份报告。
③ 报告=成功/失败/跳过+原因；冲突项高亮。
④ Document.ExportAsFixedFormat(wdExportFormatPDF)【官方支持】；预检：目标存在→跳过或覆盖询问；只读源→正常（导出不改源）；路径非法→失败记录。
⑤ ExportJob{items[], namingRule, targetDir}。
⑥ CommandDispatcher（word.exportPdf）。
⑦ Word 占用文件被锁→失败记录；长任务可取消（已完成项保留）。
⑧ 难度：低。官方支持。
⑨ T29：10 文档含 1 只读 1 路径冲突→成功 8/失败 2 且有原因记录。

## C05 章节与书签导航
① 面板快速跳转标题/书签/表格/批注。
② 侧面板树形列出结构→点击→跳转定位。
③ 非标准格式（未用标题样式的伪标题）声明"有限支持"；空结构显示空态。
④ 读取 GetCrossReferenceItems(wdRefTypeHeading)、Bookmarks、Tables、Comments【官方支持】；跳转 Range.Select()+ScrollIntoView；结构缓存在 DocDirty 时重建（节流）。
⑤ NavNode{type, label, rangePos}。
⑥ 面板→CommandDispatcher（word.navigate）。
⑦ 文档变更后树刷新；超长文档（>500 节）虚拟化；跳转不改变正文。
⑧ 难度：低-中。官方支持。
⑨ 定位准确（抽验 20 项）、正文零改动。

## C06 文本片段库
① 保存常用文本，插入当前光标位置。
② 面板管理片段（增删改分类）→点击插入到光标。
③ 插入前标注对格式的可能影响（纯文本/带格式）；支持搜索。
④ SQLite 片段表{content, isRichText, category}；插入 Selection.TypeText 或 Range.FormattedText【官方支持】；保留 Word 原生撤销（不包裹 UndoRecord 也可撤销，如包裹则单片段一单元）。
⑤ Snippet 表。
⑥ 面板→CommandDispatcher（word.insertSnippet）。
⑦ 光标在只读区/表格内→插入结果如实；富文本片段源自其他文档样式→映射说明；空选区与选区替换两种语义（有选区=替换，提示）。
⑧ 难度：低。官方支持。
⑨ 插入位置正确（光标处/选区替换），可撤销。

## C07 文档比较
① 比较两份 Word 文档并展示差异。
② 选原文/修订版→生成比较结果文档→在新窗口/标签展示。
③ 结果文档以标签形式呈现（复用标签体系）；大文档提示耗时。
④ 优先 Application.CompareDocuments 或 Document.Compare【官方支持】，结果为 Word 修订标记的新文档；**不用纯文本 diff 替代**（丢失格式/修订差异）。
⑤ CompareJob{sourcePath, targetPath, resultDocId}。
⑥ CommandDispatcher（word.compare）；结果文档注册进标签系统。
⑦ 两文件其一打开中→提示先关闭或使用副本比较；比较后关闭结果文档按 A06 流程（未保存提示）。
⑧ 难度：低。官方支持。
⑨ 新增/删除/修改均可定位（构造已知差异样例验证）。

## C08 快捷命令面板
① 快捷键唤起搜索框，搜索并执行"保存全部、打开工作区、导出 PDF"等命令。
② 快捷键→浮层搜索框→输入过滤→Enter 执行→浮层关闭。
③ 无匹配命令显示空态；不可执行命令置灰+原因 tooltip。
④ 命令注册中心：所有 CommandDispatcher 注册项{id, name, keywords, canExecute}；执行统一走 ICommandDispatcher.Execute；权限/条件校验集中。
⑤ CommandDefinition 注册表（内存）。
⑥ 面板→CommandDispatcher（唯一执行通道）。
⑦ 输入无效命令→无副作用；命令执行中再次唤起→忽略；Esc 关闭无副作用。
⑧ 难度：中。官方支持（WPF 浮层+既有管道）。
⑨ 无效输入零副作用；全部注册命令可搜索可执行。

---

# D 组：AI 智能与研究功能（P2）

> 通用安全架构：插件宿主 ↔ 本地 sidecar（FastAPI+LangGraph）HTTP 127.0.0.1+握手令牌；结构化 Schema 请求响应；一切写回 Word 前经 ratchet 校验（原文哈希+Range 位置复核）+用户确认；sidecar 与标签链路完全隔离（AI 故障不阻塞标签功能）。

## D01 AI 润色与翻译
① 选中中英文段落润色/改写/互译，预览差异后应用。
② 选中文字→AI 菜单/面板选操作→进度→差异预览（对照视图）→应用/放弃。
③ 未经确认绝不修改 Word；预览含原文/译文对照；服务不可达→面板内联错误+重试。
④ 宿主读取 Selection.Range.Text+起止位置快照→提交 sidecar→LangGraph 流程（模型 Provider Adapter，结构化 Schema 返回候选段落）→插件端 ratchet 校验（当前 Range.Text 哈希==提交时哈希，起止未漂移）→用户确认→替换 Range.Text（UndoRecord 包裹）。
⑤ AiTaskSpec{taskType, payload, contextScope}；EditCandidate{originalHash, newText, diff}。
⑥ IAiServiceGateway↔CommandDispatcher（word.applyAiEdit）。
⑦ 响应超时/超长→中止提示；原文在等待期间被用户修改→ratchet 拒绝应用并提示重新选取；空选区→提示先选中文本；超长选区分段处理。
⑧ 难度：宿主侧中，AI 侧常规。官方支持（Range 读写）。
⑨ T31：篡改/超时/原文已变三种情况均不修改 Word；正常流应用后可撤销、无虚构引用数据（人工抽验）。

## D02 全文/章节问答
① 对论文内容提问，返回带位置依据的回答。
② 面板选范围（全文/章节）→提问→回答+引用位置列表→点击引用跳转（复用 C05 导航）。
③ 未检索到依据时明确回答"不确定"，禁止无依据作答；引用可跳转。
④ 本地解析（Range 遍历提取文本+结构）→章节切分→索引（关键词+向量混合）→LangGraph RAG 链→回答附 {章节, 段落, 原文片段}；模型仅接收授权上下文。
⑤ Chunk{id, docId, sectionPath, text, embeddingRef}。
⑥ IAiServiceGateway；跳转复用 CommandDispatcher(word.navigate)。
⑦ 未保存文档→提示先保存或标注"基于当前内存内容"；文档变更→索引失效重建；超长文档分批索引。
⑧ 难度：AI 侧常规，宿主侧低。
⑨ 构造已知答案文档抽验：回答正确且引用可跳转；无依据问题得到"不确定"。

## D03 论文逻辑审查
① 识别研究目标/方法/结果/结论间不一致。
② 面板发起审查→进度→Findings 列表（位置+依据+严重性+建议）→逐条查看→用户确认后生成修改（非自动改）。
③ 每条 Finding 必含四要素；用户不确认就不改。
④ LangGraph 分层提取（各章要素抽取）→跨章节一致性审查→结构化 Finding Schema→宿主定位引用。
⑤ Finding{id, location, evidence, severity, suggestion, status}。
⑥ IAiServiceGateway→用户确认→可选 D01 通道应用。
⑦ 模型输出缺位置/依据→该条丢弃并计数（防幻觉）；超长论文分章流水线。
⑧ 难度：AI 侧中。
⑨ 每条问题有位置、依据、严重性、建议；抽验定位准确。

## D04 本地文献与外部学术检索
① 导入本地 PDF/DOCX 建立文献库；检索开放学术文献。
② 面板文献库→导入/搜索→结果列表（本地命中/外部元数据）→查看/引用。
③ 缺全文时明确标注"仅元数据"；失效来源提示；不上传用户文档正文。
④ 本地：解析入库+关键词/向量混合检索；外部：OpenAlex/Crossref 公开 API【官方公开接口】；元数据去重（DOI/标题指纹）；来源 URL 与访问权限校验。
⑤ Reference{id, type, meta, sourceUrl, hasFullText, localPath?}。
⑥ IAiServiceGateway；本地解析可复用 officecli 类工具链。
⑦ 网络不可用→仅本地结果+提示；DOI 冲突→合并；**绝不编造 DOI**（仅展示 API 返回字段）。
⑧ 难度：中。外部 API 为官方公开。
⑨ 不编造 DOI；缺全文明确标注；去重正确。

## D05 中英文术语一致性
① 识别跨章节术语/缩写/译名不一致。
② 面板发起→不一致清单（位置+变体+建议统一形式）→用户确认逐项处理。
③ **保留用户指定术语表优先级；绝不自动全局替换**——仅输出定位清单，用户逐项决定。
④ 术语候选提取（规则+频率）→词典表比对→规则检测+LLM 复核（sidecar）→宿主定位输出。
⑤ TermTable{canonical, variants[], userLocked}；Finding 复用 D03 结构。
⑥ IAiServiceGateway→清单→用户逐项走 C01 单规则替换或手动改。
⑦ 用户术语表冲突→以用户为准并提示；缩写首次出现规范检查可开关。
⑧ 难度：中。
⑨ 用户指定术语被保留；输出清单定位准确；无自动全局替换行为。

## D06 AI 自动化任务
① 用户输入"整理全文标题、生成目录、检查引用"类指令，系统生成步骤并执行。
② 输入指令→展示生成的步骤计划（受限操作序列）→用户逐项/整体批准→按白名单执行→结果报告。
③ **预览与用户批准必不可少**；模型无法执行任意本地代码、宏或未经授权的文件操作（架构红线）。
④ LangGraph 规划器输出受限操作序列（仅引用插件白名单命令枚举，如 word.formatUnify/word.insertToc/word.checkRefs）→插件端逐一 CanExecute→批准→Execute→报告。任何模型输出若映射不到白名单枚举→拒绝并记录。
⑤ PlanStep{commandName(白名单枚举), params, status}。
⑥ IAiServiceGateway↔CommandDispatcher（复用 C08 命令注册中心）。
⑦ 规划产物越权→拒绝+日志；执行中失败→剩余步骤暂停待用户决定；T32 注入测试常驻回归。
⑧ 难度：高（安全设计为主）。官方支持（命令管道已有）。
⑨ T32：越权指令（删文件/宏/任意代码）全部被拒并记录；正常计划执行结果与预览一致。

---

# E 组：PowerPoint 扩展（P3，架构预留）

## E01 PowerPoint 多标签
① 将标签管理复用到 PowerPoint；经 PowerPoint 宿主适配层实现，不套用 Word API。
② PPT 中打开/切换/关闭演示文稿，标签同步。
③ 与 Word 版一致（04 号文档）；PPT 特有：演示模式时标签栏不可见/不干扰。
④ IPresentationHostAdapter 独立实现：Presentations 事件（PresentationOpen/BeforePresentationClose/WindowActivate 等映射）【官方支持】；Core 的 Registry/TabService/CommandDispatcher 原样复用；PowerPoint.Window.Hwnd【官方支持·需确认目标版本】；**单窗口可行性必须独立验证，不继承 Word 结论**。
⑤ 复用 DocumentRecord/TabRecord（DisplayName=演示文稿名）。
⑥ HostAdapter 实现同一接口族；策略层独立实例。
⑦ 演示模式（SlideShow）期间暂停策略；嵌入对象编辑窗口；PPT 的 SDI 行为与 Word 差异实测。
⑧ 难度：中-高。风险：单窗口策略独立验证，成败未知。
⑨ PPT 版 T01–T07 等效测试集通过；打开/激活/关闭与标签一致。

## E02 幻灯片批量操作
① 批量复制/删除/移动页面，统一字体/布局。
② 面板多选幻灯片→操作→预览范围→应用。
③ 应用前预览明确列出受影响页面；未选择页面不受影响（验收核心）。
④ PowerPoint COM：Slides.Range/Duplicate/Delete/MoveTo、Slide.Fonts/Master【官方支持】；批处理前快照受影响 Slide Id 集合。
⑤ SlideSelection{slideIds[]}。
⑥ CommandDispatcher（ppt.slideBatch）。
⑦ 组合/占位符/备注页处理；撤销经 PPT 原生 Undo。
⑧ 难度：中。官方支持。
⑨ 操作后未选中页面零变化（对比快照抽验）。

## E03 Word 转 PPT
① 选中 Word 文档构建演示大纲，确认后生成 PPT。
② 选文档→提取大纲→预览编辑（页/要点增删）→生成 PPT→后续人工编辑不受限。
③ 预览可编辑是大纲而非成品；生成后提示"可继续人工编辑"。
④ Word 结构提取（复用 C05：Heading/要点）+可选 AI 摘要（D 组）→PowerPoint 对象模型创建 Slides/TextFrame【官方支持】；模板可选。
⑤ OutlineNode{level, text, slideMapping}。
⑥ CommandDispatcher（ppt.fromWord）联动 D/C 模块。
⑦ 无标题结构的文档→声明有限支持或引导 AI 提炼（用户确认）；生成的 PPT 不自动保存，走正常保存流程。
⑧ 难度：中。官方支持。
⑨ 标题层级与主要观点可追踪（对照源文档抽验）；人工可继续编辑。
