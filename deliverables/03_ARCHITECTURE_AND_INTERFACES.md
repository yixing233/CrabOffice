# 03_ARCHITECTURE_AND_INTERFACES — 技术架构与接口规范（CrabOffice v1.0 设计稿）

## 1. 候选架构对比（至少两种）

### 架构 A：VSTO 单进程分层（推荐）
VSTO Add-in 进程内分四层：HostAdapter → Core → UI → Data。AI sidecar 独立进程。
- 优点：COM 对象全程 STA 主线程访问，无跨进程封送；CustomTaskPane 原生支持；调试直接。
- 缺点：插件崩溃影响 Word 进程（需全局异常兜底）；宿主生命周期绑定 Word。

### 架构 B：独立代理进程 + 轻量 COM 桥
独立 C# 代理进程通过 ROT/COM 或消息通道驱动 Word；Word 内仅留薄壳插件创建窗格。
- 优点：代理崩溃不拖垮 Word；可独立升级。
- 缺点：COM 跨进程封送慢且 STA/MTA 复杂；CustomTaskPane 仍需 Word 进程内组件；事件（DocumentOpen 等）跨进程转发易丢失；复杂度翻倍。
- 结论：**否决为主架构**，仅当 S2 路线要求崩溃隔离时局部复用其思想（如自愈守护进程）。

### 架构 C：Office.js —— 已在 02 号文档排除（不满足 A 组窗口能力），仅留作 AI 面板可选补充。

**选定：架构 A。**

## 2. 模块图

    +--------------------------------------------------------------+
    | Word 进程 (STA 主线程)                                        |
    |  WordHostAdapter (VSTO ThisAddIn)                             |
    |     | 订阅: DocumentOpen/NewDocument/BeforeClose/             |
    |     |       BeforeSave/WindowActivate/Deactivate/             |
    |     |       WindowSelectionChange                             |
    |     v                                                         |
    |  CrabOffice.Core                                              |
    |   EventNormalizer -> DocumentRegistry/WindowRegistry          |
    |   TabService -> WindowExperienceStrategy (S0/S1/S2 可插拔)    |
    |   CommandDispatcher -> WordOperationService                   |
    |   WorkspaceService / RecentClosedService / FavoritesService   |
    |     |                                                         |
    |     v  (领域事件, 主线程)                                     |
    |  CrabOffice.UI (WPF TabBarControl via ElementHost             |
    |    in CustomTaskPane, DockPositionTop, per Window)            |
    |     |                                                         |
    |     v                                                         |
    |  CrabOffice.Data (SQLite: workspace/tabs/favorites/           |
    |    recentClosed/snippets + settings.json 版本化)              |
    +--------------------------------------------------------------+
          | HTTP 127.0.0.1 (仅 D 组)                        |
          v                                                 |
    Python sidecar: FastAPI + LangGraph + ProviderAdapter   |

## 3. 领域模型（类型化）

    // 所有 Id 为会话内 GUID；跨会话持久化一律以 canonicalPath 为元数据键，非身份主键
    class DocumentRecord {
      Guid DocumentId;          // 会话生成，绑定 Document 弱引用
      string DisplayName;       // Document.Name
      string? CanonicalPath;    // 未保存为 null；保存后回填
      SavedState SavedState;    // Saved | Dirty | NeverSaved
      bool IsReadOnly;
      string? LastKnownFullName;
      WeakReference DocRef;     // 禁止持久化 COM 引用
    }
    class WindowRecord {
      Guid WindowId;
      Guid DocumentId;
      int? Hwnd;                // 仅窗口存活期有效；Hwnd 可得性【待查证】
      bool IsActive; bool IsVisible;
      WindowStateEx State;      // Normal/Min/Max/Hidden(S2)
    }
    class TabRecord {
      Guid TabId; Guid WindowId; int Order;
      bool IsPinned; string? GroupId; string DisplayLabel;
    }
    class Workspace {
      Guid WorkspaceId; string Name;
      List<WorkspaceItem> Items;  // canonicalPath + tabOrder + pinned
      Guid? ActiveTabId; DateTime UpdatedAt; int SchemaVersion;
    }
    class CommandRequest {
      Guid CommandId; string CommandName;
      Guid[] TargetTabIds; JsonElement Options; Guid CorrelationId;
    }

身份规则（A13）：DocumentId ≠ WindowId ≠ TabId；同一文档多窗口 = 多 WindowRecord 共享 DocumentId；HWND 仅作运行期窗口定位缓存，窗口销毁即失效；路径仅是元数据，不可作唯一身份（同名/另存/重开均会碰撞）。

## 4. 接口规范（签名 + 线程约束 + 失败语义）

所有接口：仅 STA 主线程调用（UI 与 COM 同线程）；异步操作一律返回后经主线程调度器回投；所有方法失败返回 Result<T>（错误码枚举），不抛裸异常穿越层边界。

    interface IWordHostAdapter {
      HostSnapshot EnumerateDocumentsAndWindows();   // 启动/对账快照
      Result ActivateWindow(Guid windowId);           // Window.Activate + 策略钩子
      Result OpenDocument(string path, bool readOnly);
      Result NewDocument();
      Result CloseTarget(Guid windowId, SavePrompt passthrough); // 走 Word 原生流程
      Result SaveTarget(Guid documentId);
      event HostEventHandler HostEvent;               // 归一化领域事件
    }
    // 领域事件枚举: DocOpened, DocClosedConfirmed, DocSaved, DocDirty,
    // WindowCreated, WindowDestroyed, WindowActivated, HostReady, HostShutdown

    interface IDocumentRegistry {
      void ApplyHostEvent(DomainEvent e);            // 唯一写入口
      ReconcileReport Reconcile(HostSnapshot snap);  // 对账修复（事件丢失兜底）
      RegistrySnapshot GetSnapshot();
    }

    interface ITabService {
      IReadOnlyList<TabRecord> GetTabs(Guid? scopeWindowId);
      Result ActivateTab(Guid tabId);        // -> strategy.SwitchTo
      Result MoveTab(Guid tabId, int newOrder);
      Result PinTab(Guid tabId, bool pinned);
      Result CloseTab(Guid tabId);           // 经 CommandDispatcher 走 Word 关闭
    }

    interface IWindowExperienceStrategy {     // A04 策略可插拔
      StrategyKind Kind;                      // S0_Multi | S1_Minimize | S2_HideTurnover
      Result Enter();
      Result SwitchTo(Guid targetWindowId);
      Result Exit();                          // 退出前恢复原生状态
      Result RestoreNativeState();            // 故障恢复逆操作（A16）
    }

    interface IWorkspaceService {
      Result SaveWorkspace(string name);
      Result<Workspace> LoadWorkspace(Guid id);
      Result<RestoreReport> RestoreWorkspace(Guid id); // 逐项打开+失败清单
      IReadOnlyList<RecentClosedItem> ListRecentClosed();
    }

    interface ICommandDispatcher {
      CanExecuteResult CanExecute(CommandRequest req); // 权限/只读/未保存/命名文档校验
      Result<CommandResult> Execute(CommandRequest req); // 逐项执行+逐项结果
    }

    interface IAiServiceGateway {              // 独立进程，不进标签链路
      Task<AiTaskHandle> SubmitTask(AiTaskSpec spec);   // 结构化 Schema 校验
      Task<AiTaskStatus> GetTask(Guid taskId);
      Result CancelTask(Guid taskId);
    }

### 线程与 COM 安全铁律
1. 禁止在后台线程持有/访问任何 Word COM 对象；后台仅允许纯数据（字符串/DTO）。
2. Registry 不保存 COM 引用，仅 WeakReference + 会话 GUID。
3. UI 变更经统一 EventQueue 主线程消费；事件处理器重入保护（防"激活事件→改 UI→再触发"环）。
4. 每次事件处理后做浅对账（计数+签名），每 30s 或关键节点做全量 Reconcile（A14）。

### UI-事件回流协议（防循环）
UI 动作 → TabService → HostAdapter → Word 触发事件 → Registry 更新 → UI 广播。
规则：UI 主动操作设置 suppress 标记（correlationId），对应回流事件被识别后清除；超时 2s 自动清除防死锁。

## 5. AI sidecar 契约（D 组）
- 端口 127.0.0.1 固定段 + 握手令牌（防任意本地进程调用）。
- 请求 Schema：{taskType, payload, contextScope, schemaVersion}；响应：结构化 JSON（如润色=原文哈希+候选段落+差异），插件端应用前做原文哈希 ratchet 校验。
- 白名单命令表（D06）：插件端仅执行枚举内命令，任何模型输出不得映射为任意代码/宏/文件操作。

## 6. PowerPoint 预留（E 组）
- IPresentationHostAdapter 同签名实现；Presentations 事件映射（PresentationOpen/BeforePresentationClose/WindowActivate…）。
- Core 的 Registry/TabService/CommandDispatcher 直接复用；WindowExperienceStrategy 需独立验证，不继承 Word 结论。
- 数据契约复用 DocumentRecord/TabRecord（DisplayName=演示文稿名）。

## 7. 配置与数据
- SQLite 表：workspaces, workspace_items, recent_closed, favorites, snippets, tab_order(会话)。
- settings.json（版本化 schemaVersion + 迁移器）：快捷键开关、单窗口策略选择、恢复会话开关、主题跟随。
- 所有持久化记录 SchemaVersion，向前迁移只增不破。
