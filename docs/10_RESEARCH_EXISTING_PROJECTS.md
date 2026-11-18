# 10 · GitHub 成熟项目调研：Word 多标签插件（2026-11-18）

> 结论先行：发现两个高度相关的开源实现。**TabsForWord（MIT）与我们 PRD 的 A 组几乎完全重合且已解决我们正在踩的全部坑**，建议作为基础采用；officetab-rs 独立验证了我们的 S2 单窗口技术路线。原 CrabOffice 宿主层代码可作归档，Core 领域模型与 A04 成果继续沿用。

## 1. 检索方法

GitHub API + 原始 README 抓取，关键词：word add-in tabs / office multi tab vsto / word tabs plugin / wordtab；两仓库浅克隆到 Documents\CrabOffice-research 做了源码级分析。检索面覆盖后确认：**没有高星成熟项目**（该细分领域 GitHub 上 star 最高的是噪声），成熟产品在商业侧（Office Tab by ExtendOffice——TabsForWord ADR-001 提及其用 SetParent+钩子实现）。

## 2. TabsForWord（Onlgom/TabsForWord）— ★ 重点

- **协议**：MIT（可自由修改、商用、闭源衍生，需保留版权声明）
- **技术栈**：C# COM Add-in（手写 [ComImport] IDTExtensibility2 + ICustomTaskPaneConsumer），net48，**不依赖 VSTO SDK/Runtime**，dotnet SDK 即可构建，按用户安装无需管理员
- **规模**：31 个源文件，6685 行，v1.7.3，2026-07 至 2026-09 持续迭代
- **功能覆盖 vs 我们的 PRD**：
  - A01 标签栏 ✅（v1.6.0 起主模式为"窗口内宿主"，无 CTP 标题条）
  - A02 打开同步 ✅、A03 点击切换 ✅、A06 关闭保护 ✅（Reconcile 模式）
  - A07 未保存圆点 ✅、A08 拖拽排序 ✅、A09 固定+着色 ✅、A10 溢出/搜索 ✅
  - A11 右键菜单 ✅、A12 Ctrl+Tab ✅、A15 多显示器 DPI ✅（3 屏 100/125/150% 实测 13/13 PASS）
  - 受保护视图、Backstage、i18n（俄/英）均处理
- **关键技术（对 CrabOffice 的直接价值）**：
  1. **NATIVE_TAB_HOST（ADR-015）**：自建 WinForms WS_CHILD 子窗口 SetParent 进 Word 顶层窗口（OpusApp），锚定在 Ribbon 下方（定位 _WwG/_WwF/_WwB 类名，数十年稳定但未文档化）；WinEventHook（EVENT_OBJECT_LOCATIONCHANGE）即时重布局，**无子类化、无轮询**；连续 5 次失败自动回退 CTP。→ 这就是截图里橙色区域的正确打开方式（我们此前判断 CTP 无法到达该区域，它用子窗口宿主解决了）
  2. **ADR-004 幂等 Reconcile**：事件只触发共享重算（BeforeClose 后 300ms 延迟 tick + 1s 低频轮询）——与我们 Registry/Reconcile 设计互相印证
  3. **已知坑清单（KNOWN_ISSUES.md）**：CTP 最小高度 67px（标题条占 35px+）；后台窗口点击误邻（WM_MOUSEACTIVATE 需答"不激活"，v1.7.3 修复）；受保护视图编辑切换时 ActiveProtectedViewWindow 返回 null
  4. **Ctrl+Tab**：本地 WH_KEYBOARD 钩子（仅 Word 线程，非全局非低级）——绕过了 KeyBindings 局限，代价是表格内 Ctrl+Tab 制表符被屏蔽（已声明）
- **边界声明（ADR-001）**：**明确拒绝真单窗口**（合并窗口/SetParent Word 窗口），理由与我们的 S3 否决一致。**它没有单窗口模式**（源码无 SW_HIDE）

## 3. officetab-rs（washingtonwdc）— 单窗口路线的独立验证

- Apache 2.0，Rust COM Add-in，v0.1.0，"Vibe Coding"（同样 AI 辅助开发），葡语优先，声称支持 Word/Excel/PPT/Access/Project/Visio/Publisher
- **Single Window Mode = SW_HIDE 周转**（tab_manager.rs L823 注释：显示聚焦目标文档、SW_HIDE 隐藏前窗、单任务栏按钮）——与我们今日 PoC-SW-1 的实现**完全同构**，且我们本机 PoC 已见"启动后仅 1 个可见窗口"的探针证据
- 质量证据弱于 TabsForWord（0 star、无测试记录文档），但方向互相印证；含 window 子类化（subclass.rs），激进程度介于 S2 与 S3 之间

## 4. 对 CrabOffice 的策略建议（待项目负责人决策）

| 选项 | 内容 | 代价/风险 |
|---|---|---|
| **A. 采用 TabsForWord 为标签层基础（推荐）** | fork（保留 MIT 声明），验证其安装/运行，以其为 A 组底座；CrabOffice 的 S2 单窗口作为其原生宿主之上的扩展层（TabsForWord 明确不做，正是我们的差异化）；B/C/D 组按原 PRD 叠加 | 需审计 AI 生成代码质量；技术栈从 VSTO 转向 COM Add-in（偏离 D2 决策，需重新确认）；上游单维护者 |
| B. 自研继续 + 借鉴 | 保持现有 VSTO 宿主，逐步移植 native host / 已知坑修复 | 重复解决已踩坑（点击误邻、CTP 高度、PV 空指针…），周期长 |
| C. 采用 officetab-rs | 直接用其安装包 | 证据弱、成熟度低，不推荐为底座 |

**建议组合**：A 为主 + 保留我们已有的两项资产——
1. **A04 S2 单窗口**（已在本机初步验证，TabsForWord 无此能力，是 CrabOffice 的核心差异化卖点）
2. **Core 领域模型**（DocumentId/WindowId/TabId 三分离 + Reconcile 报告）比 TabsForWord 的模型更形式化，可在其上演进或用于评审对照

## 5. 风险提示

- TabsForWord 的 native host 依赖未文档化窗口类名（_WwG 等），Office 更新可能破坏——其设计含回退链（启发式→CTP），采纳后需纳入我们的回归矩阵
- 单窗口模式（S2）依赖非公开宿主行为，TabsForWord 作者的拒绝态度是保守选择而非技术证伪；我们已有初步 PoC 证据，继续按闸口推进但保持"失败即如实降级"纪律
- 两个项目均为 AI 辅助生成的单维护者代码，采纳前应跑一遍其手动/E2E 验证清单

## 6. 决策记录（2026-11-18）

- **TabsForWord 已装机实测通过**：MIT 协议，本机从源码构建（需把 csproj 的 GAC PIA 路径改为 VS 共享目录），HKCU 免管理员安装，原生窗口内标签栏、切换/固定/拖拽/着色/右键/Ctrl+Tab/搜索溢出全部正常。
- **S2 单窗口模式实验终止（用户决策）**：将 S2 隐藏周转集成到 TabsForWord 后，功能本身可用（探针证实单可见窗口、翻转正确），但 SW_HIDE/SW_SHOW 周期与 TabsForWord 的 Reserve（锚点下移）机制相互干扰，每次切换预留区累积漂移（用户截图证实空白持续增大）。已整体回滚到上游版本，放弃单窗口功能。
  - 技术备注：曾实现"隐藏时恢复锚点"的修复方向（NativeTabHost.UpdateLayout 加 IsWindowVisible 守卫），未及完整验证即按用户决策终止；若未来重启单窗口，应从该方向继续，并处理 S2 与第三方布局预留机制的耦合。
- **当前产品形态**：TabsForWord（上游原版）作为 Word 标签层；CrabOffice.WordHost 已停用（LoadBehavior=0）；CrabOffice 后续按 PRD 从 B 组（工作区）起步，复用/对照其 MIT 代码。
- 环境备忘：构建 TabsForWord 需 VS MSBuild（dotnet CLI 无法解析裸名 PIA）；git 在 D:\Program Files\Git\cmd；沙箱子进程需注入完整环境变量。
