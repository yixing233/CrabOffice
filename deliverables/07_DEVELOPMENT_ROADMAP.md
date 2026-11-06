# 07_DEVELOPMENT_ROADMAP — 研发阶段划分与依赖（CrabOffice v1.0 设计稿，未批准编码）

> 优先级是排序而非永久排除：P0 未完成不阻塞 P2 的设计与接口冻结。

## Stage 0：环境调查与可行性（预计 1–1.5 周）
- 输出 environment-report.md（02 号文档第 1 节清单）。
- 批准并执行 PoC-TP-1（顶置窗格）、PoC-KB-1（快捷键）、PoC-ID-1（身份模型）。
- **提交 PoC-SW-1 实验申请**（目标/步骤/安全措施/判据），等待批准后执行。
- 退出标准：三份 PoC 报告 + environment-report + D2/D5 决策确认。

## Stage 1：基础标签 UI（UI-PASS）——P0
- 依赖：Stage 0。
- 内容：CrabOffice.Core 骨架（Registry/TabService/CommandDispatcher）、HostAdapter 事件订阅、CustomTaskPane 顶置 + WPF 标签栏（A01、A02、A13、A14 基础、A15）。
- 退出标准：T01、T02、T03、T12、T13、T22、T23 通过。

## Stage 2：标签交互完整（SWITCH-PASS）——P0
- 内容：A03 切换、A05 新建/打开、A06 关闭与保存保护、A07 未保存标记、A08 拖拽、A09 固定、A10 搜索溢出、A11 右键、A12 快捷键、A16 安全恢复（S0 策略）。
- 退出标准：T04–T15、T20、T21 全绿。**至此仅达成 SWITCH-PASS，明确不宣称单窗口。**

## Stage 3：单窗口闸口（SINGLE-WINDOW-PASS）——P0 关键闸口
- 依赖：批准的 PoC-SW-1 结果。
- 分支：
  - PASS → 实现 IWindowExperienceStrategy.S2（含 Enter/SwitchTo/Exit/RestoreNativeState）、自愈守护、T16–T19、T25 回归。
  - FAIL → 提交失败报告 → 用户决策 D1 → 若选 S1：实现 Minimize 策略 + UI 明示“多窗口最小化模式”，更新 PRD 与验收定义。
- 退出标准：SINGLE-WINDOW-PASS 检查表全绿，或 D1 决策落地且文档已如实更名。

## Stage 4：工作区与会话（B 组）——P1
- 内容：B01–B05，SQLite schema + 版本迁移。
- 依赖：Stage 2（可并行 Stage 3）。
- 退出标准：T26–T28。

## Stage 5：Word 效率增强（C 组）——P2
- C01/C02（UndoRecord 撤销保障）→ C04/C05/C06 → C03（合并拆分，最复杂，先小样验证）→ C07/C08。
- 依赖：Stage 2；命令面板 C08 依赖 CommandDispatcher 全量注册。
- 退出标准：T29、T30 及 C 组各自验收。

## Stage 6：AI 能力（D 组）——P2
- 依赖：D3 决策；sidecar 独立项目并行启动（可与 Stage 2 并行开发）。
- 顺序：D01（往返安全闭环）→ D02（RAG）→ D05/D03 → D04 → D06（白名单执行器最后做，依赖 CommandDispatcher 成熟）。
- 退出标准：T31、T32 + D 组验收。

## Stage 7：PowerPoint 适配（E 组）——P3
- 依赖：Stage 2 架构稳定；IPresentationHostAdapter 独立 PoC（单窗口可行性**重新验证**）。
- E01 → E02 → E03。
- 退出标准：PPT 版 T01–T07 等效集 + E 组验收。

## 依赖关系总图
    Stage0 ──> Stage1 ──> Stage2 ──> Stage3(闸口) ──> Stage4(可并行)
                        ├──> Stage5          └──> Stage7(需再PoC)
                        └──> Stage6(sidecar 可提前并行)

## 每阶段通用工程要求
- 全局异常处理 + 结构化日志（文件轮转）。
- 每阶段交付更新至 01–06 号文档对应章节；验收以 05 号矩阵为准。
- 任何降级/变更必须同步更新需求定义，禁止实现与文档不一致。
