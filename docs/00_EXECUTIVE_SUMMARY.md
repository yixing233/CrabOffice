> 来源说明：本文件由 deliverables/00_SUMMARY.md 复制生成（2026-11-06），用于补齐 docs/ 系列引用的执行概要。内容与 deliverables/00_SUMMARY.md 保持一致，后续修订以 deliverables/ 版本为准。

# 00_SUMMARY — CrabOffice 交付概要（一页，供项目负责人审核）

生成日期：2026-11-05 ｜ 版本：v1.0（设计阶段，未批准编码）｜ 依据：CrabOffice_Agent_Prompt_v3.md

## 1. 推荐方案（候选，待批准）
- **宿主层**：C# VSTO Add-in（Word 2016+，.NET Framework 4.8），CustomTaskPane 顶置窗格承载 WPF 标签栏（ElementHost）。
- **核心层**：CrabOffice.Core 与 Office API 解耦——DocumentRegistry / WindowRegistry / TabService / CommandDispatcher / WorkspaceService。
- **单窗口策略分级**：S0 官方多窗口+标签切换（保底）→ S1 最小化非活动窗口（降级候选，需用户决策）→ S2 SW_HIDE 窗口隐藏周转（唯一可能达成 SINGLE-WINDOW-PASS 的路线，必须 PoC 闸口）→ S3 SetParent/宿主化（默认否决）。
- **AI 层**：独立 Python FastAPI + LangGraph sidecar（本机回环），与标签链路完全隔离。
- **数据层**：SQLite + 版本化 JSON 配置。
- **PowerPoint**：IPresentationHostAdapter 独立适配层，复用 Core 契约。

## 2. 核心不确定性（决定项目成败）
1. **真正单窗口（A04）无官方路线**：Word 2013+ 为 SDI（每文档一顶层窗口），对象模型不提供单窗口承载多文档的能力。S2/S3 均依赖非官方 Win32 行为，且随 Office 更新可能失效。**顶部标签栏成功 ≠ 单窗口成功。**
2. 顶置 CustomTaskPane 与 Ribbon/正文的实际相对位置、高度行为：官方 API 存在，但布局细节【待实机验证】。
3. Word.Window 是否暴露 Hwnd【待查证】；Ctrl+Tab 能否经 KeyBindings 重绑【待实机验证】。
4. S2 方案下 Word 自身是否会重新显示被隐藏窗口、模态对话框归属行为【待实机验证】。

## 3. 关键技术实验（最小 PoC，未执行）
| 编号 | 名称 | 验证问题 | 判据 |
|---|---|---|---|
| PoC-TP-1 | 顶置窗格 | Top Dock 位置/高度/多窗口行为 | UI-PASS 检查表 |
| PoC-SW-1 | 单窗口闸口 | SW_HIDE 周转能否过 SINGLE-WINDOW-PASS | 任务栏/Alt+Tab/焦点/对话框/保存/打印全绿，否则判失败 |
| PoC-KB-1 | 快捷键 | Ctrl+Tab 重绑与 IME 兼容 | A12 验收 |
| PoC-ID-1 | 身份模型 | 同名/多窗口/HWND 映射不串换 | A13 验收 |

PoC-SW-1 安全措施：仅用户态 ShowWindow，不注入、不 SetParent；VM 快照中执行；内置 RestoreNativeState 逆操作与启动自愈。

## 4. 待批准的决策
- D1：若 PoC-SW-1 失败，是否接受"S1 最小化降级"并如实更名（不得称真单窗口）。
- D2：确认 VSTO + .NET Framework 4.8 技术栈（Office.js 无法满足 A 组）。
- D3：AI sidecar 的部署形态（用户机独立安装包 vs 插件内嵌）。
- D4：Ctrl+Tab 若不可绑定的替代键位方案。

## 5. 交付物索引
01 功能 PRD ｜ 02 技术可行性 ｜ 03 架构与接口 ｜ 04 UI 规格 ｜ 05 验证矩阵 ｜ 06 风险与决策 ｜ 07 研发路线

## 6. 声明
本文档为设计产物。**任何功能在通过 05 号文档对应验收项之前均视为未实现**；文件存在不代表功能完成。
