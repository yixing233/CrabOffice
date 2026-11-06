# 02_TECH_FEASIBILITY — 技术可行性研究（CrabOffice v1.0 设计稿）

> 标注体系：【官方支持】微软官方文档明确提供的 API/行为｜【官方行为】官方确认的宿主行为（需实机复核细节）｜【理论可行】有公开机制支撑但未经实机验证｜【待实机验证】需在目标环境实测｜【待查证】API 存在性未确认，不得臆断｜【高风险/非公开】依赖未公开行为｜【不可行】官方路线缺失。

## 1. 环境调查清单（Stage 0 必做，未执行）
- Windows 10/11 版本与构建号；Windows 11 24H2 前后窗口管理差异。
- Word 版本：2016/2019/2021/2024/Microsoft 365；安装类型 MSI vs Click-to-Run；32/64 位（注册表 HKLM\SOFTWARE\Microsoft\Office\...\Outlook\Bitness 或 Process 判断）。
- Visual Studio 2019/2022（VSTO 工具负载）、.NET Framework 4.8 Developer Pack、VSTO Runtime。
- Python 3.10+（AI sidecar）、SQLite。
- 显卡/DPI：主副屏缩放组合（100/125/150/200%）。
- 输出：environment-report.md，所有后续结论标注适用版本。

## 2. 宿主技术选型对比

| 维度 | A: VSTO Add-in（推荐候选） | B: COM Add-in（IDTExtensibility2） | C: Office.js (Web Add-in) | 结论 |
|---|---|---|---|---|
| CustomTaskPane 顶置窗格 | 支持 | 需自封装 ICustomTaskPane | 不支持原生窗格 | A 胜 |
| Word COM 事件完整订阅 | 支持 | 支持 | 不支持窗口/文档级事件 | A 胜 |
| 窗口/句柄级操作前置能力 | 等同（均可调 Win32） | 等同 | 无 | A=B |
| UI 现代化（AI 面板） | WinForms/WPF/WebView2 | 同左 | 自身即 Web | 补充项 |
| 部署 | ClickOnce/MSI+VSTO Runtime | MSI+注册表 | AppSource/侧载 | B 部署更重 |
| 团队成本 | C#+VS 工具链最完整 | 手写互操作多 | 需前端栈+宿主受限 | A 胜 |
| 满足 A01–A16 | 可（除 A04 官方路线） | 可（同左） | **不可行**（无窗口句柄/无法宿主标签） | 排除 C 作为宿主 |

结论：主宿主选 A（VSTO）。Office.js 仅保留为未来 AI 面板的**可选补充**，不作为任何 A/B 组功能依赖。【官方支持】
另一候选 B'（独立代理进程 + 轻量 COM 桥）详见 03 号文档架构对比。

## 3. 各组关键 API 官方能力清单

### A 组（标签与窗口）
| API | 用途 | 状态 |
|---|---|---|
| Microsoft.Office.Tools.CustomTaskPanes.Add(control, title, window) | 为指定 Word.Window 创建专属窗格 | 【官方支持】 |
| CustomTaskPane.DockPosition = msoCTPDockPositionTop；.Height | 顶置与高度控制 | 【官方支持】，与 Ribbon 相对位置/最小高度【待实机验证】（PoC-TP-1） |
| Application.DocumentOpen / NewDocument / DocumentBeforeClose / DocumentBeforeSave / WindowActivate / WindowDeactivate / WindowSelectionChange | 文档与窗口事件源 | 【官方支持】 |
| Application.Documents.Add / Open(ReadOnly:=...) | 新建/打开 | 【官方支持】 |
| Window.Activate() | 切换窗口 | 【官方支持】 |
| Window.WindowState = wdWindowStateMinimize/Maximize | 窗口状态控制（S1 路线基础） | 【官方支持】 |
| Document.Saved | 未保存标记 | 【官方支持】 |
| Document.Close/Save/SaveAs2 | 关闭与保存流程 | 【官方支持】 |
| Application.KeyBindings | 快捷键绑定 | 【官方支持】；但只能绑定内置命令名或宏，VSTO 回调需宏 shim 或内置命令转发；Ctrl+Tab 是否可重绑【待实机验证】（PoC-KB-1） |
| Word.Window.Hwnd | 句柄获取 | 【待查证】：PowerPoint.Window 有 Hwnd 官方属性；Word.Window 属性表未见 Hwnd，备选 Caption+EnumWindows/GetForegroundWindow 匹配 |
| SW_HIDE/ShowWindow/SetForegroundWindow (user32) | S2 单窗口路线 | 【理论可行】+【高风险/非公开行为组合】，仅 PoC |
| SetParent | 窗口宿主化 | 默认否决（S3） |

关键官方行为：Word 2013 起改为 SDI——每个文档对应独立顶层框架窗口，"New Window" 只为同一文档创建额外视图窗口。对象模型不提供"一个主窗口承载多文档"的能力。【官方行为】→ **A04 无官方路线，见第 4 节。**

### B 组
- 文件枚举与路径：Document.FullName/Path/Name【官方支持】。
- 无官方"会话恢复"API，需自行持久化（SQLite/JSON）+ 逐项 Documents.Open【官方支持】。
- Word AutoRecover 与插件会话恢复是两套机制，不可替代【官方行为】。

### C 组
- Range.Find.Execute（Find.Text 上限约 255 字符，长文本需切片）【官方支持】。
- Application.UndoRecord（BeginCustomRecord/EndCustomRecord）合并撤销单元【官方支持，Word 2013+】。
- 样式操作：Document.Styles / Range.ParagraphFormat / Font【官方支持】。
- Document.SaveAs2 + ExportAsFixedFormat(PDF)【官方支持】。
- Document.GetCrossReferenceItems(wdRefTypeHeading)、Bookmarks、Tables、Comments【官方支持】。
- Document.Compare / Application.CompareDocuments【官方支持】。
- 禁止直接改 DOCX zip 内部 XML 字符串（破坏 content types/关系/校验和）——规约性结论。

### D 组
- 读取 Range.Text 传给本地服务【官方支持】；写回需保存 Range 起止与原文快照，应用前重校验版本（ratchet 校验）。
- Word 无官方 Python/LangGraph 接口；sidecar 走 HTTP(127.0.0.1) 或命名管道，插件端白名单命令执行——架构约定，非宿主能力。

### E 组
- PowerPoint.Window.Hwnd【官方支持（需确认目标版本）】；Presentations/Slides/Slide COM 模型【官方支持】。
- PPT 2013+ 同样 SDI；单窗口可行性与 Word 需**独立**验证，不得沿用结论。

## 4. 真正单窗口（A04）专项研究结论

### 4.1 三层验收闸口（不得混用）
1. UI-PASS：顶置标签栏可见可交互。
2. SWITCH-PASS：标签准确切换文档，关闭/保存与原生兼容。
3. SINGLE-WINDOW-PASS：仅一个主要可交互 Word 工作窗口；任务栏单条目、Alt+Tab 单条目、键盘焦点唯一、模态对话框归属正确、保存/打印/关闭行为正常。

### 4.2 策略矩阵
| 策略 | 描述 | 依据 | 可达层级 | 风险 |
|---|---|---|---|---|
| S0 | 多窗口 + 标签同步切换（保底） | 全官方 API | SWITCH-PASS | 低；但任务栏/Alt+Tab 多条目，不满足单窗口 |
| S1 | 唯一活动窗口，其余 WindowState=Minimize | 官方 API | 不达 SINGLE-WINDOW-PASS（最小化窗口仍占任务栏/Alt+Tab） | 低；可作为**如实命名的降级项**，需用户决策 D1 |
| S2 | SW_HIDE 隐藏非活动窗口 + 激活周转（先显示新→激活→再隐藏旧） | user32 公开 API + 非公开宿主反应 | **唯一可能达 SINGLE-WINDOW-PASS**（隐藏窗口自动移出任务栏/Alt+Tab） | 高：Word 可能自行重新显示；焦点序列；模态对话框孤儿；崩溃遗留隐藏窗口；Office 更新致失效。必须 PoC-SW-1 + 逆操作 + 自愈 |
| S3 | SetParent/宿主窗口/DLL 注入 | 非公开用法 | 理论上 | 极高（DPI 混合、COM 生命周期、更新破坏、卸载残留）→ 默认否决，仅隔离评估，不进 PoC 主线 |

### 4.3 判定
- **SINGLE-WINDOW-PASS 无官方实现路线。** S2 是唯一候选且依赖非公开宿主反应，成败只能由 PoC-SW-1 实测裁定。
- 无论结果如何，禁止把 S0/S1 重新定义为"真单窗口"（需求方第 9 节红线）。
- 若 PoC-SW-1 失败：提交失败报告（现象、Office 版本、失败检查项），由项目负责人决策 D1。

## 5. Word 版本适用性
- 最低目标：Word 2016（SDI 已稳定，UndoRecord 可用）。
- VSTO 仅支持已安装 PIA/VSTO Runtime 的桌面 MSI/C2R 版；不支持 Word Online/macOS。
- 32/64 位匹配：插件进程位随 Office 而非系统，Win32 调用不受影响（用户态同进程）。
- Microsoft 365 更新通道风险：每月通道可能改变窗口行为 → S2 类方案必须列入回归矩阵。

## 6. 可行性结论速览
- A01/A02/A03/A05–A16：官方 API 路线明确，个别细节待实机验证 → 可进入设计开发。
- A04：无官方路线；S2 PoC 闸口决定成败；存在失败可能，需预设降级决策。
- B/C 组：基本全官方 API，工程量问题而非可行性问题。
- D 组：宿主侧官方 API 足够；AI 侧为常规 Python 工程，无宿主可行性风险。
- E 组：官方 API 类似 Word，但必须独立验证，不继承 Word 结论。
