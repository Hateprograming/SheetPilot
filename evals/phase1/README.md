# 第一阶段测试：280例怎么运行

本工具有两种测试方式。**表格工具测试**直接给SheetPilot写好的操作规则，检查计划、处理文件和结果；它不检查云端AI是否能读懂用户的话。**完整流程测试**把案例中的自然语言要求交给当前配置的云端AI，再经过追问、计划检查、确认执行和结果检查。两种方式运行同一案例，案例总数仍算一例。

280例由原有40例和新增240例组成。新增案例包括160例正常处理、40例追问、20例拒绝、20例故障保护。新增案例中，24例使用SpreadsheetBench公开原文件，48例根据不同的调研问题另建可复现表格，168例按固定规则生成。公开原文件的要求经过改写，不算SpreadsheetBench原题成绩。故障保护只在表格工具测试中运行；其余256例可以进入完整流程测试。

## 先准备和检查

在 `SheetPilotApp` 目录的 PowerShell 中运行：

```powershell
.\.venv\Scripts\python.exe -m evals.phase1 prepare
.\.venv\Scripts\python.exe -m evals.phase1 validate
```

这会在 `.phase1/cases/` 中生成每例的输入文件、`case.json` 和独立答案。`.phase1/case-coverage.md` 是280例清单。再次准备会得到相同的文件内容；运行记录放在 `.phase1/runs/`，不会因准备案例而删除。原始调研文件只读取，不在原位置修改。

## 表格工具测试：一次跑全部280例

```powershell
.\.venv\Scripts\python.exe -m evals.phase1 run --mode offline
```

这里用预先写好的操作规则代替云端AI的回复。它能够发现文件处理、计划、版本和答案检查问题；无法证明云端AI理解正确。20例故障保护会故意制造超时、无效回复、步骤出错或计划版本失效。个别案例失败不会使其他案例没有结果。

## 完整流程测试：你按批次启动

确保项目 `.env` 已配置云端AI，然后在同一目录反复运行：

```powershell
.\.venv\Scripts\python.exe -m evals.phase1 run --mode live
```

每次最多运行20例、记录60次请求尝试、等待30分钟；每例最多6次请求。下一次会自动选择尚未运行的案例，不重复计入已完成的案例。内部重试也计数。全部256例有记录后，命令会提示没有剩余案例。测试工具不替你决定未知业务规则：计划不符合案例要求时不确认执行，遇到没有预设答案的追问会停止并记为需要人看。若确实要重跑已有记录，明确加 `--rerun`；原记录仍保留。

只跑指定案例、缩小单批数量或查看累计结果：

```powershell
.\.venv\Scripts\python.exe -m evals.phase1 run --mode live --ids L1 P-LOOKUP-01 --rerun
.\.venv\Scripts\python.exe -m evals.phase1 run --mode live --max-tasks 5 --max-calls 10
.\.venv\Scripts\python.exe -m evals.phase1 suite-report
```

案例筛选还有 `--kinds normal clarification refusal fault`，多个编号用空格分开。网页上的4条流程可用 `.\.venv\Scripts\python.exe -m evals.phase1 browser` 另行检查；网页测试使用写好的操作规则，不发云端AI请求，也不额外增加案例数。

## 结果在哪里

每批结果在 `.phase1/runs/日期-方式/`。`report.md` 用中文写逐例结果，`results.jsonl` 和 `summary.csv` 便于筛选，`run.json` 保存批次上限和版本，每例文件夹保存输入副本、对话、计划、调用记录、输出和具体差异。`.phase1/suite-report.md` 与 `.phase1/suite-summary.csv` 汇总所有批次。失败和需要人看的案例都保留；尚未启动的完整流程测试不会写成已通过。

`../Guidance/Phase1Evaluation.md` 是人工整理的阶段报告，不会随着运行命令自动改写。每次查看最新数量，以 `.phase1/suite-report.md` 或 `suite-report` 命令为准；阶段报告需要在复核测试结果后更新。

256例完整流程测试的逐项分析见[完整流程测试报告](../../../Guidance/Phase1FullFlowEvaluation.md)。

工具只在 `.phase1` 下写测试文件。它不会使用项目 `data` 或 `data_web` 中的私人工作簿。公开文件副本的来源和许可证写在案例里。公式案例目前只检查公式文字；没有独立计算结果时明确写“计算值未验证”。
