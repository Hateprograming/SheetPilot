# 测试入口

第一阶段280例测试请看[使用说明](phase1/README.md)。下面只介绍较早的意图分类检查，不能运行280例表格处理测试。

# SheetPilot Intent Eval

`intent_cases.jsonl` 固定记录正常、歧义、越权、公式和复合任务表达。

在配置好 `.env` 后运行：

```powershell
$env:PYTHONPATH = "src"
python evals/run_intent_eval.py
```

退出码为 0 表示全部通过。每次修改模型、Prompt、Context 或 Skill 契约后都应执行；新增真实失败案例时，先将其匿名化，再加入 JSONL。

