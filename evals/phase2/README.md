# Phase 2 evaluation

This suite measures user-visible success, latency, interaction count, stability,
token use and cost without changing the Phase 1 benchmark.

```powershell
python -m evals.phase2 prepare
python -m evals.phase2 validate
python -m evals.phase2 run --suite locked --dry-run
python -m evals.phase2 run --suite locked
python -m evals.phase2 run --suite stability
python -m evals.phase2 suite-report
```

`prepare` creates the locked 2.1 suite under `.phase2/deepseek-flash-v2.1`. If its manifest
already exists, preparation only succeeds when every locked file still matches.
Official runs never overwrite completed instances. Use `--diagnostic` for a
rerun that must not affect the official score.

The suite contains a dated `pricing.json` snapshot. Reports always derive Token
counts from raw `calls.json` logs and calculate cost from that snapshot. Remove
the snapshot when a run should retain Token counts but report cost as `unpriced`.

