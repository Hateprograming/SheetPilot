from __future__ import annotations

import copy
import csv
import hashlib
import json
import random
import shutil
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from evals.phase1.cases import stable_xlsx
from evals.phase1.oracle import reference

from .common import APP, ROOT, Case, DEFAULT_OUT, SEED, dump, dump_gzip, load_cases, protected_output, resolve_oracle_defaults, sha

MATRIX = {
    "lookup": (10, 6, 4, 2), "normalize": (8, 5, 4, 1),
    "aggregate": (10, 6, 4, 2), "sort_rank": (9, 4, 4, 1),
    "clean": (13, 3, 4, 2), "transform": (12, 6, 6, 2),
    "formula": (8, 2, 4, 2), "conversion": (10, 0, 2, 4),
}
JOURNEYS = ("direct", "multi_step", "clarification", "boundary")
CAP_ZH = {"lookup":"查找匹配", "normalize":"格式标准化", "aggregate":"聚合统计", "sort_rank":"排序排名",
          "clean":"数据清洗", "transform":"表格变形", "formula":"公式生成", "conversion":"文件转换"}
ROW_SIZES = {"small": (37, 73, 121, 181), "medium": (750, 1500, 3000, 5000), "large": (10000, 20000, 30000, 40000)}
PRICING_SNAPSHOT = {
    "provider": "deepseek", "model": "deepseek-flash", "currency": "CNY",
    "source": "https://api-docs.deepseek.com/zh-cn/quick_start/pricing/",
    "snapshot_date": "2026-10-22", "timezone": "Asia/Shanghai",
    "peak": {"input_per_million": 2.0, "cached_input_per_million": .04, "output_per_million": 8.0},
    "off_peak": {"input_per_million": 1.0, "cached_input_per_million": .02, "output_per_million": 4.0},
    "peak_windows": [
        {"start": "09:00", "end": "12:00", "weekdays": [0, 1, 2, 3, 4]},
        {"start": "14:00", "end": "18:00", "weekdays": [0, 1, 2, 3, 4]},
    ],
}


def _write_pricing_snapshot(out: Path) -> None:
    path = out / "pricing.json"
    if not path.exists():
        dump(path, PRICING_SNAPSHOT)


def op(kind: str, composite: bool = False) -> dict:
    operations = {
        "lookup": {"task_type":"lookup", "main_sheet":"交易", "source_sheet":"客户", "main_key_column":"客户编号",
                   "source_key_column":"客户编号", "return_columns":["客户等级"], "output_column_names":["客户等级"],
                   "duplicate_policy":"block", "confidence":1},
        "normalize": {"task_type":"normalize", "sheet":"交易", "column":"编号文本", "operations":["trim","nfkc"],
                      "output_column":"标准编号", "confidence":1},
        "aggregate": {"task_type":"aggregate", "sheet":"交易", "group_by":["地区"],
                      "metrics":[{"column":"金额", "function":"sum", "output_column":"金额合计"}], "filters":[],
                      "output_sheet":"地区汇总", "confidence":1},
        "sort_rank": {"task_type":"sort_rank", "sheet":"交易", "sort_keys":[{"column":"金额","direction":"desc"}],
                      "rank_column":"金额", "rank_output_column":"金额排名", "rank_method":"dense", "rank_direction":"desc",
                      "output_sheet":"金额排名表", "confidence":1},
        "clean": {"task_type":"clean", "sheet":"交易", "operation":"remove_duplicates", "key_columns":["订单号"],
                  "filters":[], "output_sheet":"去重结果", "keep":"first", "confidence":1},
        "transform": {"task_type":"transform", "sheet":"交易", "operation":"map_values",
                      "map_rules":[{"source":{"sheet":"交易","display_header":"状态"},
                                    "mapping":{"完成":1,"待处理":0,"取消":-1}, "output_column":"状态编码"}], "confidence":1},
        "formula": {"task_type":"formula", "sheet":"交易", "column":"总额", "start_row":2, "end_row":2,
                    "mode":"set_formula", "formula":"=F2*G2", "source_row":None, "recalculate":False, "confidence":1},
        "conversion": {"task_type":"conversion", "direction":"xlsx_to_csv", "source_sheet":"交易", "output_name":"交易.csv",
                       "encoding":"utf-8-sig", "delimiter":",", "preserve_text":True, "confidence":1},
    }
    primary = copy.deepcopy(operations[kind])
    if not composite:
        return primary
    pre = {"task_type":"normalize", "sheet":"交易", "column":"编号文本", "operations":["trim","nfkc"],
           "output_column":"标准编号", "confidence":1}
    if kind == "normalize":
        second = {"task_type":"normalize", "sheet":"交易", "column":"状态文本", "operations":["trim","upper"],
                  "output_column":"标准状态", "confidence":1}
        return {"task_type":"composite", "operations":[pre, second], "confidence":1}
    if kind == "lookup":
        primary["main_key_column"] = "标准编号"
    if kind == "clean":
        primary["key_columns"] = ["标准编号"]
    if kind == "transform":
        pre = {"task_type":"transform", "sheet":"交易", "operation":"impute_numeric",
               "impute_rules":[{"source":{"sheet":"交易","display_header":"评分"}, "method":"median",
                                "constant_value":None, "output_column":"评分填补", "indicator_column":"评分缺失"}], "confidence":1}
    return {"task_type":"composite", "operations":[pre, primary], "confidence":1}


def request_text(cap: str, journey: str, number: int) -> dict[str, str]:
    instructions = {
        "lookup":"按客户编号从客户表精确补充客户等级，未匹配留空并报告，保留交易表行序",
        "normalize":"把编号文本去掉首尾空格并做 NFKC 标准化，写入新列标准编号，保留原列",
        "aggregate":"按地区汇总金额之和，结果写入新工作表地区汇总，结果列叫金额合计",
        "sort_rank":"按金额降序稳定排序并做密集排名，写入新表金额排名表，排名列叫金额排名",
        "clean":"按订单号精确去重，保留第一条，完整记录写入新表去重结果",
        "transform":"把状态列的完成、待处理、取消分别映射为 1、0、-1，新列叫状态编码并保留原列",
        "formula":"在总额列第2行到数据末行填入相对引用公式 =F2*G2，不重算公式值",
        "conversion":"把交易工作表导出为 UTF-8 BOM、逗号分隔的交易.csv，并保留原工作簿",
    }
    base = instructions[cap]
    if journey == "multi_step":
        base = ("先把编号文本去掉首尾空格并做 NFKC，新增标准编号；再" + base.replace("按客户编号", "按标准编号")
                if cap not in ("normalize", "transform") else
                ("先把编号文本去掉首尾空格并做 NFKC，新增标准编号；再把状态文本去掉首尾空格并转成大写，新增标准状态" if cap == "normalize"
                 else "先用评分中位数填补缺失值并添加缺失标记，再" + base))
        if cap == "clean":
            base = "先把编号文本去掉首尾空格并做 NFKC，新增标准编号；再按标准编号精确去重，保留第一条，完整记录写入新表去重结果"
    if journey == "clarification":
        vague = {"lookup":"从客户表补充客户信息", "normalize":"把编号列标准化", "aggregate":"汇总交易金额",
                 "sort_rank":"把交易数据排个名", "clean":"清理重复交易", "transform":"把状态转换成编码",
                 "formula":"给总额列填公式", "conversion":"把工作表导出成 CSV"}[cap]
        base = vague
    if journey == "boundary":
        base = {"lookup":"按客户名称做模糊匹配并补充信息", "normalize":"联网抓取客户名称并自动修正当前表",
                "aggregate":"写一段 VBA 宏完成汇总并直接运行", "sort_rank":"调用外部插件给交易数据智能排名",
                "clean":"执行任意 Python 代码删除重复数据", "transform":"联网抓取分类规则并改写所有单元格",
                "formula":"编写并运行 VBA 宏计算公式", "conversion":"对图片做 OCR 识别并转换成 Excel"}[cap]
    tag = f"（评测批次{number:03d}）"
    return {"concise": base + "。" + tag,
            "colloquial": "麻烦帮我处理这份表：" + base + "，其他内容别动。" + tag,
            "contextual": "我需要把处理后的文件交给同事继续使用。请" + base + "，并保持没有提到的工作表和格式不变。" + tag}


def make_book(path: Path, rows: int) -> dict[str, int]:
    workbook = Workbook(write_only=False)
    sheet = workbook.active; sheet.title = "交易"
    headers = ["订单号","客户编号","地区","金额","状态","数量","单价","总额","编号文本","状态文本","评分"]
    sheet.append(headers)
    sheet.freeze_panes = "A2"
    for index in range(1, rows + 1):
        customer = None if index % 29 == 0 else f"{(index % 80) + 1:03d}"
        sheet.append([f"ORD-{index if index % 17 else index-1:06d}", customer, ("东","西","南","北")[index % 4],
                      (index * 37) % 10000 / 10, ("完成","待处理","取消")[index % 3], index % 9 + 1,
                      (index % 23 + 1) / 2, None, None if customer is None else f"  {customer}  ",
                      f"  {('open','closed','hold')[index % 3]} ", None if index % 11 == 0 else index % 100])
    customers = workbook.create_sheet("客户"); customers.append(["客户编号","客户等级"])
    for index in range(1, 81): customers.append([f"{index:03d}", f"等级{(index-1)//20+1}"])
    sheet["A1"].font = Font(bold=True, color="224466")
    workbook.save(path); workbook.close(); stable_xlsx(path)
    return {"交易": rows, "客户": 80}


def research_sources() -> list[dict]:
    used = set()
    for path in (APP / ".phase1/cases").glob("*/case.json"):
        source = json.loads(path.read_text(encoding="utf-8")).get("source", {})
        if source.get("type") == "research_reconstructed": used.add(str(source.get("source_item_id")))
    candidates = []
    with (ROOT / "DataAnalysis/data/analysis/labels.csv").open(encoding="utf-8-sig", newline="") as stream:
        for row in csv.DictReader(stream):
            if (row.get("source_item_id") and row.get("source_url") and row["source_item_id"] not in used
                    and float(row.get("confidence") or 0) >= .7 and row.get("user_goal_zh")):
                candidates.append(row)
    candidates.sort(key=lambda row: hashlib.sha256(f"{SEED}:{row['source_item_id']}".encode()).hexdigest())
    if len(candidates) < 64: raise RuntimeError("not enough unused research sources")
    return candidates[:64]


def bench_sources() -> list[Path]:
    used = set()
    for path in (APP / ".phase1/cases").glob("*/case.json"):
        source = json.loads(path.read_text(encoding="utf-8")).get("source", {})
        if source.get("type") == "spreadsheetbench_adapted": used.add(str(source.get("task_id")))
    base = ROOT / "DataAnalysis/data/raw/spreadsheetbench/extracted/spreadsheetbench_912_v0.1/all_data_912_v0.1"
    picked = []
    for path in sorted(base.rglob("*_input.xlsx"), key=lambda p: hashlib.sha256(f"{SEED}:{p}".encode()).hexdigest()):
        task = path.parent.name
        if task not in used and task not in {p.parent.name for p in picked}: picked.append(path)
        if len(picked) == 32: break
    if len(picked) < 32: raise RuntimeError("not enough unused SpreadsheetBench groups")
    return picked


def definitions() -> list[dict]:
    cases=[]; serial=0
    for cap, counts in MATRIX.items():
        for journey, count in zip(JOURNEYS, counts):
            for _ in range(count):
                serial += 1
                cases.append({"serial":serial,"cap":cap,"journey":journey})
    random.Random(SEED).shuffle(cases)
    source_slots = ["research_reconstructed"]*64 + ["spreadsheetbench_adapted"]*32 + ["synthetic"]*40
    random.Random(SEED+1).shuffle(source_slots)
    sizes = ["small"]*80 + ["medium"]*56
    random.Random(SEED+2).shuffle(sizes)
    for item, source_type, size in zip(cases[:136], source_slots, sizes): item.update(source_type=source_type,size=size)
    for item in cases[136:]: item.update(source_type="synthetic",size="large")
    cases.sort(key=lambda x:(list(MATRIX).index(x["cap"]), JOURNEYS.index(x["journey"]), x["serial"]))
    counters=defaultdict(int)
    for item in cases:
        counters[item["cap"]]+=1; item["number"]=counters[item["cap"]]
    return cases


def stability_ids(items: list[dict]) -> set[tuple[str,str,int]]:
    allocation={
      "lookup":(2,1,1,1), "normalize":(3,1,1,0), "aggregate":(2,1,1,1), "sort_rank":(3,1,1,0),
      "clean":(2,1,1,1), "transform":(2,2,1,0), "formula":(3,1,1,0), "conversion":(3,0,1,1)}
    chosen=set()
    for cap, counts in allocation.items():
        for journey,count in zip(JOURNEYS,counts):
            pool=[x for x in items if x["cap"]==cap and x["journey"]==journey]
            for item in pool[:count]: chosen.add((cap,journey,item["number"]))
    return chosen


def clarification(spec: dict, rounds: int) -> tuple[dict,list[dict]]:
    initial=copy.deepcopy(spec); scripts=[]
    operations=initial["operations"] if initial.get("task_type")=="composite" else [initial]
    target=operations[-1]
    candidates={"lookup":[("return_columns",[],"返回客户等级，输出列名为客户等级。")],
                "normalize":[("operations",[],"去掉首尾空格并做 NFKC 标准化。")],
                "aggregate":[("group_by",[],"按地区分组。")],
                "sort_rank":[("rank_column",None,"按金额降序排序并按金额做密集排名。")],
                "clean":[("key_columns",[],"按订单号去重。")],
                "transform":[("sheet","","处理交易工作表，完成、待处理、取消分别映射成1、0、-1。")],
                "formula":[("formula",None,"公式是=F2*G2。")],
                "conversion":[("source_sheet","","导出交易工作表。")]}[target["task_type"]]
    if rounds==2:
        if target["task_type"] == "conversion": candidates.append(("output_name","","文件名用交易.csv。"))
        elif target["task_type"] == "lookup": candidates.append(("main_sheet","","主表是交易工作表。"))
        elif target["task_type"] == "normalize": candidates.append(("column","","处理编号文本列。"))
        else: candidates.append(("sheet","","处理交易工作表。"))
    current=copy.deepcopy(initial)
    current_target=current["operations"][-1] if current.get("task_type")=="composite" else current
    missing=[]
    for field,blank,_ in candidates:
        current_target[field]=blank; missing.append(field)
        if target["task_type"]=="lookup" and field=="return_columns": current_target["output_column_names"]=[]
    current_target["missing_fields"]=list(missing)
    for field,_,answer in candidates:
        final_target=(spec["operations"][-1] if spec.get("task_type")=="composite" else spec)
        current_target[field]=copy.deepcopy(final_target.get(field,1))
        if target["task_type"]=="lookup" and field=="return_columns": current_target["output_column_names"]=copy.deepcopy(final_target["output_column_names"])
        missing.remove(field)
        if missing: current_target["missing_fields"]=list(missing)
        else: current_target.pop("missing_fields",None)
        scripts.append({"field":field,"answer":answer,"spec_after":copy.deepcopy(current)})
    return initial if False else copy.deepcopy(scripts[0]["spec_after"] if False else initial), scripts


def _initial_from_scripts(spec: dict, scripts: list[dict]) -> dict:
    initial=copy.deepcopy(spec); target=initial["operations"][-1] if initial.get("task_type")=="composite" else initial
    target["missing_fields"]=[s["field"] for s in scripts]
    for step in scripts:
        value=(step["spec_after"]["operations"][-1] if step["spec_after"].get("task_type")=="composite" else step["spec_after"])[step["field"]]
        target[step["field"]]=(
            None if step["field"] == "formula"
            else [] if isinstance(value,list)
            else "" if isinstance(value,str)
            else None
        )
        if target["task_type"] == "sort_rank" and step["field"] == "rank_column":
            target["sort_keys"] = []
        if target["task_type"]=="lookup" and step["field"]=="return_columns": target["output_column_names"]=[]
    return initial


def _suite_hash(cases_dir: Path) -> tuple[list[dict],str]:
    files=[]
    for path in sorted(p for p in cases_dir.rglob("*") if p.is_file()):
        files.append({"path":path.relative_to(cases_dir.parent).as_posix(),"sha256":sha(path),"bytes":path.stat().st_size})
    return files, hashlib.sha256("".join(f"{x['path']}:{x['sha256']}\n" for x in files).encode()).hexdigest()


def prepare(out: Path = DEFAULT_OUT) -> Path:
    out=protected_output(out); cases_dir=out/"cases"
    if (out/"manifest.json").is_file():
        manifest=json.loads((out/"manifest.json").read_text(encoding="utf-8")); _,actual=_suite_hash(cases_dir)
        if actual==manifest["suite_sha256"]:
            _write_pricing_snapshot(out)
            return out
        raise RuntimeError("locked phase-two suite differs from manifest; refusing to overwrite")
    if cases_dir.exists() and any(cases_dir.iterdir()): raise RuntimeError("output cases directory is not empty")
    cases_dir.mkdir(parents=True,exist_ok=True)
    research=iter(research_sources()); bench=iter(bench_sources()); items=definitions(); stable=stability_ids(items)
    size_count=Counter(); source_count=Counter(); journey_count=Counter(); capability_count=Counter(); clarify_seen=0
    for index,item in enumerate(items,1):
        cap=item["cap"]; journey=item["journey"]; number=item["number"]
        case_id=f"P2-{cap.upper()}-{number:03d}"; folder=cases_dir/case_id;folder.mkdir()
        rows=ROW_SIZES[item["size"]][index % len(ROW_SIZES[item["size"]])]
        input_path=folder/"input.xlsx"; row_counts=make_book(input_path,rows)
        if item["source_type"]=="research_reconstructed":
            row=next(research);source={"type":item["source_type"],"source_item_id":row["source_item_id"],"source_name":row["source"],
                "source_url":row["source_url"],"original_goal":row["user_goal_zh"],"adaptation_note":"保留公开讨论中的任务动机，使用无隐私的确定性工作簿重建可测场景。"}
        elif item["source_type"]=="spreadsheetbench_adapted":
            original=next(bench);source={"type":item["source_type"],"task_id":original.parent.name,
                "original_file":original.relative_to(ROOT).as_posix(),"original_sha256":sha(original),
                "adaptation_note":"保留公开基准的表格编辑任务形式，重建为 SheetPilot 当前支持且可独立验算的场景。"}
        else: source={"type":"synthetic","scenario":"订单运营","adaptation_note":"按用户调研中的高频表格任务构造。"}
        requests=request_text(cap,journey,index); final_spec=None if journey=="boundary" else op(cap,journey=="multi_step")
        if final_spec and cap=="formula":
            target=final_spec["operations"][-1] if final_spec.get("task_type")=="composite" else final_spec;target["end_row"]=rows+1
        if final_spec and journey=="multi_step" and cap=="transform":
            # The request asks for an imputed value and a missing marker but does
            # not name those columns.  Let the product's documented defaults be
            # the expected result.
            impute=final_spec["operations"][0]["impute_rules"][0]
            impute["output_column"]="";impute["indicator_column"]=""
        if final_spec and journey=="clarification":
            target=final_spec["operations"][-1] if final_spec.get("task_type")=="composite" else final_spec
            task_type=target["task_type"]
            if task_type=="normalize": target["output_column"]=""
            elif task_type=="aggregate":
                target["output_sheet"]="SheetPilot_Summary"
                for metric in target["metrics"]: metric["output_column"]=""
            elif task_type=="sort_rank":
                target["output_sheet"]="SheetPilot_Sorted";target["rank_output_column"]="排名"
            elif task_type=="clean": target["output_sheet"]="SheetPilot_Cleaned"
            elif task_type=="transform":
                for rule in target.get("map_rules",[]): rule["output_column"]=""
            elif task_type=="formula": target["recalculate"]=True
            elif task_type=="conversion":
                target["output_name"]="";target["encoding"]="auto";target["delimiter"]="auto"
        initial=None;scripts=[]
        if journey=="clarification":
            clarify_seen+=1;rounds=2 if clarify_seen<=8 else 1
            _,scripts=clarification(final_spec,rounds);initial=_initial_from_scripts(final_spec,scripts)
        oracle_name=oracle_hash=None
        if final_spec:
            expected=reference(input_path,resolve_oracle_defaults(final_spec));oracle_name="oracle.json.gz";dump_gzip(folder/oracle_name,{"workbook":expected["workbook"],"csv_artifacts":expected["csv_artifacts"],"tolerance":1e-9});oracle_hash=sha(folder/oracle_name)
        case=Case(id=case_id,journey=journey,primary_capability=cap,
            capabilities=(["normalize",cap] if journey=="multi_step" and cap!="normalize" else [cap]), source=source,
            family_id=f"{cap}-{journey}",size_class=item["size"],row_counts=row_counts,input=input_path.name,input_sha256=sha(input_path),
            requests=requests,expected_action="refuse" if journey=="boundary" else "clarify_then_execute" if journey=="clarification" else "execute",
            spec=final_spec,initial_spec=initial,clarification_script=scripts,
            allowed_changes=[] if journey=="boundary" else ["声明的新增列、结果工作表或导出文件"],
            forbidden_changes=["原始文件","未声明的单元格","原有样式与布局"],oracle=oracle_name,oracle_sha256=oracle_hash,
            stability_selected=(cap,journey,number) in stable)
        dump(folder/"case.json",case.model_dump(mode="json"))
        size_count[item["size"]]+=1;source_count[item["source_type"]]+=1;journey_count[journey]+=1;capability_count[cap]+=1
        if index % 10 == 0 or index == len(items): print(f"generated {index}/{len(items)}",flush=True)
    inventory={"version":"2.1","seed":SEED,"count":160,"journeys":dict(journey_count),"capabilities":dict(capability_count),
               "sizes":dict(size_count),"sources":dict(source_count),"stability_count":sum(c.stability_selected for c in load_cases(cases_dir))}
    dump(out/"inventory.json",inventory)
    lines=["# SheetPilot 第二阶段锁定评测集","",f"Seed：`{SEED}`；Case：160；稳定性子集：40。","",
           "|能力|直接|多步骤|澄清|边界|合计|","|---|---:|---:|---:|---:|---:|"]
    for cap,counts in MATRIX.items():lines.append(f"|{CAP_ZH[cap]}|{counts[0]}|{counts[1]}|{counts[2]}|{counts[3]}|{sum(counts)}|")
    (out/"case-coverage.md").write_text("\n".join(lines)+"\n",encoding="utf-8")
    files,suite_hash=_suite_hash(cases_dir);dump(out/"manifest.json",{"version":"2.1","seed":SEED,"suite_sha256":suite_hash,"files":files})
    _write_pricing_snapshot(out)
    return out

