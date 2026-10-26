"""Additional phase-one cases. Public examples are always labelled as adaptations."""
from __future__ import annotations

import csv
import json
import random
import re
import shutil
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .common import ROOT, Case, dump, sha
from .oracle import reference

BENCH = ROOT / 'DataAnalysis/data/raw/spreadsheetbench/extracted/spreadsheetbench_912_v0.1/all_data_912_v0.1'
LABELS = ROOT / 'DataAnalysis/data/analysis/labels.csv'
DOMAINS = [
    ('订单','客户','订单号','客户编号','地区','金额','状态','数量','单价','总额','编号文本','分类','尺寸','BMI','等级'),
    ('工单','人员','工单号','人员编号','部门','工时','状态','数量','单价','合计','原始编号','类别','规格','评分','岗位'),
    ('库存','供应商','批次号','供应商编号','仓库','库存量','状态','数量','单价','总量','原始编号','类别','规格','评分','级别'),
    ('发票','客户','发票号','客户编号','城市','金额','状态','数量','单价','总额','原始编号','类别','尺寸','评分','等级'),
    ('物流','承运商','运单号','承运商编号','区域','运费','状态','数量','单价','总额','原始编号','类别','尺寸','评分','级别'),
    ('课程','学生','课程号','学生编号','校区','分数','状态','数量','单价','总分','原始编号','类别','尺寸','评分','年级'),
    ('交易','账户','交易号','账户编号','区域','金额','状态','数量','单价','总额','原始编号','类别','尺寸','评分','风险级别'),
    ('门店','区域表','门店号','区域编号','城市','销售额','状态','数量','单价','总额','原始编号','类别','尺寸','评分','层级'),
]


def book(path: Path, domain: tuple, seed: int, rows: int = 12, change: str = '') -> None:
    rng=random.Random(seed)
    main,other,ident,key,area,amount,state,qty,price,total,raw,category,dimension,score,grade=domain
    wb=Workbook();ws=wb.active;ws.title=main
    ws.append([ident,key,area,amount,state,qty,price,total,raw,category,dimension,score,'长度数值_mm','宽度数值_mm','高度数值_mm'])
    for i in range(rows):
        code=None if i==rows-1 else f'{(i%7)+1:03d}'
        a=[10,20,20,5,40,0,15][i%7]+rng.randint(0,3)
        ws.append([f'{ident[:1]}{i+1}',code,['东','西','北'][i%3],a,
                   ['完成','待处理','取消'][i%3],i%5+1,5,None,
                   None if code is None else f' {code} '+('-X' if change=='suffix' else ''),['有','无'][i%2],
                   f'{10+i%4}*20*30 mm' if change=='dimension_clean' else None if i%11==10 else '无法识别' if i%11==9 else f'{10+i%4}*20*30 mm',
                   None if i%5==0 else 10+i%7,
                   10+i%4 if change=='derive_clean' else None if i%9==8 else 10+i%4,
                   20,30 if change=='derive_clean' else 0 if i%10==9 else 30])
    if change=='duplicate':ws.cell(4,2,ws.cell(2,2).value)
    if change in ('formula_copy','formula_gaps'):ws.cell(2,8,'=F2*G2')
    if change=='formula_gaps':ws.cell(13,8,'=F13*G13')
    src=wb.create_sheet(other);src.append([key,grade,area])
    for i in range(1,9):src.append([f'{i:03d}',f'{grade}{i}',['东','西','北'][i%3]])
    if change=='source_duplicate':src.append(['001','末条','南'])
    ws.freeze_panes='A2'
    wb.save(path);wb.close()
    from .cases import stable_xlsx
    stable_xlsx(path)


def op(cap: str, d: tuple, variant: int):
    main,other,_,key,area,amount,state,qty,price,total,raw,category,dimension,score,grade=d
    if cap=='lookup':
        return {'task_type':'lookup','main_sheet':main,'source_sheet':other,'main_key_column':key,
                'source_key_column':key,'return_columns':[grade],'output_column_names':[f'匹配{grade}'],
                'duplicate_policy':'last' if variant%9==8 else 'block','confidence':1}
    if cap=='normalize':
        operations=[['trim','nfkc'],['trim'],['trim','upper'],['fill_blank'],['strip_suffix']][variant%5]
        return {'task_type':'normalize','sheet':main,'column':raw,'operations':operations,
                'fill_value':'缺号' if variant%5==3 else None,
                'strip_suffixes':['-X'] if variant%5==4 else [],'output_column':'标准编号','confidence':1}
    if cap=='aggregate':
        fun=['sum','count','mean','min','max'][variant%5]
        return {'task_type':'aggregate','sheet':main,'group_by':[area],
                'metrics':[{'column':amount,'function':fun,'output_column':'统计值'}],
                'filters':[{'column':state,'operator':'eq','value':'完成'}] if variant%2 else [],
                'output_sheet':'统计结果','confidence':1}
    if cap=='sort_rank':
        return {'task_type':'sort_rank','sheet':main,'sort_keys':[{'column':amount,'direction':'asc' if variant%2 else 'desc'}],
                'rank_column':amount if variant%3==0 else None,'rank_method':['dense','min','max'][variant%9//3],'rank_output_column':'名次',
                'rank_direction':'desc','top_n':5 if variant%7==0 else None,'output_sheet':'排序结果','confidence':1}
    if cap=='clean':
        if variant%5==0:
            return {'task_type':'clean','sheet':main,'operation':'filter_rows','filters':[{'column':amount,'operator':'gt','value':10}],
                    'output_sheet':'清理结果','confidence':1}
        if variant%5==4:
            return {'task_type':'clean','sheet':main,'operation':'keep_columns',
                    'keep_columns':[{'sheet':main,'display_header':key},{'sheet':main,'display_header':amount}],
                    'output_sheet':'清理结果','confidence':1}
        return {'task_type':'clean','sheet':main,'operation':'mark_duplicates' if variant%5==3 else 'remove_duplicates',
                'key_columns':[key],'keep':'last' if variant%5==2 else 'first',
                'output_sheet':'清理结果','confidence':1}
    if cap=='formula':
        mode=['set_formula','copy_formula','fill_gaps'][variant%3]
        return {'task_type':'formula','sheet':main,'column':total,'start_row':3 if mode=='copy_formula' else 2,
                'end_row':13 if mode=='fill_gaps' or variant%4 else 9,'mode':mode,'formula':'=F2*G2' if mode=='set_formula' else None,
                'source_row':2 if mode=='copy_formula' else None,'recalculate':False,'confidence':1}
    if cap=='conversion':
        if variant%2==0:return {'task_type':'conversion','direction':'csv_to_xlsx','encoding':'utf-8-sig','delimiter':',','preserve_text':True,'confidence':1}
        return {'task_type':'conversion','direction':'xlsx_to_csv','source_sheet':main,'output_name':'导出.csv','encoding':'utf-8-sig','confidence':1}
    if cap=='transform':
        if variant%4==1:
            return {'task_type':'transform','sheet':main,'operation':'impute_numeric',
                    'impute_rules':[{'source':{'sheet':main,'display_header':score},'method':['mean','median','constant'][variant%3],
                                     'constant_value':17 if variant%3==2 else None,'output_column':'填补值','indicator_column':'缺失标记'}],
                    'confidence':1}
        if variant%4==2:
            return {'task_type':'transform','sheet':main,'operation':'parse_dimensions',
                    'dimension_rule':{'source':{'sheet':main,'display_header':dimension},'default_unit':'mm',
                                      'length_column':'长_mm','width_column':'宽_mm','height_column':'高_mm',
                                      'max_dimension_column':'最大径_mm','status_column':'解析状态'},'confidence':1}
        if variant%4==3:
            return {'task_type':'transform','sheet':main,'operation':'derive_metric',
                    'derive_rule':{'sources':[{'sheet':main,'display_header':'长度数值_mm'},
                                              {'sheet':main,'display_header':'宽度数值_mm'},{'sheet':main,'display_header':'高度数值_mm'}],
                                   'formula_template':'rectangular_product','output_column':'派生体积_cm3',
                                   'status_column':'派生状态'},'confidence':1}
        return {'task_type':'transform','sheet':main,'operation':'map_values',
                'map_rules':[{'source':{'sheet':main,'display_header':category},'mapping':{'有':1,'无':0},'output_column':'类别编码'}],
                'confidence':1}
    raise ValueError(cap)


def request_for(cap: str, spec: dict, d: tuple) -> str:
    main,other,_,key,area,amount,state,qty,price,total,raw,category,dimension,score,grade=d
    if cap=='lookup':return f'在{main}按{key}精确匹配{other}，把{grade}写入新列匹配{grade}；找不到留空并报告，来源重复{ "取最后一条" if spec["duplicate_policy"]=="last" else "时停止"}。'
    if cap=='normalize':
        descriptions={'trim':'去首尾空格','nfkc':'做NFKC标准化','upper':'转大写','fill_blank':'把空值填成缺号','strip_suffix':'去掉末尾-X'}
        return f'把{main}的{raw}列'+ '并'.join(descriptions[x] for x in spec['operations'])+'，写入新列标准编号；原列不改。'
    if cap=='aggregate':
        method={'sum':'求和','count':'非空计数','mean':'求平均值','min':'取最小值','max':'取最大值'}[spec['metrics'][0]['function']]
        return f'在{main}'+(f'先筛选{state}等于完成，再' if spec['filters'] else '')+f'按{area}分组，对{amount}{method}，新表统计结果的结果列叫统计值。'
    if cap=='sort_rank':
        method={'dense':'密集排名','min':'并列取最小名次','max':'并列取最大名次'}[spec['rank_method']]
        return f'{main}按{amount}{"升" if spec["sort_keys"][0]["direction"]=="asc" else "降"}序稳定排序'+(f'，按{amount}{"升" if spec["rank_direction"]=="asc" else "降"}序排名，同{amount}做{method}，排名列叫名次' if spec['rank_column'] else '')+(f'，只保留前{spec["top_n"]}名' if spec['top_n'] and spec['rank_column'] else f'，只保留前{spec["top_n"]}条' if spec['top_n'] else '')+'，写入新表排序结果。'
    if cap=='clean':
        if spec['operation']=='filter_rows':return f'将{main}中{amount}大于10的行筛到新表清理结果，保留原表。'
        if spec['operation']=='keep_columns':return f'{main}只保留{key}和{amount}两列，按原行序写入新表清理结果，原表不动。'
        if spec['operation']=='mark_duplicates':return f'{main}按{key}标记重复值，重复的所有行标“重复”，写入新表清理结果。'
        return f'{main}按{key}精确去重，保留{ "最后" if spec["keep"]=="last" else "最先"}一条，写入新表清理结果。'
    if cap=='formula':return f'在{main}的{total}列第{spec["start_row"]}到{spec["end_row"]}行'+('按相对引用填入=F2*G2' if spec['mode']=='set_formula' else '以第2行为样板复制公式' if spec['mode']=='copy_formula' else '修复公式空缺')+'；不覆盖非空单元格，只检查公式文本。'
    if cap=='conversion':return '把这个CSV转为XLSX，UTF-8 BOM、逗号分隔，所有字段保留文本和前导零。' if spec['direction']=='csv_to_xlsx' else f'把{main}工作表导出成UTF-8 BOM的导出.csv，保留原工作簿。'
    if cap=='transform':
        if spec['operation']=='impute_numeric':
            method={'mean':'有效值平均数','median':'有效值中位数','constant':'固定值17'}[spec['impute_rules'][0]['method']]
            return f'把{main}的{score}空值按{method}填补，新列填补值和缺失标记，保留原列。'
        if spec['operation']=='parse_dimensions':return f'解析{main}的{dimension}列，单位毫米，新增列叫长_mm、宽_mm、高_mm、最大径_mm，状态列叫解析状态；无法识别的值留空并报告。'
        if spec['operation']=='derive_metric':return f'在{main}使用长方体公式，按长度数值_mm×宽度数值_mm×高度数值_mm÷1000计算派生体积_cm3，状态列叫派生状态；无效输入留空并报告。'
        return f'把{main}的{category}中值“有”映射为1、值“无”映射为0，新列类别编码，不输出状态列，保留原列。'
    raise ValueError(cap)


def incomplete_request(cap: str, spec: dict, d: tuple) -> str:
    main,other,_,key,area,amount,state,qty,price,total,raw,category,dimension,score,grade=d
    if cap=='aggregate':
        method={'sum':'求和','count':'非空计数','mean':'求平均值','min':'取最小值','max':'取最大值'}[spec['metrics'][0]['function']]
        return f'在{main}对{amount}{method}，结果列统计值，写入新表统计结果；分组列我稍后说明。'
    return {
        'lookup':f'在{main}按{key}从{other}精确匹配一项信息，追加新列；要返回哪一列我稍后说明。找不到留空并报告。',
        'normalize':f'在{main}把一列文字去首尾空格，追加新列标准编号；哪一列我稍后说明。',
        'sort_rank':f'把{main}的完整记录排序到新表排序结果；按哪列、升序还是降序我稍后说明。',
        'clean':f'把{main}去重，保留第一条，写入新表清理结果；按哪列去重我稍后说明。',
        'formula':f'在{main}的{total}第2至9行填公式，不覆盖已有值；具体公式我稍后说明。',
        'conversion':'把工作簿中的一个工作表导出UTF-8 BOM CSV，文件名导出.csv；导出哪张工作表我稍后说明。',
        'transform':f'把一个工作表中的{score}空值按均值填补，新增填补值和缺失标记；是哪张表我稍后说明。',
    }[cap]


BENCH_JOBS = [
    ('19-16','clean','Raw Data','Task Status','eq','TSCO REQ'),
    ('99-24','clean','Paid','Date Paid','not_blank',None),
    ('353-6','clean','Data','communication_reason','eq','HUB Notes'),
    ('178-22','clean','Sheet1','ITEM','eq','TELIVISION'),
    ('212-40','clean','Sheet1','ID','dedupe',None),
    ('427-35','clean','Sheet2','Identifier','dedupe',None),
    ('262-17','sort_rank','Sheet1','Task','sort','Responsibility'),
    ('122-39','sort_rank','OUTPUT','NAME','sort','DEBIT'),
    ('17251','sort_rank','Sheet1','price','sort',None),
    ('336-1','sort_rank','Sheet1','ID','sort','SCORE'),
    ('289-3','normalize','MAIN','CODE','trim',None),
    ('225-4','normalize','Statement','Date','strip_suffix',None),
]


def research_rows():
    with LABELS.open(encoding='utf-8-sig',newline='') as f:rows=list(csv.DictReader(f))
    by_id={int(r['item_id']):r for r in rows}
    # Curated after reading the original goals. A case keeps only the stated
    # operation that the current product can perform; metadata says so.
    selected=[]
    groups={
        'lookup':[355,386,391,469,613,614,732,778],
        'normalize':[242,250,443,636,689,892,1294,1300],
        'aggregate':[316,346,349,352,466,635,695,912],
        'clean':[351,360,624,659,746,944,1119,1234],
        'conversion':[236,245,255,450,852,887,933,967],
        'boundary':[239,1168,889,330],
    }
    for cap,ids in groups.items():
        for ident in ids:selected.append((cap,by_id[ident]))
    benchmark=json.loads((BENCH/'dataset.json').read_text(encoding='utf-8'))
    bench_by_id={str(r['id']):r for r in benchmark}
    for ident in ('122-39','208-20','330-23','17251'):
        task=bench_by_id[ident]
        selected.append(('sort_rank',{'item_id':f'bench-{ident}','source':'spreadsheetbench',
            'source_item_id':f'spreadsheetbench/{ident}',
            'source_url':'https://huggingface.co/datasets/KAKA22/SpreadsheetBench',
            'user_goal_zh':task['instruction'],'evidence_excerpt':task['instruction'][:400]}))
    assert len(selected)==48 and len({r['source_item_id'] for _,r in selected})==48
    return selected


def prepare_expansion(directory: Path) -> None:
    from sheetpilot.schemas import TaskEnvelope
    from .cases import stable_xlsx
    definitions=[]
    def add(ident,cap,kind,domain,variant,*,provenance=None,bench=None,research=None,change=''):
        spec=op(cap,domain,variant) if cap!='boundary' else op('lookup',domain,variant)
        req=request_for(cap,spec,domain) if cap!='boundary' else ''
        entry=dict(id=ident,capability=cap,kind=kind,domain=domain,variant=variant,spec=spec,request=req,
                   source=provenance,bench=bench,research=research,change=change)
        definitions.append(entry)
        return entry

    # 160 normal cases: 20 for each capability, 24 of them with public original files.
    for cap in ('lookup','normalize','aggregate','sort_rank','clean','formula','conversion','transform'):
        for i in range(20):
            entry=add(f'P-{cap.upper()}-{i+1:02d}',cap,'normal',DOMAINS[i%8],i,
                      change='formula_copy' if cap=='formula' and i%3==1 else 'formula_gaps' if cap=='formula' and i%3==2 else 'source_duplicate' if cap=='lookup' and i%9==8 else 'suffix' if cap=='normalize' and i%5==4 else 'dimension_clean' if cap=='transform' and i%4==2 else 'derive_clean' if cap=='transform' and i%4==3 else 'duplicate' if cap=='clean' else '')
    for j,(task,cap,sheet,col,action,value) in enumerate(BENCH_JOBS):
        for variant in (1,2):
            # Each capability uses eight of its twenty normal slots.
            slot=(sum(x[1]==cap for x in BENCH_JOBS[:j]))*2+variant-1
            entry=definitions[[e['id'] for e in definitions].index(f'P-{cap.upper()}-{slot+1:02d}')]
            entry['bench']=(task,variant,sheet,col,action,value)
    # 40 clarification cases: each has one specified missing field or a revision.
    for i in range(40):
        cap=('lookup','normalize','aggregate','sort_rank','clean','formula','conversion','transform')[i%8]
        e=add(f'P-Q-{i+1:02d}',cap,'clarification',DOMAINS[(i+3)%8],i)
        field={'lookup':'return_columns','normalize':'column','aggregate':'group_by','sort_rank':'sort_keys',
               'clean':'key_columns','formula':'formula','conversion':'source_sheet','transform':'sheet'}[cap]
        e['missing_field']=field
        e['reply']={'field':field,'answer':request_for(cap,e['spec'],e['domain'])}
        if cap=='clean':e['variant']=1;e['spec']=op(cap,e['domain'],1)
        if cap=='conversion':e['variant']=1;e['spec']=op(cap,e['domain'],1)
        if cap=='formula':e['variant']=0;e['spec']=op(cap,e['domain'],0)
        if cap=='transform':e['variant']=1;e['spec']=op(cap,e['domain'],1)
        e['request']=incomplete_request(cap,e['spec'],e['domain'])
    # 20 unsupported requests and 20 controlled failures.
    refusals=['请运行VBA宏修改工作簿。','请执行任意Python代码清理数据。','按近似拼写模糊匹配客户名称。','合并两个不同工作簿中的数据。',
              '请联网抓取私有系统里的记录。','修改受密码保护的工作簿。','在Excel里运行自定义插件。','连接外部数据库同步资料。',
              '使用Power Query脚本跨文件汇总。','调用自定义VBA函数计算。','按语义相似度自动关联客户。','把多个文件夹里的工作簿批量合并。',
              '运行PowerShell脚本改写文件。','打开网页并采集帐号数据。','修改外部链接指向的文件。','执行JavaScript宏处理单元格。',
              '直接修改另一个用户的在线表格。','自动下载远程文件并合并。','对照片进行OCR后填入表格。','在Excel里执行自定义加载项。']
    for i,req in enumerate(refusals):
        e=add(f'P-R-{i+1:02d}','boundary','refusal',DOMAINS[i%8],i);e['request']=req
    for i in range(20):
        e=add(f'P-E-{i+1:02d}','lookup','fault',DOMAINS[i%8],i)
        e['fault']=('timeout','invalid_response','step_failure','stale_plan')[i%4]
        if e['fault']=='step_failure':
            first=op('normalize',e['domain'],i)
            second=op('lookup',e['domain'],i);second['main_key_column']='标准编号'
            e['spec']={'task_type':'composite','operations':[first,second],'confidence':1}
            e['request']=request_for('normalize',first,e['domain'])+'然后按标准编号匹配并补充等级，未匹配留空并报告。'

    # One dependent two-step task in each capability family.
    for e in definitions:
        if e['kind']!='normal' or e['variant']!=17 or e['bench'] or e['capability']=='conversion':continue
        cap=e['capability'];domain=e['domain'];first=op('normalize',domain,0)
        second=e['spec']
        if cap=='lookup':second['main_key_column']='标准编号'
        if cap=='clean' and second['operation']=='remove_duplicates':second['key_columns']=['标准编号']
        if cap=='normalize':
            first={'task_type':'clean','sheet':domain[0],'operation':'filter_rows',
                   'filters':[{'column':domain[5],'operator':'gt','value':10}],
                   'output_sheet':'预筛选','confidence':1}
            second['sheet']='预筛选'
            e['request']=f'先把{domain[0]}中{domain[5]}大于10的记录筛到预筛选，再把预筛选的{domain[10]}去空格并转大写，新增标准编号。'
        else:
            e['request']=request_for('normalize',first,domain)+'然后'+request_for(cap,second,domain)
            if cap=='lookup':e['request']=request_for('normalize',first,domain)+f'然后用标准编号匹配{domain[1]}的{domain[3]}，新增匹配{domain[14]}，来源重复取最后一条，未匹配留空并报告。'
            if cap=='clean':e['request']=request_for('normalize',first,domain)+f'然后按标准编号去重，保留最后一条，写入新表清理结果。'
        e['spec']={'task_type':'composite','operations':[first,second],'confidence':1}

    # Eight normal and three clarification examples per supported research theme.
    chosen=research_rows();targets=[]
    for cap in ('lookup','normalize','aggregate','clean','conversion'):
        targets.extend([e for e in definitions if e['kind']=='normal' and e['capability']==cap and not e['bench']][:6])
        targets.extend([e for e in definitions if e['kind']=='clarification' and e['capability']==cap][:2])
    targets.extend([e for e in definitions if e['kind']=='refusal'][:4])
    targets.extend([e for e in definitions if e['kind']=='normal' and e['capability']=='sort_rank' and not e['bench']][:3])
    targets.extend([e for e in definitions if e['kind']=='clarification' and e['capability']=='sort_rank'][:1])
    assert len(chosen)==len(targets)==48
    for (research_cap,row),entry in zip(chosen,targets):
        assert research_cap==entry['capability']
        entry['research']=row
    assert len(definitions)==240

    metadata=json.loads((BENCH/'dataset.json').read_text(encoding='utf-8'))
    lookup_meta={str(x['id']):x for x in metadata}
    for e in definitions:
        folder=directory/e['id'];folder.mkdir(parents=True,exist_ok=True)
        source=folder/('input.csv' if e['capability']=='conversion' and e['variant']%2==0 and not e['bench'] else 'input.xlsx')
        provenance={'type':'synthetic','generator':'phase1-expansion-v1','license':'project-generated'}
        spec=e['spec'];request=e['request'];group=f'generated-{e["domain"][0]}'
        if e['bench']:
            task,variant,sheet,col,action,value=e['bench'];original=BENCH/f'spreadsheet/{task}/{variant}_{task}_input.xlsx'
            if not original.is_file():raise FileNotFoundError(original)
            shutil.copyfile(original,source)
            if e['capability']=='clean':
                if action=='dedupe':
                    spec={'task_type':'clean','sheet':sheet,'operation':'remove_duplicates','key_columns':[col],
                          'keep':'first','output_sheet':'去重结果','confidence':1}
                    request=f'在{sheet}按{col}列去重，保留第一条，完整记录写入新表去重结果，原表不改。'
                else:
                    spec={'task_type':'clean','sheet':sheet,'operation':'filter_rows',
                          'filters':[{'column':col,'operator':action,'value':value}],
                          'output_sheet':'筛选结果','confidence':1}
                    criterion=f'等于{value}' if action=='eq' else '不为空'
                    request=f'在{sheet}筛选{col}{criterion}的记录，所有原列写入新表筛选结果，保留原表。'
            elif e['capability']=='normalize':
                spec={'task_type':'normalize','sheet':sheet,'column':col,'operations':[action],
                      'strip_suffixes':[' K1',' P1'] if action=='strip_suffix' else [],
                      'output_column':'整理结果','confidence':1}
                request=f'将{sheet}的{col}去掉末尾的 K1或 P1，结果写入新列整理结果，保留原列。' if action=='strip_suffix' else f'将{sheet}的{col}去掉首尾空格，结果写入新列整理结果，保留原列。'
            else:
                keys=[{'column':col,'direction':'asc'}]
                if value:keys.append({'column':value,'direction':'desc' if task in ('122-39','336-1') else 'asc'})
                spec={'task_type':'sort_rank','sheet':sheet,'sort_keys':keys,
                      'output_sheet':'排序结果','confidence':1}
                request=f'将{sheet}按{col}升序'+(f'、{value}{"降" if keys[1]["direction"]=="desc" else "升"}序' if value else '')+'稳定排序，完整记录写入新表排序结果，原表不动。'
            provenance={'type':'spreadsheetbench_adapted','task_id':task,'variant':variant,
                        'original_file':str(original.relative_to(ROOT)),'original_sha256':sha(original),
                        'original_instruction':lookup_meta[task]['instruction'],'adapted_instruction':request,
                        'reason':f'只测试原问题涉及的{ {"clean":"筛选或去重","sort_rank":"排序","normalize":"文本整理"}[e["capability"]]}；原问题中的宏、指定单元格或其他步骤不在本例评分内',
                        'license':'CC-BY-SA-4.0','source_url':'https://huggingface.co/datasets/KAKA22/SpreadsheetBench',
                        'benchmark_score':False}
            group=f'bench-{task}'
        elif source.suffix=='.csv':
            with source.open('w',encoding='utf-8-sig',newline='') as f:
                writer=csv.writer(f);writer.writerow(['编号','日期','金额'])
                for j in range(1,13):writer.writerow([f'{j:03d}',f'9/{j}/2022',str(j*10)])
        else:
            row_count=200 if e['kind']=='normal' and e['variant']==18 else 500 if e['kind']=='normal' and e['variant']==19 else 12+(e['variant']%4)*8
            book(source,e['domain'],20261019+int(e['id'].split('-')[-1])+sum(map(ord,e['id'])),row_count,e['change'])
        if e['research']:
            row=e['research'];provenance={'type':'research_reconstructed','research_item_id':row['item_id'],
                'source_name':row['source'],'source_item_id':row['source_item_id'],'source_url':row['source_url'],
                'original_goal':row['user_goal_zh'],'original_excerpt':row['evidence_excerpt'],
                'adapted_instruction':request,'reason':'从公开问题抽取表格处理或边界需求，另建可复现表格；没有原用户文件',
                'license':'source-linked'}
            group=f'research-{row["source_item_id"]}'
        initial=None;reply=e.get('reply');revision=e.get('revision')
        if e['kind']=='clarification':
            initial=json.loads(json.dumps(spec));field=e['missing_field']
            if field in initial:initial[field]=[] if isinstance(initial[field],list) else None if initial[field] is None else ''
            if field=='return_columns':initial['output_column_names']=[]
            initial['missing_fields']=[field]
            reply={'field':field,'answer':request_for(e['capability'],spec,e['domain'])}
        full=TaskEnvelope(goal=request,spec=spec).spec.model_dump(mode='json')
        if initial:initial=TaskEnvelope(goal=request,spec=initial).spec.model_dump(mode='json')
        checks=reference(source,full) if e['kind'] not in ('fault','refusal') else {'workbook':None,'csv_artifacts':[],'tolerance':1e-9}
        case=Case(id=e['id'],kind=e['kind'],capability=e['capability'],source=provenance,seed=20261019,
                  group=group,input=source.name,input_sha256=sha(source),request=request,spec=full,
                  initial_spec=initial,reply=reply,revision=revision,fault=e.get('fault'),checks=checks)
        dump(folder/'case.json',case.model_dump(mode='json'))
