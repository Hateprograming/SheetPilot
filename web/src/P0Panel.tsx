import { useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import type { Workspace, FeedbackPayload } from './types'
import { api } from './api'

export function P0Panel({ workspace, locate }: { workspace: Workspace; locate: (sheet: string, row: number, version: number) => void }) {
  const client = useQueryClient()
  const id = workspace.task.task_id
  const op = workspace.current_operation
  const auth = workspace.data_authorization
  const policy = useQuery({queryKey:['data-policy'], queryFn:api.policy})
  const decision = useMutation({ mutationFn: (approved: boolean) => api.authorize(id, auth!.grant_id, approved), onSuccess: () => client.invalidateQueries({ queryKey: ['workspace', id] }) })
  const stop = useMutation({mutationFn: () => api.cancelPlan(id), onSuccess: () => client.invalidateQueries({queryKey:['workspace',id]})})
  return <>
    {auth && <section className="single-card p0-card" aria-label="异常数据授权"><h2>发送异常值前，请确认范围</h2><p>接收方：{policy.data?.recipient || '读取中'} · 模型：{policy.data?.model || '读取中'}</p><p>{auth.purpose}</p><p>工作版本 v{auth.version} · 本次最多发送 {auth.count} 条，仅包含下面的目标值。其他列及完整文件不会发送。</p>
      <details><summary>查看将发送的目标值与位置</summary><div className="table-wrap"><table><thead><tr><th>工作表</th><th>Excel 行号</th><th>目标列和值</th></tr></thead><tbody>{auth.rows.map((r, i) => <tr key={i}><td>{r.sheet}</td><td>{r.excel_row}</td><td>{Object.entries(r.target_values).map(([k,v]) => `${k}：${String(v)}`).join('；')}</td></tr>)}</tbody></table></div></details>
      <footer><button className="button primary" disabled={decision.isPending || stop.isPending || !policy.data} onClick={() => decision.mutate(true)}>允许本次发送并继续</button><button className="button secondary" disabled={decision.isPending || stop.isPending} onClick={() => decision.mutate(false)}>不发送，手动指定规则</button><button className="text-button" disabled={decision.isPending || stop.isPending} onClick={() => stop.mutate()}>停止本次处理</button></footer><p>拒绝后可在对话中输入规则，或停止本次处理。</p>{(decision.error || stop.error) && <p role="alert">{(decision.error || stop.error)?.message}</p>}</section>}
    {op && <section className="single-card p0-card" aria-label="本次处理结果">
      <h2>{op.delivery_status === 'ready' ? `本次结果 · v${op.result_version}` : op.delivery_status === 'unavailable' ? '结果文件当前不可交付' : '本次处理进度与耗时'}</h2>
      {op.delivery_status === 'unavailable' && <p role="alert">结果文件缺失或完整性检查未通过，不能下载为本次成功结果。</p>}
      {op.summary && <><p>处理范围：{op.summary.scope.join('、')}</p>{op.summary.steps.map((step, i) => <div className="p0-step" key={i}><b>步骤 {i+1} · 影响 {step.rows_affected ?? '未知'} 行</b><p>{step.metrics.map(m => `${m.label}：${m.value}`).join(' · ')}</p></div>)}<small>{op.summary.row_count_note}</small>
        <h3>已完成以下检查</h3><ul>{op.summary.checks.map((c,i) => <li key={i}>{c.passed ? '✓' : '需检查'} {c.detail || c.name}</li>)}</ul>
        {op.summary.warnings.length > 0 && <div role="note"><h3>提醒与异常</h3><ul>{op.summary.warnings.map((w,i) => <li key={i}>{w}</li>)}</ul></div>}
        <h3>仍需人工确认</h3><ul>{op.summary.not_verified.map((w,i) => <li key={i}>{w}</li>)}</ul>
        {op.summary.manual_review_required && <p>本次结果存在需要你进一步检查的项目。</p>}
      </>}
      {op.delivery_status === 'ready' && <><footer><a className="button primary" href={`/api/workbooks/${id}/versions/${op.result_version}/download`}>下载本次结果 v{op.result_version}</a><a className="text-button" href={`/api/workbooks/${id}/original`}>下载原文件</a></footer><Exceptions key={op.operation_id} workspace={workspace} locate={locate} /></>}
      <details className="p0-timing"><summary>处理耗时与交互记录</summary><p>服务端处理 {Math.round(op.metrics.server_ms/1000)} 秒；等待用户响应 {Math.round(op.metrics.waiting_ms/1000)} 秒。</p><p>这些时间不代表你的实际劳动时间；页面停留不计作工作时间。</p><p>澄清／修订 {op.metrics.clarification_turns} 次；自动重试 {op.metrics.retries} 次；暂停 {op.metrics.pauses} 次、恢复 {op.metrics.resumes} 次。</p><ul>{Object.entries(op.metrics.stage_ms).filter(([k]) => !['waiting','pause_count','resume_count'].includes(k)).map(([k,v]) => <li key={k}>{({planning:'规划',precheck:'预检',execution:'执行',verification:'核验',commit:'终检与提交',preparation:'准备及协调'} as Record<string,string>)[k] || k}：{Math.round(v/1000)} 秒</li>)}</ul></details>
      {['completed','failed','paused','cancelled','superseded'].includes(op.status) && <Feedback key={op.operation_id} workspace={workspace} />}
    </section>}
    {workspace.retention && <p className="p0-retention">闲置 {workspace.retention.days} 天后清理文件、对话及任务记录。预计到期：{workspace.retention.expires_at ? new Date(workspace.retention.expires_at).toLocaleString('zh-CN') : '待确定'}。可通过工作簿菜单主动删除；轮询不延长保留期。</p>}
  </>
}

function Exceptions({ workspace, locate }: { workspace: Workspace; locate: (sheet: string, row: number, version: number) => void }) {
  const [issue, setIssue] = useState('')
  const [offset, setOffset] = useState(0)
  const [opened, setOpened] = useState(false)
  const op = workspace.current_operation!
  const query = useQuery({ queryKey: ['exceptions', workspace.task.task_id, op.operation_id, issue, offset], queryFn: () => api.exceptions(workspace.task.task_id, op.operation_id, issue, offset), enabled: opened })
  return <details onToggle={e => setOpened(e.currentTarget.open)}><summary>按类型查看异常并定位原始行</summary>
    {query.isLoading && <p>正在读取异常…</p>}{query.error && <p role="alert">{query.error.message}</p>}
    {query.data && <><label>异常类型 <select value={issue} onChange={e => { setIssue(e.target.value); setOffset(0) }}><option value="">全部</option>{Object.entries(query.data.counts).map(([k,v]) => <option key={k} value={k}>{k}（{v}）</option>)}</select></label>
      <p>共 {query.data.total} 条相关异常；目标值只在本地异常文件中查看。</p><ul>{query.data.rows.map((row,i) => <li key={i}>{row.sheet} · Excel 第 {row.excel_row} 行 · {row.column} · {row.issue_code}<p>{row.description}</p><button className="text-button" onClick={() => locate(row.sheet, row.excel_row, query.data!.version!)}>定位该行</button></li>)}</ul>
      <footer><button className="button secondary" disabled={offset===0} onClick={() => setOffset(Math.max(0, offset-50))}>上一页</button><button className="button secondary" disabled={offset+50>=query.data.total} onClick={() => setOffset(offset+50)}>下一页</button></footer></>}
  </details>
}

function Feedback({ workspace }: { workspace: Workspace }) {
  const client = useQueryClient()
  const op = workspace.current_operation!
  const [value, setValue] = useState<FeedbackPayload>(op.feedback || { outcome: 'completed', reasons: [], note: '' })
  const [saved, setSaved] = useState(false)
  const [outcome, setOutcome] = useState<FeedbackPayload['outcome'] | ''>(op.feedback?.outcome || '')
  const save = useMutation({ mutationFn: () => api.feedback(workspace.task.task_id, op.operation_id, value), onSuccess: () => { setSaved(true); client.invalidateQueries({ queryKey: ['workspace', workspace.task.task_id] }) } })
  const net = op.metrics.net_saved_minutes
  return <details className="p0-feedback"><summary>这次是否完成目标？反馈与净省时</summary><form onSubmit={e => { e.preventDefault(); save.mutate() }}>
    <label>目标完成情况<select aria-label="目标完成情况" value={outcome} required onChange={e => { setOutcome(e.target.value as FeedbackPayload['outcome']); setSaved(false); setValue({...value, outcome: e.target.value as FeedbackPayload['outcome']}) }}><option value="" disabled>请选择完成情况</option><option value="completed">目标完成</option><option value="partial">部分完成</option><option value="not_completed">未完成</option></select></label>
    <label>原因（可选）<select value={value.reasons[0] || ''} onChange={e => setValue({...value,reasons:e.target.value ? [e.target.value] : []})}><option value="">请选择</option><option>结果不符合业务规则</option><option>异常难以处理</option><option>需要太多复核或返工</option><option>等待过久或未拿到文件</option><option>其他</option></select></label>
    <label>补充说明（可选）<textarea maxLength={2000} value={value.note} onChange={e => setValue({...value,note:e.target.value})} /></label>
    {([['review_minutes','复核与返工分钟数'],['baseline_minutes','原方法完整耗时（分钟）'],['actual_total_minutes','本次完整耗时（分钟，含描述、等待、复核与交付）']] as const).map(([key,label]) => <label key={key}>{label}<input aria-label={label} type="number" min={0} max={100000} step="any" value={value[key] ?? ''} onChange={e => { setSaved(false); setValue({...value,[key]:e.target.value==='' ? null : Number(e.target.value)}) }} /></label>)}
    <p>{net == null ? '净省时：数据不足。填写可比较的两次完整耗时后计算。' : `净省时：${net} 分钟${net<0 ? '（本次耗时更长）' : ''}。`}</p><button className="button secondary" disabled={save.isPending || !outcome}>保存反馈</button>{saved && <p role="status">反馈已保存，可继续修改。</p>}{save.error && <p role="alert">{save.error.message}</p>}
  </form></details>
}
