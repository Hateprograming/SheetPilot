import { P0Panel } from './P0Panel'
import { CSSProperties, FormEvent, useEffect, useRef, useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Link, Navigate, Route, Routes, useNavigate, useParams } from 'react-router-dom'
import ReactMarkdown from 'react-markdown'
import {
  AlertTriangle, ArrowLeft, BookOpen, Check, ChevronDown, Download, FileSpreadsheet,
  FolderOpen, GripVertical, History, Info, LoaderCircle, LockKeyhole, MoreHorizontal, Paperclip, Pause,
  Pencil, Play, RefreshCw, RotateCcw, Send, ShieldCheck, Table2, Trash2, Upload, X, XCircle,
} from 'lucide-react'
import { api, ApiError } from './api'
import type { Dict, Message, Preview, RecentWorkbook, Workspace } from './types'

const ACTIVE = new Set(['approved', 'executing', 'analyzing_request', 'intent_parsed'])
const STATUS: Record<string, string> = {
  uploaded: '准备分析', profiled: '等待任务', needs_conversion: '等待转换要求',
  intent_parsed: '正在理解', analyzing_request: '正在理解', needs_clarification: '等待补充',
  awaiting_approval: '等待确认', approved: '准备执行', executing: '正在执行',
  completed: '已完成', failed: '执行暂停', verification_failed: '已安全回滚', paused: '已暂停',
  awaiting_data_authorization: '等待数据发送授权', cancelled: '已取消',
}

function formatTime(value?: string) {
  if (!value) return '刚刚'
  return new Intl.DateTimeFormat('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }).format(new Date(value))
}

function errorText(error: unknown) {
  return error instanceof ApiError || error instanceof Error ? error.message : '操作没有完成，请重试。'
}

function Logo() {
  return <Link className="brand" to="/"><span className="brand-mark"><Table2 size={26} /></span><span><strong>SheetPilot</strong><small>本地表格处理 Agent</small></span></Link>
}

function AmbientTexture() {
  return <svg className="ambient-texture" aria-hidden="true" focusable="false">
    <filter id="sheetpilot-grain" x="-20%" y="-20%" width="140%" height="140%">
      <feTurbulence type="fractalNoise" baseFrequency="0.72 0.36" numOctaves="3" seed="17" stitchTiles="stitch" />
      <feColorMatrix type="saturate" values="0" />
      <feComponentTransfer><feFuncA type="table" tableValues="0 0.2" /></feComponentTransfer>
    </filter>
    <rect width="100%" height="100%" filter="url(#sheetpilot-grain)" />
  </svg>
}

function DesktopFrame({ sidebar, children }: { sidebar: React.ReactNode; children: React.ReactNode }) {
  return <div className="page-root"><AmbientTexture /><div className="app-frame">{sidebar}<main className="app-surface">{children}</main></div></div>
}

function RecentWorkbooks({ currentId }: { currentId?: string }) {
  const recent = useQuery({ queryKey: ['workbooks'], queryFn: api.workbooks })
  const client = useQueryClient()
  const navigate = useNavigate()
  const [menu, setMenu] = useState<string | null>(null)
  const [menuAbove, setMenuAbove] = useState(false)
  const [dialog, setDialog] = useState<{ type: 'rename' | 'delete'; id: string; name: string } | null>(null)
  const [name, setName] = useState('')
  useEffect(() => {
    if (!menu) return
    const close = () => setMenu(null)
    document.addEventListener('pointerdown', close)
    return () => document.removeEventListener('pointerdown', close)
  }, [menu])
  const refresh = (id: string) => { client.invalidateQueries({ queryKey: ['workbooks'] }); client.invalidateQueries({ queryKey: ['workspace', id] }) }
  const rename = useMutation({ mutationFn: ({ id, value }: { id: string; value: string }) => api.renameWorkbook(id, value), onSuccess: (result, variables) => {
    client.setQueryData<Workspace>(['workspace', variables.id], current => current ? { ...current, task: { ...current.task, original_name: result.name } } : current)
    client.setQueryData<RecentWorkbook[]>(['workbooks'], current => current?.map(item => item.task_id === variables.id ? { ...item, name: result.name } : item))
    refresh(variables.id)
    setDialog(null)
  } })
  const remove = useMutation({ mutationFn: (id: string) => api.deleteWorkbook(id), onSuccess: (_, id) => {
    client.setQueryData<RecentWorkbook[]>(['workbooks'], current => current?.filter(item => item.task_id !== id))
    client.removeQueries({ queryKey: ['workspace', id] })
    client.invalidateQueries({ queryKey: ['workbooks'] })
    setDialog(null)
    if (id === currentId) navigate('/')
  } })
  const scrollable = (recent.data?.length ?? 0) > 3
  return <section className="rail-section recent-workbooks"><p className="rail-label">最近工作簿</p><div className={`recent-list${scrollable ? ' is-scrollable' : ''}`} role="region" aria-label="最近工作簿列表" tabIndex={scrollable ? 0 : undefined}>
    {recent.data?.map((item, index) => <div key={item.task_id} className={`recent-item ${item.task_id === currentId || (!currentId && index === 0) ? 'selected' : ''}`}>
      <Link className="recent-item-main" to={`/workbooks/${item.task_id}`}><FileSpreadsheet size={18} /><span><b>{item.name}</b><small>{item.version ? `v${item.version} · ` : ''}{STATUS[item.status] || item.status}</small></span></Link>
      <button className="more-button" aria-label={`管理 ${item.name}`} aria-haspopup="menu" aria-expanded={menu === item.task_id} onPointerDown={event => event.stopPropagation()} onClick={event => { const list = event.currentTarget.closest('.recent-list'); setMenuAbove(!!list && list.scrollHeight > list.clientHeight && event.currentTarget.getBoundingClientRect().bottom + 110 > list.getBoundingClientRect().bottom); setMenu(menu === item.task_id ? null : item.task_id) }}><MoreHorizontal size={19} /></button>
      {menu === item.task_id && <div className={`workbook-menu ${menuAbove ? 'opens-above' : ''}`} role="menu" onPointerDown={event => event.stopPropagation()}>
        <button role="menuitem" onClick={() => { rename.reset(); remove.reset(); setName(item.name); setDialog({ type: 'rename', id: item.task_id, name: item.name }); setMenu(null) }}><Pencil size={17} />重命名</button>
        <button role="menuitem" className="danger" onClick={() => { rename.reset(); remove.reset(); setDialog({ type: 'delete', id: item.task_id, name: item.name }); setMenu(null) }}><Trash2 size={17} />删除</button>
      </div>}
    </div>)}
    {!recent.isLoading && !recent.data?.length && <div className="empty-recent"><FolderOpen size={22} /><span>上传后的工作簿会显示在这里</span></div>}
  </div>
    {dialog && <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="workbook-dialog-title"><div className="modal workbook-dialog">
      <header>{dialog.type === 'rename' ? <Pencil size={20} /> : <Trash2 size={20} />}<div><h2 id="workbook-dialog-title">{dialog.type === 'rename' ? '重命名工作簿' : '删除这个工作簿？'}</h2><p>{dialog.type === 'rename' ? '扩展名会自动保留' : '文件、版本与处理记录都会被删除'}</p></div></header>
      {dialog.type === 'rename' ? <div className="dialog-field"><label htmlFor="workbook-name">工作簿名称</label><input id="workbook-name" autoFocus value={name} onChange={event => setName(event.target.value)} onKeyDown={event => { if (event.key === 'Enter' && name.trim()) rename.mutate({ id: dialog.id, value: name }) }} /></div> : <p>确认删除“{dialog.name}”？此操作无法撤销，正在执行的工作簿不会被删除。</p>}
      <footer><button className="button secondary" onClick={() => setDialog(null)}>取消</button>{dialog.type === 'rename' ? <button className="button primary" disabled={!name.trim() || rename.isPending} onClick={() => rename.mutate({ id: dialog.id, value: name })}>保存名称</button> : <button className="button destructive" disabled={remove.isPending} onClick={() => remove.mutate(dialog.id)}>确认删除</button>}</footer>
      {(rename.error || remove.error) && <p className="inline-error"><XCircle size={15} />{errorText(rename.error || remove.error)}</p>}
    </div></div>}
  </section>
}

function SideRail({ workspace, home = false, selectedSheet, setSelectedSheet, selectedVersion, setSelectedVersion }: { workspace?: Workspace; home?: boolean; selectedSheet?: string; setSelectedSheet?: (sheet: string) => void; selectedVersion?: number; setSelectedVersion?: (version: number) => void }) {
  return <aside className="side-rail"><Logo />
    {!home && <Link className="button primary rail-upload" to="/"><Upload size={16} />上传工作簿</Link>}
    <RecentWorkbooks currentId={workspace?.task.task_id} />
    {workspace && <>
      <section className="rail-section sheet-navigation"><p className="rail-label">Sheet 列表 <span>{workspace.profile.sheets.length}</span></p><div className="sheet-list">
        {workspace.profile.sheets.map(sheet => <button key={sheet.name} className={`sheet-item ${selectedSheet === sheet.name ? 'selected' : ''}`} onClick={() => setSelectedSheet?.(sheet.name)}><Table2 size={15} /><span>{sheet.name}</span><small>{Math.max(0, sheet.max_row - sheet.header_row).toLocaleString()}</small></button>)}
      </div></section>
      <section className="rail-section version-navigation"><p className="rail-label">版本记录</p><ol className="rail-timeline">
        {[...workspace.versions].reverse().map((version, index) => <li key={version.version} className={`${index === 0 ? 'current' : ''} ${selectedVersion === version.version ? 'selected' : ''}`}><i /><button onClick={() => { if (version.sheets?.length && selectedSheet && !version.sheets.includes(selectedSheet)) setSelectedSheet?.(version.sheets[0]); setSelectedVersion?.(version.version) }}><span><b>v{version.version}{index === 0 ? ' · 当前' : ''}</b><small>{formatTime(version.created_at)}</small></span><em>{selectedVersion === version.version ? '正在预览' : '查看'}</em></button></li>)}
      </ol></section>
    </>}
    {!workspace && <div className="rail-bottom"><span className="rail-note"><LockKeyhole size={16} />当前浏览器独立会话</span></div>}
  </aside>
}

function Home() {
  const navigate = useNavigate()
  const client = useQueryClient()
  const [file, setFile] = useState<File | null>(null)
  const [consent, setConsent] = useState(false)
  const policy = useQuery({queryKey:['data-policy'], queryFn:api.policy})
  const [dragging, setDragging] = useState(false)
  const input = useRef<HTMLInputElement>(null)
  const upload = useMutation({ mutationFn: () => api.upload(file!, consent), onSuccess: ({ task_id }) => { client.invalidateQueries({ queryKey: ['workbooks'] }); navigate(`/workbooks/${task_id}`) } })
  const choose = (next?: File) => { if (next) { setFile(next); setConsent(false) } }
  return <DesktopFrame sidebar={<SideRail home />}><div className="home-main">
    <div className="home-intro"><h1>Hello User</h1><p>How can I help with you today?</p></div>
    <div className={`upload-bar ${dragging ? 'dragging' : ''}`} onDragOver={event => { event.preventDefault(); setDragging(true) }} onDragLeave={() => setDragging(false)} onDrop={event => { event.preventDefault(); setDragging(false); choose(event.dataTransfer.files[0]) }}>
      <div><b>{file ? file.name : '把工作簿拖到这里'}</b><small>{file ? `${(file.size / 1024 / 1024).toFixed(2)} MB · 已准备分析` : '支持 .xlsx / .csv，单个文件不超过 20 MB'}</small></div>
      <input ref={input} type="file" accept=".xlsx,.csv" hidden onChange={event => choose(event.target.files?.[0])} />
      <button className="button primary" onClick={() => input.current?.click()}><FolderOpen size={16} />{file ? '重新选择' : '选择文件'}</button>
    </div>
    <div className="p0-policy">{policy.data ? <><b>模型接收方：{policy.data.recipient} · {policy.data.model}</b><p>工作簿在本应用保存，闲置 {policy.data.retention_days} 天后清理；可随时通过工作簿菜单主动删除。{policy.data.extra_consent}</p><small>{policy.data.deletion}</small></> : <p>{policy.error ? '无法读取数据说明，请稍后重试。' : '正在读取数据处理说明…'}</p>}</div>
    {file && <div className="upload-actions"><label className="consent"><input type="checkbox" checked={consent} onChange={event => setConsent(event.target.checked)} /><span>我了解：任务文字及我主动填写的信息、文件名、Sheet 名、表头和统计会发给模型；异常目标值另行授权。</span></label>
      <button disabled={!consent || !policy.data || upload.isPending} className="button primary" onClick={() => upload.mutate()}>{upload.isPending ? <LoaderCircle className="spin" size={16} /> : <Play size={16} />}分析这个工作簿</button></div>}
    {upload.error && <p className="inline-error"><XCircle size={16} />{errorText(upload.error)}</p>}
    <div className="feature-row"><Feature icon={<BookOpen />} title="读结构，不猜" text="列出 Sheet、行列数、字段类型、空值和重复情况。" /><Feature icon={<Info />} title="说人话，能追问" text="不明确的规则先确认，不把模型猜测当业务规则。" /><Feature icon={<History />} title="留版本，可回退" text="每次成功执行生成新版本，原文件始终只读。" /></div>
  </div></DesktopFrame>
}

function Feature({ icon, title, text }: { icon: React.ReactNode; title: string; text: string }) {
  return <section className="feature-card"><span>{icon}</span><h2>{title}</h2><p>{text}</p></section>
}

function useWorkspace(id: string) {
  return useQuery({ queryKey: ['workspace', id], queryFn: () => api.workspace(id), refetchInterval: query => {
    const workspace = query.state.data
    if (!workspace) return 1000
    const planning = workspace.planning_run?.status
    const execution = workspace.active_run?.status
    return ACTIVE.has(workspace.task.status) || ['queued', 'waiting_for_execution', 'running'].includes(planning) || execution === 'running' ? 1000 : false
  } })
}

function WorkspaceHeader({ workspace }: { workspace: Workspace }) {
  return <header className="workspace-header"><div className="file-heading"><span className="file-mark"><FileSpreadsheet size={20} /></span><div><h1>{workspace.task.original_name}</h1><p>{STATUS[workspace.task.status] || workspace.task.status} · v{workspace.task.current_version} · {workspace.profile.sheets.length} 个 Sheet</p></div></div><div className="workspace-actions"><span className="privacy-pill"><LockKeyhole size={14} />会话隔离 · 原文件只读</span><a className="button secondary" href={`/api/workbooks/${workspace.task.task_id}/download`}><Download size={16} />{workspace.current_operation && workspace.current_operation.delivery_status !== 'ready' ? '下载上一正式版本' : '下载当前正式版本'} v{workspace.task.current_version}</a></div></header>
}

function WorkspacePage() {
  const { taskId = '' } = useParams()
  const query = useWorkspace(taskId)
  const [selectedSheet, setSelectedSheet] = useState('')
  const [previewVersion, setPreviewVersion] = useState<number | null>(null)
  const [targetRow, setTargetRow] = useState<number | undefined>()
  useEffect(() => { if (query.data && !selectedSheet) setSelectedSheet(query.data.profile.sheets[0]?.name || '') }, [query.data, selectedSheet])
  if (query.isLoading) return <LoadingScreen />
  if (query.error || !query.data) return <NotFound message={errorText(query.error)} />
  const workspace = query.data
  const shownVersion = previewVersion ?? workspace.task.current_version
  return <DesktopFrame sidebar={<SideRail workspace={workspace} selectedSheet={selectedSheet} setSelectedSheet={setSelectedSheet} selectedVersion={shownVersion} setSelectedVersion={version => setPreviewVersion(version === workspace.task.current_version ? null : version)} />}><div className="workspace-page"><WorkspaceHeader workspace={workspace} /><WorkspaceBody workspace={workspace} selectedSheet={selectedSheet} setSelectedSheet={setSelectedSheet} previewVersion={previewVersion ?? undefined} targetRow={targetRow} onLocate={(sheet,row,version) => { setSelectedSheet(sheet); setTargetRow(row); setPreviewVersion(version) }} /></div></DesktopFrame>
}

function WorkspaceBody({ workspace, selectedSheet, setSelectedSheet, previewVersion, targetRow, onLocate }: { workspace: Workspace; selectedSheet: string; setSelectedSheet: (sheet: string) => void; previewVersion?: number; targetRow?:number; onLocate:(sheet:string,row:number,version:number)=>void }) {
  const client = useQueryClient()
  const [text, setText] = useState('')
  const [editing, setEditing] = useState(false)
  const textarea = useRef<HTMLTextAreaElement>(null)
  const columns = useRef<HTMLDivElement>(null)
  const resizing = useRef(false)
  const [taskRatio, setTaskRatio] = useState(() => {
    const stored = window.localStorage.getItem('sheetpilot-task-pane-ratio')
    if (stored == null) return .61
    const saved = Number(stored)
    return Number.isFinite(saved) && saved >= 0 && saved <= 1 ? saved : .61
  })
  const [isResizing, setIsResizing] = useState(false)
  const id = workspace.task.task_id
  const invalidate = () => client.invalidateQueries({ queryKey: ['workspace', id] })
  const send = useMutation({ mutationFn: (value: string) => api.message(id, value), onSuccess: () => { setText(''); setEditing(false); invalidate() } })
  const choice = useMutation({ mutationFn: (value: string) => api.message(id, value), onSuccess: invalidate })
  const latestPlan = [...workspace.messages].reverse().find(item => item.kind === 'plan')
  const latestClarification = [...workspace.messages].reverse().find(item => item.kind === 'clarification')
  const status = workspace.task.status
  const submit = (event: FormEvent) => { event.preventDefault(); if (text.trim()) send.mutate(text.trim()) }
  useEffect(() => {
    const field = textarea.current
    if (!field) return
    field.style.height = '0px'
    const nextHeight = Math.min(field.scrollHeight, 176)
    field.style.height = `${Math.max(46, nextHeight)}px`
    field.style.overflowY = field.scrollHeight > 176 ? 'auto' : 'hidden'
  }, [text, taskRatio])
  useEffect(() => { window.localStorage.setItem('sheetpilot-task-pane-ratio', String(taskRatio)) }, [taskRatio])
  useEffect(() => {
    const container = columns.current
    if (!container) return
    const observer = new ResizeObserver(entries => {
      const width = entries[0]?.contentRect.width || 0
      if (width) setTaskRatio(current => Math.min(current, Math.max(0, (width - 332) / width)))
    })
    observer.observe(container)
    return () => observer.disconnect()
  }, [])
  const moveDivider = (clientX: number) => {
    const rect = columns.current?.getBoundingClientRect()
    if (!rect?.width) return
    const taskWidth = Math.max(0, Math.min(clientX - rect.left, rect.width - 332))
    setTaskRatio(taskWidth / rect.width)
  }
  const adjustDivider = (direction: number) => {
    const width = columns.current?.getBoundingClientRect().width || 1
    const max = Math.max(0, (width - 332) / width)
    setTaskRatio(current => Math.max(0, Math.min(max, current + direction)))
  }
  return <div className="workspace-body"><div ref={columns} className={`workspace-columns ${isResizing ? 'resizing' : ''}`} style={{ '--task-pane-width': `${taskRatio * 100}%` } as CSSProperties}><section className="task-pane"><div className="task-scroll">
    {['profiled', 'needs_conversion'].includes(status) && !workspace.messages.some(message => message.role === 'assistant') && <p className="structure-note">工作簿结构已经准备好。直接描述你希望完成的任务。</p>}
    <Conversation messages={workspace.messages} compact={['awaiting_approval', 'completed'].includes(status)} /><PlanningActivity workspace={workspace} />
    {status === 'needs_clarification' && latestClarification && <ClarificationBlock message={latestClarification} choose={value => choice.mutate(value)} pending={choice.isPending} />}
    {status === 'awaiting_approval' && latestPlan && <PlanBlock workspace={workspace} message={latestPlan} onEdit={() => setEditing(true)} />}
    {['approved', 'executing'].includes(status) && <ExecutionBlock workspace={workspace} />}
    {['paused', 'failed', 'verification_failed'].includes(status) && <RecoveryBlock workspace={workspace} />}
    <P0Panel workspace={workspace} locate={onLocate} />
    {(send.error || choice.error) && <div className="error-banner"><XCircle size={16} />{errorText(send.error || choice.error)}</div>}
  </div>
    <form className="composer" onSubmit={submit}><Paperclip size={18} /><textarea ref={textarea} rows={1} aria-label="任务描述" value={text} onChange={event => setText(event.target.value)} placeholder={editing ? '告诉我需要怎样调整…' : status === 'completed' ? '继续描述下一项修改…' : '描述你希望怎样处理这个工作簿…'} onKeyDown={event => { if (event.key === 'Enter' && !event.shiftKey) { event.preventDefault(); submit(event as any) } }} /><button aria-label="发送" disabled={!text.trim() || send.isPending} className="send-button">{send.isPending ? <LoaderCircle className="spin" size={17} /> : <Send size={17} />}</button></form>
  </section><button type="button" className="workspace-divider" role="separator" aria-orientation="vertical" aria-label="调整对话与表格宽度" aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(taskRatio * 100)} title="拖动调整宽度，双击恢复默认" onDoubleClick={() => setTaskRatio(.61)} onPointerDown={event => { resizing.current = true; setIsResizing(true); event.currentTarget.setPointerCapture(event.pointerId); moveDivider(event.clientX) }} onPointerMove={event => { if (resizing.current) moveDivider(event.clientX) }} onPointerUp={event => { resizing.current = false; setIsResizing(false); event.currentTarget.releasePointerCapture(event.pointerId) }} onPointerCancel={event => { resizing.current = false; setIsResizing(false); if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId) }} onKeyDown={event => { if (event.key === 'ArrowLeft') { event.preventDefault(); adjustDivider(-.04) } else if (event.key === 'ArrowRight') { event.preventDefault(); adjustDivider(.04) } else if (event.key === 'Home') { event.preventDefault(); setTaskRatio(0) } else if (event.key === 'End') { event.preventDefault(); adjustDivider(1) } }}><GripVertical size={16} aria-hidden="true" /></button><WorkbookPreview workspace={workspace} sheet={selectedSheet} setSheet={setSelectedSheet} fixedVersion={previewVersion} startRow={targetRow} /></div></div>
}

function Conversation({ messages, compact }: { messages: Message[]; compact?: boolean }) {
  const visible = messages.filter(message => !['plan', 'result', 'clarification'].includes(message.kind)).slice(compact ? -3 : -6)
  if (!visible.length) return null
  return <div className="conversation">{visible.map(message => <article key={message.id} className={`message ${message.role}`}><div><ReactMarkdown>{message.content}</ReactMarkdown></div>{message.created_at && <time>{formatTime(message.created_at)}</time>}</article>)}</div>
}

function PlanningActivity({ workspace }: { workspace: Workspace }) {
  const run = workspace.planning_run
  const active = ACTIVE.has(workspace.task.status) || ['queued', 'waiting_for_execution', 'running'].includes(run?.status)
  if (!active || !run) return null
  const events = workspace.activity_events.filter(event => event.run_kind === 'planning' && (!run.job_id || event.run_id === run.job_id)).slice(-3)
  return <div className="activity-lines" aria-live="polite" aria-label="处理进度">{events.length ? events.map(event => <p key={event.id} className={event.status === 'completed' ? 'done' : 'active'}>{event.status === 'completed' ? <Check size={15} /> : <span className="pulse-dot" />}{event.label}</p>) : <p className="active"><span className="pulse-dot" />正在理解你的要求…</p>}</div>
}

function ClarificationBlock({ message, choose, pending }: { message: Message; choose: (value: string) => void; pending: boolean }) {
  const choices: string[] = message.metadata?.choices || []
  const review = message.metadata?.review
  return <section className="single-card clarification-block"><header><AlertTriangle size={19} /><div><h2>需要你确认</h2><p>确认前不会修改工作簿</p></div></header><div className="markdown"><ReactMarkdown>{message.content}</ReactMarkdown></div>
    {choices.length > 0 && <div className="choice-row">{choices.map((item, index) => <button disabled={pending} key={item} className={`choice ${index === 0 ? 'recommended' : ''}`} onClick={() => choose(item)}>{item}</button>)}</div>}
    {review?.rows?.length ? <details className="inline-details"><summary>查看 {review.total || review.rows.length} 条待确认记录</summary><ReviewRows rows={review.rows} /></details> : null}
  </section>
}

function ReviewRows({ rows }: { rows: Dict[] }) {
  const headers = Object.keys(rows[0] || {}).slice(0, 5)
  const labels: Dict = { sheet: '工作表', row: 'Excel 行号', source_columns: '涉及的源列', raw_values: '原始值', issue_code: '问题代码', description: '说明' }
  return <div className="table-wrap review-table"><table><thead><tr>{headers.map(header => <th key={header}>{labels[header] || header}</th>)}</tr></thead><tbody>{rows.slice(0, 6).map((row, index) => <tr key={index}>{headers.map(header => <td key={header}>{formatReviewValue(row[header])}</td>)}</tr>)}</tbody></table></div>
}

function formatReviewValue(value: unknown): string {
  if (value == null || value === '') return '(空)'
  if (Array.isArray(value)) return value.map(formatReviewValue).join('、')
  if (typeof value === 'object') {
    const entries = Object.entries(value as Dict)
    if (entries.length === 1) return formatReviewValue(entries[0][1])
    return entries.map(([key, item]) => `${key}：${formatReviewValue(item)}`).join('；')
  }
  return String(value)
}

function PlanBlock({ workspace, message, onEdit }: { workspace: Workspace; message: Message; onEdit: () => void }) {
  const client = useQueryClient()
  const plan = message.metadata?.plan || {}
  const steps: Dict[] = plan.steps || []
  const requirements: string[] = plan.confirmation_requirements || []
  const [ack, setAck] = useState(requirements.length === 0)
  const run = useMutation({ mutationFn: () => api.execute(workspace.task.task_id, requirements), onSuccess: () => client.invalidateQueries({ queryKey: ['workspace', workspace.task.task_id] }) })
  const cancel = useMutation({ mutationFn: () => api.cancelPlan(workspace.task.task_id), onSuccess: () => client.invalidateQueries({ queryKey: ['workspace', workspace.task.task_id] }) })
  const estimated = steps.reduce((sum, step) => sum + Number(step.expected_changes?.estimated_rows || 0), 0)
  const sheets = [...new Set<string>(steps.flatMap(step => step.expected_changes?.modified_sheets || []))]
  return <section className="single-card plan-block"><h2>准备这样处理</h2>{message.metadata?.envelope?.goal && <p className="plan-goal">{message.metadata.envelope.goal}</p>}
    <ul>{steps.map((step, index) => <li key={step.id || index}><span /><div><b>{step.description}</b>{step.expected_changes?.summary && <p>{step.expected_changes.summary}</p>}</div></li>)}</ul>
    <p className="plan-meta">{estimated ? `预计影响 ${estimated.toLocaleString()} 行` : '将按计划处理工作簿'}{sheets.length ? `，涉及「${sheets.join('」「')}」` : ''}，原文件保持不变。</p>
    {requirements.length > 0 && <label className="acknowledge"><input type="checkbox" checked={ack} onChange={event => setAck(event.target.checked)} />我确认该计划包含覆盖已有列等高风险写入，历史值仍保留在旧版本中。</label>}
    <footer><button className="button primary" disabled={!ack || run.isPending} onClick={() => run.mutate()}><Play size={16} />确认并执行</button><button className="button secondary" onClick={onEdit}>调整计划</button><button className="text-button" onClick={() => cancel.mutate()}>取消计划</button></footer>{(run.error || cancel.error) && <p className="inline-error">{errorText(run.error || cancel.error)}</p>}
  </section>
}

function ExecutionBlock({ workspace }: { workspace: Workspace }) {
  const client = useQueryClient()
  const id = workspace.task.task_id
  const planMessage = [...workspace.messages].reverse().find(item => item.kind === 'plan')
  const steps: Dict[] = planMessage?.metadata?.plan?.steps || []
  const runByStep = new Map(workspace.tool_runs.map(run => [run.plan_step_id, run]))
  const completed = steps.filter(step => runByStep.get(step.id)?.status === 'completed').length
  const progress = steps.length ? Math.round((completed / steps.length) * 100) : 0
  const current = steps.find(step => runByStep.get(step.id)?.status !== 'completed')
  const stage = [...workspace.activity_events].reverse().find(event => event.run_kind === 'execution' && event.run_id === workspace.active_run?.run_id)
  const control = useMutation({ mutationFn: (action: string) => api.control(id, action), onSuccess: () => client.invalidateQueries({ queryKey: ['workspace', id] }) })
  return <section className="process-block"><div className="process-heading"><LoaderCircle className="spin" size={18} /><div><h2>正在执行</h2><p>{stage?.label || current?.description || '正在准备安全工作副本'} · {progress}%</p></div></div><div className="progress"><span style={{ width: `${Math.max(progress, 3)}%` }} /></div>
    <details className="inline-details"><summary>查看执行细节</summary><ul className="execution-list">{steps.map((step, index) => { const state = runByStep.get(step.id)?.status || (index === completed ? 'running' : 'pending'); return <li key={step.id || index} className={state}>{state === 'completed' ? <Check size={15} /> : state === 'running' ? <LoaderCircle className="spin" size={15} /> : <span />}{step.description}<small>{state === 'completed' ? '已完成' : state === 'running' ? '执行中' : '等待'}</small></li> })}</ul></details>
    <div className="process-actions"><button className="button secondary" disabled={control.isPending} onClick={() => control.mutate('pause')}><Pause size={16} />在安全步骤暂停</button><button className="text-button danger" disabled={control.isPending} onClick={() => control.mutate('cancel')}><X size={16} />停止处理</button></div>
  </section>
}

function RecoveryBlock({ workspace }: { workspace: Workspace }) {
  const client = useQueryClient(); const id = workspace.task.task_id
  const canResume = workspace.current_operation?.can_resume
  const control = useMutation({ mutationFn: (action: string) => api.control(id, action), onSuccess: () => client.invalidateQueries({ queryKey: ['workspace', id] }) })
  return <section className="single-card recovery-block"><header><AlertTriangle size={19} /><div><h2>{canResume ? '任务停在安全检查点' : '本次处理未完成'}</h2><p>当前正式版本 v{workspace.task.current_version}</p></div></header><p>{canResume ? workspace.active_run?.error?.safe_message || '可以从已核验检查点恢复。' : '未产生可交付的新版本。请调整要求后重新发送，历史结果仍可在版本记录中查看。'}</p>
    {canResume && <footer><button disabled={control.isPending} className="button primary" onClick={() => control.mutate('resume')}><RefreshCw size={16} />从检查点重试</button><button disabled={control.isPending} className="text-button danger" onClick={() => control.mutate('cancel')}>取消运行</button></footer>}{control.error && <p className="inline-error">{errorText(control.error)}</p>}
  </section>
}

function WorkbookPreview({ workspace, sheet, setSheet, fixedVersion, startRow }: { workspace: Workspace; sheet: string; setSheet: (sheet: string) => void; fixedVersion?: number; startRow?:number }) {
  const completed = workspace.task.status === 'completed'
  const version = fixedVersion ?? workspace.task.current_version
  const versionSheets = workspace.versions.find(item => item.version === version)?.sheets
  const availableSheets = versionSheets?.length ? versionSheets : workspace.profile.sheets.map(item => item.name)
  const preview = useQuery({ queryKey: ['preview', workspace.task.task_id, sheet, version, startRow], queryFn: () => api.preview(workspace.task.task_id, sheet, version, startRow), enabled: !!sheet })
  const isCurrent = version === workspace.task.current_version
  const title = completed && isCurrent ? '修改后预览' : version === 0 ? '原始数据预览' : isCurrent ? '当前工作版本预览' : `v${version} 历史预览`
  return <section className="preview-section"><header className="preview-heading"><div><div className="preview-title-line"><h2>{title}</h2><span>v{version}{isCurrent ? ' · 当前' : ''}</span></div><p>{isCurrent ? '这里始终展示最新正式版本。' : '正在查看历史只读版本；点击左侧版本即可切换。'}</p></div><div className="preview-controls"><label>Sheet<select value={sheet} onChange={event => setSheet(event.target.value)}>{availableSheets.map(name => <option key={name} value={name}>{name}</option>)}</select></label>
    </div></header>
    <PreviewTable preview={preview.data} loading={preview.isLoading} error={preview.error} /><div className="preview-details"><details><summary>查看全部变更 <ChevronDown size={16} /></summary><Changes workspace={workspace} /></details><details><summary>查看核验详情 <ChevronDown size={16} /></summary><Verification workspace={workspace} /></details></div>
  </section>
}

function PreviewTable({ preview, loading, error }: { preview?: Preview; loading: boolean; error?: Error | null }) {
  if (loading) return <div className="preview-empty"><LoaderCircle className="spin" />正在读取预览…</div>
  if (error) return <div className="preview-empty"><XCircle />{errorText(error)}</div>
  if (!preview) return <div className="preview-empty"><Table2 />选择一个 Sheet 查看数据</div>
  const rows = preview.rows.slice(0, 8)
  return <div className="preview-table-block"><div className="table-meta"><b>{preview.sheet}</b><span>显示第 {preview.start_row || 2}–{(preview.start_row || 2)+rows.length-1} 行 · 共 {preview.total_rows.toLocaleString()} 行</span></div><div className="table-wrap"><table><thead><tr><th>#</th>{preview.headers.map((header, index) => <th key={`${header}-${index}`}>{header}</th>)}</tr></thead><tbody>{rows.map((row, index) => <tr key={index}><td>{index + (preview.start_row || 2)}</td>{preview.headers.map((header, col) => <td key={`${header}-${col}`} className={row[header] == null || row[header] === '' ? 'empty-cell' : ''}>{row[header] == null || row[header] === '' ? '(空)' : String(row[header])}</td>)}</tr>)}</tbody></table></div></div>
}

function Changes({ workspace }: { workspace: Workspace }) {
  const summary = workspace.current_operation?.summary
  const steps = summary?.steps || []
  if (!steps.length) return <p className="details-empty">本次操作尚无已提交变更，请查看当前计划。</p>
  return <ul className="detail-list">{steps.map((step,index) => <li key={index}><b>步骤 {index+1}</b><span>{summary?.scope.join('、')} · {step.rows_affected ?? '未知'} 行</span></li>)}</ul>
}

function Verification({ workspace }: { workspace: Workspace }) {
  const summary = workspace.current_operation?.summary
  const report = summary ? {status:summary.verification_status,checks:summary.checks} : null
  if (!report) return <p className="details-empty">任务执行完成后，这里会展示核验依据。</p>
  const checks: Dict[] = report.checks || []
  return <div><p className="verify-lead"><ShieldCheck size={18} />{report.status === 'passed' ? '已完成以下检查' : report.status === 'passed_with_warnings' ? '核验通过，存在提醒' : '核验失败，已回滚'}</p><ul className="detail-list">{checks.map((check, index) => <li key={`${check.name}-${index}`}><b>{check.passed ? '通过' : '请检查'} · {check.name}</b><span>{check.detail}</span></li>)}</ul></div>
}

function VersionsPage() {
  const { taskId = '' } = useParams()
  const query = useWorkspace(taskId)
  const client = useQueryClient()
  const [selectedSheet, setSelectedSheet] = useState('')
  const [selectedVersion, setSelectedVersion] = useState<number | null>(null)
  const [confirm, setConfirm] = useState(false)
  const restore = useMutation({ mutationFn: (version: number) => api.restore(taskId, version), onSuccess: () => { setConfirm(false); client.invalidateQueries({ queryKey: ['workspace', taskId] }); client.invalidateQueries({ queryKey: ['workbooks'] }) } })
  useEffect(() => { if (query.data) { if (!selectedSheet) setSelectedSheet(query.data.profile.sheets[0]?.name || ''); if (selectedVersion == null) setSelectedVersion(query.data.task.current_version) } }, [query.data, selectedSheet, selectedVersion])
  if (query.isLoading) return <LoadingScreen />
  if (query.error || !query.data) return <NotFound message={errorText(query.error)} />
  const workspace = query.data
  const current = workspace.task.current_version
  const selected = selectedVersion ?? current
  const selectedRecord = workspace.versions.find(version => version.version === selected)
  return <DesktopFrame sidebar={<SideRail workspace={workspace} selectedSheet={selectedSheet} setSelectedSheet={setSelectedSheet} selectedVersion={selected} setSelectedVersion={setSelectedVersion} />}><div className="workspace-page versions-page"><WorkspaceHeader workspace={workspace} /><div className="versions-layout"><div className="versions-content">
    <Link className="back-link" to={`/workbooks/${taskId}`}><ArrowLeft size={16} />回到工作区</Link><header className="versions-heading"><h1>版本轨迹</h1><p>每次成功处理都会留下可核验的版本节点；恢复不会删除后续历史。</p></header>
    <div className="version-list">{[...workspace.versions].reverse().map(version => <button key={version.version} className={`version-row ${selected === version.version ? 'selected' : ''}`} onClick={() => setSelectedVersion(version.version)}><i>{version.version === current ? <Check size={15} /> : `v${version.version}`}</i><span><b>v{version.version} · {version.version === 0 ? '初始上传' : version.version === current ? '当前版本' : '历史版本'}</b><small>{formatTime(version.created_at)} · {version.parent_version == null ? '原始文件' : `基于 v${version.parent_version}`}</small></span><em>{version.version === current ? '正在使用' : '查看详情'}</em></button>)}</div>
    <section className="version-detail"><div><h2>v{selected} 版本详情</h2><p>{selected === current ? '当前工作版本' : '历史只读版本'} · {selectedRecord ? formatTime(selectedRecord.created_at) : ''}</p></div><p>{selected === 0 ? '这是最初上传的工作簿，系统始终保留且不会覆盖。' : '该版本由一次经过确认的计划生成，并记录了处理步骤与核验结果。'}</p><footer><a className="button secondary" href={`/api/workbooks/${taskId}/versions/${selected}/download`}><Download size={16} />下载 v{selected}</a>{selected !== current && <button className="button primary" onClick={() => setConfirm(true)}><RotateCcw size={16} />恢复到 v{selected}</button>}</footer></section>
  </div><WorkbookPreview workspace={workspace} sheet={selectedSheet} setSheet={setSelectedSheet} fixedVersion={selected} /></div></div>
    {confirm && <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="restore-title"><div className="modal"><header><RotateCcw size={20} /><div><h2 id="restore-title">恢复到 v{selected}？</h2><p>当前版本仍会保留在历史记录中</p></div></header><p>恢复后，工作簿内容回到所选版本。这个动作不会覆盖原文件，也不会删除已有版本。</p><footer><button className="button secondary" onClick={() => setConfirm(false)}>取消</button><button className="button primary" disabled={restore.isPending} onClick={() => restore.mutate(selected)}>确认恢复</button></footer>{restore.error && <p className="inline-error">{errorText(restore.error)}</p>}</div></div>}
  </DesktopFrame>
}

function LoadingScreen() { return <div className="page-root"><AmbientTexture /><div className="center-state"><LoaderCircle className="spin" /><p>正在读取工作区…</p></div></div> }

function NotFound({ message = '这个工作簿不存在或已不可用。' }: { message?: string }) {
  return <div className="page-root"><AmbientTexture /><div className="center-state"><FileSpreadsheet /><h1>无法打开工作簿</h1><p>{message}</p><Link className="button primary" to="/">返回首页</Link></div></div>
}

export default function App() {
  useQuery({ queryKey: ['session'], queryFn: api.session, staleTime: Infinity })
  return <><div className="desktop-only"><Routes><Route path="/" element={<Home />} /><Route path="/workbooks/:taskId" element={<WorkspacePage />} /><Route path="/workbooks/:taskId/versions" element={<VersionsPage />} /><Route path="*" element={<Navigate to="/" replace />} /></Routes></div><div className="small-screen"><Table2 /><h1>请使用桌面浏览器</h1><p>SheetPilot 的表格预览和计划审阅需要至少 1180px 的屏幕宽度。</p></div></>
}
