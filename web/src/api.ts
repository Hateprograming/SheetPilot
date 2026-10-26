import type { Preview, RecentWorkbook, Workspace, FeedbackPayload } from './types'

export class ApiError extends Error {
  constructor(message: string, public code = 'RequestError', public detail: unknown = null) {
    super(message)
  }
}

async function request<T>(url: string, init?: RequestInit): Promise<T> {
  const response = await fetch(url, { credentials: 'same-origin', ...init })
  if (!response.ok) {
    let payload: any = null
    try { payload = await response.json() } catch { /* response was not JSON */ }
    const body = payload?.detail && typeof payload.detail === 'string' ? payload.detail : payload
    throw new ApiError(body?.message || body || '操作没有完成，请稍后重试。', body?.code, body?.detail)
  }
  return response.json()
}

const json = (body: unknown): RequestInit => ({
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(body),
})

const jsonWithMethod = (method: string, body: unknown): RequestInit => ({
  method,
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(body),
})

export const api = {
  policy: () => request<{recipient:string;provider:string;model:string;retention_days:number;default_sent:string[];extra_consent:string;deletion:string}>('/api/data-policy'),
  feedback: (id:string, op:string, payload:FeedbackPayload) => request(`/api/workbooks/${id}/operations/${op}/feedback`, jsonWithMethod('PUT',payload)),
  authorize: (id:string, grant:string, approved:boolean) => request(`/api/workbooks/${id}/authorizations/${grant}`, json({approved})),
  exceptions: (id:string, op:string, issue:string, offset:number) => request<{rows:Array<{sheet:string;excel_row:number;column:string;issue_code:string;description:string}>;counts:Record<string,number>;total:number;version:number|null}>(`/api/workbooks/${id}/operations/${op}/exceptions?offset=${offset}&limit=50${issue ? `&issue=${encodeURIComponent(issue)}` : ''}`),
  session: () => request<{ status: string }>('/api/session'),
  workbooks: () => request<RecentWorkbook[]>('/api/workbooks'),
  workspace: (id: string) => request<Workspace>(`/api/workbooks/${id}/workspace`),
  renameWorkbook: (id: string, name: string) => request<{ name: string }>(`/api/workbooks/${id}`, jsonWithMethod('PATCH', { name })),
  deleteWorkbook: (id: string) => request<{ status: string }>(`/api/workbooks/${id}`, { method: 'DELETE' }),
  preview: (id: string, sheet: string, version?: number, startRow?: number) => request<Preview>(
    `/api/workbooks/${id}/preview?sheet=${encodeURIComponent(sheet)}&limit=50${version == null ? '' : `&version=${version}`}${startRow == null ? '' : `&start_row=${startRow}`}`,
  ),
  upload: async (file: File, consent: boolean) => {
    const body = new FormData()
    body.append('file', file)
    body.append('consent', String(consent))
    return request<{ task_id: string; workspace: Workspace }>('/api/workbooks', { method: 'POST', body })
  },
  message: (id: string, text: string) => request(`/api/workbooks/${id}/messages`, json({ text })),
  customClarification: (id: string, label: string) => request(`/api/workbooks/${id}/clarifications/custom`, json({ label })),
  execute: (id: string, acknowledgements: string[]) => request(`/api/workbooks/${id}/plan/execute`, json({ acknowledgements })),
  cancelPlan: (id: string) => request(`/api/workbooks/${id}/plan/cancel`, { method: 'POST' }),
  control: (id: string, action: string) => request(`/api/workbooks/${id}/run/${action}`, { method: 'POST' }),
  restore: (id: string, version: number) => request(`/api/workbooks/${id}/versions/${version}/restore`, { method: 'POST' }),
}
