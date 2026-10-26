export type Dict = Record<string, any>

export interface RecentWorkbook {
  task_id: string
  name: string
  version: number
  status: string
  updated_at: string
}

export interface SheetProfile {
  name: string
  max_row: number
  max_column: number
  header_row: number
  columns?: Array<{ name: string; inferred_type?: string; null_count?: number; unique_count?: number }>
}

export interface Message {
  id: number
  role: 'user' | 'assistant' | 'system'
  kind: string
  content: string
  metadata: Dict
  created_at?: string
}

export interface Workspace {
  current_operation?: CurrentOperation | null
  data_authorization?: { grant_id: string; purpose: string; count: number; version: number; rows: Array<{sheet:string; excel_row:number; target_values:Dict}> } | null
  retention?: {days:number; expires_at:string | null}
  task: Dict & { task_id: string; original_name: string; current_version: number; status: string }
  profile: Dict & { sheets: SheetProfile[]; size_bytes?: number; risks?: string[] }
  messages: Message[]
  versions: Array<Dict & { version: number; created_at: string; parent_version?: number; sheets?: string[] }>
  result?: Dict | null
  verification?: Dict | null
  active_run?: Dict | null
  tool_runs: Dict[]
  planning_run?: Dict | null
  activity_events: Dict[]
}

export interface Preview {
  start_row?: number
  sheet: string
  version: number
  headers: string[]
  rows: Dict[]
  total_rows: number
}

export interface FeedbackPayload {
  outcome: 'completed' | 'partial' | 'not_completed'
  reasons: string[]
  note: string
  review_minutes?: number | null
  baseline_minutes?: number | null
  actual_total_minutes?: number | null
}

export interface CurrentOperation {
  can_resume: boolean
  operation_id: string
  status: string
  result_version: number | null
  delivery_status: 'ready' | 'unavailable' | 'not_ready'
  summary: {verification_status:string; scope:string[]; steps:Array<{tool:string; rows_affected:number|null; metrics:Array<{key:string;label:string;value:number}>}>; row_count_note:string; checks:Array<{name:string;passed:boolean;detail:string}>; warnings:string[]; not_verified:string[]; manual_review_required:boolean} | null
  feedback: FeedbackPayload | null
  metrics: {stage_ms:Record<string,number>; server_ms:number; waiting_ms:number; clarification_turns:number; retries:number; pauses:number; resumes:number; net_saved_minutes:number|null}
}
