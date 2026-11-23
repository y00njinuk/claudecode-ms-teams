// Microsoft To Do 도구 (todo_*). azure.scopes 에 Tasks.ReadWrite / Tasks.Read 가 있을 때만 server.ts 가 노출한다.
// 인증 · Graph 호출은 Teams 와 같은 Auth · Graph 를 쓴다 (로그인 1회로 함께 쓴다).
import {
  SSL_HINT,
  TODO_SCOPE_R,
  TODO_SCOPE_RW,
  todoAccess,
  type Auth,
  type Config,
  type Graph,
  type GraphMethod,
  type GraphResult,
  type Log,
  type TodoAccess,
} from "./core.js"
import type { Tool, ToolResult } from "./mcp.js"

const LISTS_URL = "https://graph.microsoft.com/v1.0/me/todo/lists"
const GRAPH_ORIGIN = "https://graph.microsoft.com/"
const enc = encodeURIComponent
// @odata.nextLink 를 따라가는 최대 페이지 수.
const MAX_PAGES = 10
const DEFAULT_TOP = 50
const MAX_TOP = 200
const IMPORTANCE = ["low", "normal", "high"]
const STATUSES = ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"]
const LIST_STATUS = ["open", "completed", "all"]
const CHECKLIST_ACTIONS = ["add", "update", "delete"]

const LOGIN_HINT = "사용자에게 터미널에서 /microsoft-teams:connect 로 로그인하라고 안내하세요."
const RELOGIN_HINT =
  "Azure 앱의 delegated Tasks 권한(관리자 동의 포함)을 확인한 뒤, 사용자에게 터미널에서 /microsoft-teams:connect 로 다시 로그인(relogin)하라고 안내하세요."

export type TodoDeps = {
  cfg: Config
  log: Log
  auth: Pick<Auth, "hasRefreshToken" | "tokenScope" | "todoConsentRequired" | "refreshWithTodoScopes">
  graph: Pick<Graph, "request">
  // 토큰 scope 가 모자라면 설정을 다시 읽어 in-place 갱신한다 (Auth 가 같은 cfg 객체를 참조한다).
  reloadConfig?: () => Config
}

export function describeTodoAccess(access: TodoAccess): string {
  if (access === "readwrite") return `사용 (${TODO_SCOPE_RW})`
  return access === "read" ? `읽기 전용 (${TODO_SCOPE_R})` : "미사용 (scope 없음)"
}

// 저장된 토큰이 To Do 를 쓸 수 있는지. 로그인 상태와 함께 status · doctor 에 표시한다.
export function describeTodoToken(tokenScope: string | null): string {
  if (tokenScope == null) return "토큰 scope 알 수 없음"
  const access = todoAccess(tokenScope)
  if (access === "none") return "저장된 토큰에 Tasks 권한 없음 — /microsoft-teams:connect 로 다시 로그인(relogin) 필요"
  return `토큰 권한 ${access === "readwrite" ? TODO_SCOPE_RW : TODO_SCOPE_R}`
}

export function todoInstructions(access: TodoAccess): string {
  if (access === "none") return ""
  return [
    "",
    `Microsoft To Do (같은 Microsoft 로그인 사용${access === "read" ? `, ${TODO_SCOPE_R} 라 조회만 가능` : ""}).`,
    "- 사용자가 Microsoft To Do 작업(할 일)을 요청하면 todo_* 도구를 써라.",
    "- 목록을 모르면 todo_list_lists 로 확인하라. 목록 id · task id 는 추측하지 말고 도구 결과에서 받아라.",
    ...(access === "readwrite" ? ["- task 삭제(todo_delete_task)는 사용자에게 확인받은 뒤에만 하라."] : []),
  ].join("\n")
}

// ---------- 시간 ----------

export type DateTimeTimeZone = { dateTime: string; timeZone: string }

export function todoTimeZone(cfg: Config): string {
  return cfg.todo.timeZone || Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC"
}

function isIanaZone(timeZone: string): boolean {
  try {
    new Intl.DateTimeFormat("en-US", { timeZone })
    return true
  } catch {
    return false
  }
}

function isCalendarDate(y: number, m: number, d: number): boolean {
  const t = new Date(Date.UTC(y, m - 1, d))
  return t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d
}

// To Do 는 기한을 날짜로만 쓴다 (보낸 시간대 기준 날짜). 빈 문자열이면 해제(null).
export function parseDueDate(v: string, timeZone: string): DateTimeTimeZone | null {
  if (v === "") return null
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v.trim())
  if (!m || !isCalendarDate(+m[1], +m[2], +m[3])) throw new Error(`dueDate 는 YYYY-MM-DD 형식이어야 합니다 (빈 문자열은 해제): ${v}`)
  return { dateTime: `${m[1]}-${m[2]}-${m[3]}T00:00:00`, timeZone }
}

// 시각을 생략하면 09:00. 빈 문자열이면 해제(null).
export function parseReminder(v: string, timeZone: string): DateTimeTimeZone | null {
  if (v === "") return null
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?$/.exec(v.trim())
  const valid = m && isCalendarDate(+m[1], +m[2], +m[3]) && (m[4] == null || (+m[4] < 24 && +m[5] < 60))
  if (!m || !valid) throw new Error(`reminder 는 YYYY-MM-DDTHH:mm 또는 YYYY-MM-DD(09:00) 형식이어야 합니다 (빈 문자열은 해제): ${v}`)
  return { dateTime: `${m[1]}-${m[2]}-${m[3]}T${m[4] ?? "09"}:${m[5] ?? "00"}:00`, timeZone }
}

// Graph 는 응답의 날짜를 UTC 로 준다. 설정 시간대로 바꿔 보여준다. 설정이 Windows 이름이면 변환 표가 없으므로
// 시스템 시간대(IANA)로 바꾼다 — UTC 그대로 두면 기한 날짜가 하루 어긋나 보인다.
export function localDateTime(d: DateTimeTimeZone | null | undefined, timeZone: string): { value: string; timeZone: string } | undefined {
  if (!d?.dateTime) return undefined
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(d.dateTime)
  if (!m) return { value: d.dateTime, timeZone: d.timeZone }
  if (!/^utc$/i.test(d.timeZone ?? "")) return { value: `${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}`, timeZone: d.timeZone }
  if (!isIanaZone(timeZone)) timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC"
  const instant = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]))
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat("en-CA", {
      timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hourCycle: "h23",
    })
      .formatToParts(instant)
      .map((p) => [p.type, p.value]),
  )
  return { value: `${parts.year}-${parts.month}-${parts.day}T${parts.hour}:${parts.minute}`, timeZone }
}

const showDate = (d: DateTimeTimeZone | null | undefined, timeZone: string) => localDateTime(d, timeZone)?.value.slice(0, 10)
const showDateTime = (d: DateTimeTimeZone | null | undefined, timeZone: string) => {
  const local = localDateTime(d, timeZone)
  return local ? `${local.value} (${local.timeZone})` : undefined
}

// ---------- 입력 · 변환 ----------

function optString(a: Record<string, unknown>, key: string): string | undefined {
  const v = a[key]
  if (v === undefined || v === null) return undefined
  if (typeof v !== "string") throw new Error(`${key} 는 문자열이어야 합니다.`)
  return v
}

function requireString(a: Record<string, unknown>, key: string): string {
  const v = optString(a, key)?.trim()
  if (!v) throw new Error(`${key} 가 필요합니다.`)
  return v
}

function optBoolean(a: Record<string, unknown>, key: string): boolean | undefined {
  const v = a[key]
  if (v === undefined || v === null) return undefined
  if (typeof v !== "boolean") throw new Error(`${key} 는 true 또는 false 여야 합니다.`)
  return v
}

function optEnum(a: Record<string, unknown>, key: string, allowed: string[]): string | undefined {
  const v = optString(a, key)
  if (v !== undefined && !allowed.includes(v)) throw new Error(`${key} 는 ${allowed.join(" | ")} 중 하나여야 합니다: ${v}`)
  return v
}

// 도구 인자 → todoTask 본문. 넘긴 필드만 담는다 (PATCH 는 빠진 필드를 바꾸지 않는다).
export function taskPatch(a: Record<string, unknown>, timeZone: string): Record<string, unknown> {
  const p: Record<string, unknown> = {}
  const title = optString(a, "title")
  if (title !== undefined) {
    if (!title.trim()) throw new Error("title 이 비어 있습니다.")
    p.title = title
  }
  const body = optString(a, "body")
  if (body !== undefined) p.body = { content: body, contentType: "text" }
  const importance = optEnum(a, "importance", IMPORTANCE)
  if (importance !== undefined) p.importance = importance
  const status = optEnum(a, "status", STATUSES)
  if (status !== undefined) p.status = status
  const dueDate = optString(a, "dueDate")
  if (dueDate !== undefined) p.dueDateTime = parseDueDate(dueDate, timeZone)
  const reminder = optString(a, "reminder")
  if (reminder !== undefined) {
    p.reminderDateTime = parseReminder(reminder, timeZone)
    p.isReminderOn = reminder !== ""
  }
  return p
}

export type TodoList = { id: string; displayName: string; wellknownListName?: string; isShared?: boolean }

// list 는 목록 id 또는 표시 이름. 생략하면 기본 목록(defaultList). 표시 이름은 로캘마다 다르므로 기본 목록은 wellknownListName 으로 찾는다.
export function resolveList(lists: TodoList[], list: string | undefined): TodoList {
  const wanted = list?.trim()
  if (!wanted) {
    const def = lists.find((l) => l.wellknownListName === "defaultList")
    if (!def) throw new Error("기본 목록(defaultList)을 찾지 못했습니다. todo_list_lists 로 목록을 확인하고 list 를 지정하세요.")
    return def
  }
  const byId = lists.find((l) => l.id === wanted)
  if (byId) return byId
  const byName = lists.filter((l) => l.displayName === wanted)
  if (byName.length === 1) return byName[0]
  if (byName.length > 1) throw new Error(`이름이 "${wanted}" 인 목록이 ${byName.length}개입니다. todo_list_lists 결과의 id 로 지정하세요.`)
  throw new Error(`목록 "${wanted}" 을 찾지 못했습니다. todo_list_lists 로 확인하세요.`)
}

type GraphTask = Record<string, any>

function summarizeChecklistItem(c: Record<string, any>) {
  return { id: c.id, displayName: c.displayName, isChecked: c.isChecked }
}

export function summarizeTask(t: GraphTask, timeZone: string, full = false): Record<string, unknown> {
  const s: Record<string, unknown> = {
    id: t.id,
    title: t.title,
    status: t.status,
    importance: t.importance,
    due: showDate(t.dueDateTime, timeZone),
    reminder: t.isReminderOn ? showDateTime(t.reminderDateTime, timeZone) : undefined,
    completed: showDateTime(t.completedDateTime, timeZone),
    created: t.createdDateTime,
  }
  if (full) {
    s.body = t.body?.content || undefined
    s.checklist = Array.isArray(t.checklistItems) ? t.checklistItems.map(summarizeChecklistItem) : undefined
    s.lastModified = t.lastModifiedDateTime
  }
  return s
}

// ---------- 오류 ----------

function graphMessage(body: unknown): string {
  const msg = (body as { error?: { message?: unknown } } | null)?.error?.message
  return typeof msg === "string" && msg ? msg : ""
}

export function describeFailure(r: Extract<GraphResult, { ok: false }>, deps: TodoDeps, withDates: boolean): string {
  const msg = graphMessage(r.body)
  const detail = msg ? `: ${msg}` : ""
  switch (r.error) {
    case "auth":
      if (deps.auth.todoConsentRequired()) return `Microsoft To Do 권한(Tasks) 동의가 필요합니다. ${RELOGIN_HINT}`
      if (!deps.auth.hasRefreshToken()) return `Microsoft 로그인이 필요합니다 (저장된 토큰 없음 또는 만료). ${LOGIN_HINT}`
      return "Microsoft 로그인 토큰을 갱신하지 못했습니다. 네트워크를 확인하고 다시 시도하세요. 계속되면 " + LOGIN_HINT
    case "ssl":
      return `Microsoft Graph 연결 중 인증서 오류가 났습니다. ${SSL_HINT}`
    case "network":
      return "Microsoft Graph 에 연결하지 못했습니다 (네트워크 오류). 연결을 확인한 뒤 다시 시도하세요."
    case "badBody":
      return "Microsoft Graph 응답을 해석하지 못했습니다 (프록시 · 캡티브 포털 응답일 수 있음). 네트워크를 확인한 뒤 다시 시도하세요."
  }
  if (r.status === 403) {
    return `Microsoft To Do 권한이 없습니다 (HTTP 403${detail}). azure.scopes 에 ${TODO_SCOPE_RW}(쓰기) 또는 ${TODO_SCOPE_R}(조회)가 있는지 확인하고, ${RELOGIN_HINT}`
  }
  if (r.status === 404) {
    return `대상을 찾지 못했습니다 (HTTP 404${detail}). 목록 id 는 todo_list_lists, task id 는 todo_list_tasks, checklist 항목 id 는 todo_get_task 결과에서 확인하세요.`
  }
  if (r.status === 400) {
    return (
      `Microsoft Graph 가 요청을 거부했습니다 (HTTP 400${detail}). 입력 형식: dueDate YYYY-MM-DD, reminder YYYY-MM-DDTHH:mm, importance low|normal|high.` +
      (withDates ? ` 시간대 오류라면 설정의 todo.timeZone 을 Windows 시간대 이름(예: Korea Standard Time)으로 바꿔 보세요 (현재: ${todoTimeZone(deps.cfg)}).` : "")
    )
  }
  if (r.status === 429) return "요청이 많아 Microsoft Graph 가 일시적으로 제한했습니다 (HTTP 429). 잠시 후 다시 시도하세요."
  return `Microsoft Graph 오류 (HTTP ${r.status}${detail}). 잠시 후 다시 시도하세요.`
}

// ---------- 도구 ----------

type Need = "read" | "readwrite"

const satisfies = (have: TodoAccess, need: Need) => (need === "read" ? have !== "none" : have === "readwrite")
const scopeLabel = (need: Need) => (need === "read" ? `${TODO_SCOPE_R} 또는 ${TODO_SCOPE_RW}` : TODO_SCOPE_RW)

export function todoTools(access: TodoAccess, deps: TodoDeps): Tool[] {
  if (access === "none") return []
  const { auth, graph, log } = deps
  const tz = () => todoTimeZone(deps.cfg)

  // 여러 세션이 tokens.json 을 공유하므로, 예전 scope 를 가진 세션이 refresh 하면 Tasks 권한이 빠진 토큰이 저장될 수 있다.
  // 그러면 설정을 다시 읽고 전체 scope 로 한 번 갱신해 본다. 토큰 scope 를 모르면(null) 그대로 진행해 Graph 응답에 맡긴다.
  const ensureScope = async (need: Need): Promise<void> => {
    if (!auth.hasRefreshToken()) throw new Error(`Microsoft 로그인이 필요합니다. ${LOGIN_HINT}`)
    const tokenOk = () => {
      const scope = auth.tokenScope()
      return scope == null || satisfies(todoAccess(scope), need)
    }
    if (tokenOk()) return
    if (deps.reloadConfig) Object.assign(deps.cfg, deps.reloadConfig())
    if (!satisfies(todoAccess(deps.cfg.azure.scopes), need)) {
      throw new Error(`설정의 azure.scopes 에 ${scopeLabel(need)} 가 없습니다. 설정을 고친 뒤 Claude Code 세션을 다시 시작하세요.`)
    }
    const refreshed = await auth.refreshWithTodoScopes()
    if (tokenOk()) return
    if (auth.todoConsentRequired()) throw new Error(`Microsoft To Do 권한(Tasks) 동의가 필요합니다. ${RELOGIN_HINT}`)
    if (!refreshed) {
      if (!auth.hasRefreshToken()) throw new Error(`Microsoft 로그인이 만료되었습니다. ${LOGIN_HINT}`)
      throw new Error("Tasks 권한을 받기 위한 토큰 갱신에 실패했습니다 (네트워크 등). 잠시 후 다시 시도하세요.")
    }
    throw new Error(`저장된 로그인 토큰에 ${scopeLabel(need)} 권한이 없습니다 (scope 추가 전에 로그인했거나 다른 세션이 예전 scope 로 갱신). ${RELOGIN_HINT}`)
  }

  const call = async (method: GraphMethod, url: string, body?: Record<string, unknown>): Promise<any> => {
    const r = await graph.request(method, url, body)
    if (!r.ok) throw new Error(describeFailure(r, deps, Boolean(body?.dueDateTime || body?.reminderDateTime)))
    return r.body
  }

  // 토큰이 Graph 밖으로 나가지 않도록 nextLink 는 Graph 주소일 때만 따라간다.
  const getAll = async (url: string, max = Infinity): Promise<any[]> => {
    const items: any[] = []
    let next: string | undefined = url
    for (let page = 0; next && page < MAX_PAGES && items.length < max; page++) {
      const body: any = await call("GET", next)
      if (Array.isArray(body?.value)) items.push(...body.value)
      const link = body?.["@odata.nextLink"]
      next = typeof link === "string" && link.startsWith(GRAPH_ORIGIN) ? link : undefined
    }
    return items.slice(0, max)
  }

  const listOf = async (a: Record<string, unknown>): Promise<TodoList> => resolveList(await getAll(LISTS_URL), optString(a, "list"))
  const tasksUrl = (listId: string) => `${LISTS_URL}/${enc(listId)}/tasks`
  const taskUrl = (listId: string, taskId: string) => `${tasksUrl(listId)}/${enc(taskId)}`

  const listArg = { type: "string", description: "목록 id 또는 표시 이름 (todo_list_lists 결과). 생략하면 기본 목록 (wellknownListName=defaultList)." }
  const taskIdArg = { type: "string", description: "task id (todo_list_tasks · todo_create_task 결과)" }
  const zone = tz()
  const taskFields = {
    title: { type: "string", description: "제목" },
    body: { type: "string", description: "메모 (plain text)" },
    importance: { type: "string", enum: IMPORTANCE },
    dueDate: { type: "string", description: `기한 YYYY-MM-DD (${zone} 기준). 빈 문자열이면 해제.` },
    reminder: { type: "string", description: `알림 YYYY-MM-DDTHH:mm (${zone} 기준, 시각 생략 시 09:00). 빈 문자열이면 해제.` },
  }

  type Def = { need: Need; tool: Omit<Tool, "handler">; run: (a: Record<string, unknown>) => Promise<unknown> }
  const defs: Def[] = [
    {
      need: "read",
      tool: {
        name: "todo_list_lists",
        description: "Microsoft To Do 목록(list) 전체를 조회한다.",
        inputSchema: { type: "object", properties: {} },
      },
      run: async () =>
        (await getAll(LISTS_URL)).map((l: TodoList) => ({
          id: l.id,
          displayName: l.displayName,
          wellknownListName: l.wellknownListName,
          isShared: l.isShared,
        })),
    },
    {
      need: "read",
      tool: {
        name: "todo_list_tasks",
        description: "Microsoft To Do 목록의 task 를 최근 생성순으로 조회한다. 기본은 미완료 task 50개.",
        inputSchema: {
          type: "object",
          properties: {
            list: listArg,
            status: { type: "string", enum: LIST_STATUS, description: "open(기본, 미완료) | completed | all" },
            top: { type: "integer", minimum: 1, maximum: MAX_TOP, description: `최대 개수 (기본 ${DEFAULT_TOP})` },
          },
        },
      },
      run: async (a) => {
        const status = optEnum(a, "status", LIST_STATUS) ?? "open"
        const top = a.top ?? DEFAULT_TOP
        if (typeof top !== "number" || !Number.isInteger(top) || top < 1 || top > MAX_TOP) {
          throw new Error(`top 은 1–${MAX_TOP} 정수여야 합니다.`)
        }
        const list = await listOf(a)
        const filter = status === "open" ? "status ne 'completed'" : status === "completed" ? "status eq 'completed'" : ""
        const url =
          `${tasksUrl(list.id)}?%24top=${top}&%24orderby=${enc("createdDateTime desc")}` + (filter ? `&%24filter=${enc(filter)}` : "")
        const zoneNow = tz()
        return { listId: list.id, tasks: (await getAll(url, top)).map((t) => summarizeTask(t, zoneNow)) }
      },
    },
    {
      need: "read",
      tool: {
        name: "todo_get_task",
        description: "Microsoft To Do task 한 건을 메모 · checklist 포함해 조회한다.",
        inputSchema: { type: "object", properties: { list: listArg, taskId: taskIdArg }, required: ["taskId"] },
      },
      run: async (a) => {
        const taskId = requireString(a, "taskId")
        const list = await listOf(a)
        return summarizeTask(await call("GET", `${taskUrl(list.id, taskId)}?%24expand=checklistItems`), tz(), true)
      },
    },
    {
      need: "readwrite",
      tool: {
        name: "todo_create_task",
        description: "Microsoft To Do task 를 만든다. checklist 를 주면 하위 항목도 함께 만든다.",
        inputSchema: {
          type: "object",
          properties: {
            list: listArg,
            ...taskFields,
            checklist: { type: "array", items: { type: "string" }, description: "checklist 항목 이름 목록" },
          },
          required: ["title"],
        },
      },
      run: async (a) => {
        requireString(a, "title")
        const checklist = a.checklist ?? []
        if (!Array.isArray(checklist) || !checklist.every((c) => typeof c === "string" && c.trim())) {
          throw new Error("checklist 는 비어 있지 않은 문자열 배열이어야 합니다.")
        }
        const zoneNow = tz()
        const body = taskPatch({ ...a, status: undefined }, zoneNow)
        const list = await listOf(a)
        const task = await call("POST", tasksUrl(list.id), body)
        // 생성 본문에 checklistItems 를 함께 넣는 방식은 확인하지 못해, 문서화된 checklistItems 생성 API 로 하나씩 만든다.
        const items: Array<Record<string, any>> = []
        for (const displayName of checklist as string[]) {
          try {
            items.push(await call("POST", `${taskUrl(list.id, task.id)}/checklistItems`, { displayName }))
          } catch (e) {
            throw new Error(
              `task 는 만들었지만 checklist 항목 "${displayName}" 추가에 실패했습니다 (task id: ${task.id}, 추가된 항목 ${items.length}개). ` +
                `todo_checklist(action=add) 로 나머지를 추가하세요. 원인: ${(e as Error).message}`,
            )
          }
        }
        return { listId: list.id, task: summarizeTask({ ...task, checklistItems: items }, zoneNow, true) }
      },
    },
    {
      need: "readwrite",
      tool: {
        name: "todo_update_task",
        description: "Microsoft To Do task 의 제목 · 메모 · 중요도 · 상태 · 기한 · 알림을 수정한다. 완료 처리는 status=completed.",
        inputSchema: {
          type: "object",
          properties: { list: listArg, taskId: taskIdArg, ...taskFields, status: { type: "string", enum: STATUSES } },
          required: ["taskId"],
        },
      },
      run: async (a) => {
        const taskId = requireString(a, "taskId")
        const zoneNow = tz()
        const patch = taskPatch(a, zoneNow)
        if (!Object.keys(patch).length) throw new Error("수정할 필드가 없습니다 (title · body · importance · status · dueDate · reminder).")
        const list = await listOf(a)
        return summarizeTask(await call("PATCH", taskUrl(list.id, taskId), patch), zoneNow, true)
      },
    },
    {
      need: "readwrite",
      tool: {
        name: "todo_delete_task",
        description: "Microsoft To Do task 를 삭제한다 (복구 불가). 사용자에게 확인받은 뒤에만 호출한다.",
        inputSchema: { type: "object", properties: { list: listArg, taskId: taskIdArg }, required: ["taskId"] },
      },
      run: async (a) => {
        const taskId = requireString(a, "taskId")
        const list = await listOf(a)
        await call("DELETE", taskUrl(list.id, taskId))
        return { deleted: taskId }
      },
    },
    {
      need: "readwrite",
      tool: {
        name: "todo_checklist",
        description: "Microsoft To Do task 의 checklist 항목을 추가(add) · 수정(update) · 삭제(delete)한다.",
        inputSchema: {
          type: "object",
          properties: {
            list: listArg,
            taskId: taskIdArg,
            action: { type: "string", enum: CHECKLIST_ACTIONS },
            itemId: { type: "string", description: "update · delete 대상 checklist 항목 id (todo_get_task 결과)" },
            displayName: { type: "string", description: "항목 이름 (add 필수)" },
            isChecked: { type: "boolean", description: "완료 여부" },
          },
          required: ["taskId", "action"],
        },
      },
      run: async (a) => {
        const taskId = requireString(a, "taskId")
        const action = optEnum(a, "action", CHECKLIST_ACTIONS)
        if (!action) throw new Error(`action 이 필요합니다 (${CHECKLIST_ACTIONS.join(" | ")}).`)
        const displayName = optString(a, "displayName")
        const isChecked = optBoolean(a, "isChecked")
        const itemId = action === "add" ? "" : requireString(a, "itemId")
        if (action === "add" && !displayName?.trim()) throw new Error("add 에는 displayName 이 필요합니다.")
        if (displayName !== undefined && !displayName.trim()) throw new Error("displayName 이 비어 있습니다.")
        const patch = { ...(displayName !== undefined ? { displayName } : {}), ...(isChecked !== undefined ? { isChecked } : {}) }
        if (action === "update" && !Object.keys(patch).length) throw new Error("update 에는 displayName 또는 isChecked 가 필요합니다.")
        const list = await listOf(a)
        const base = `${taskUrl(list.id, taskId)}/checklistItems`
        if (action === "add") return summarizeChecklistItem(await call("POST", base, patch))
        if (action === "update") return summarizeChecklistItem(await call("PATCH", `${base}/${enc(itemId)}`, patch))
        await call("DELETE", `${base}/${enc(itemId)}`)
        return { deleted: itemId }
      },
    },
  ]

  return defs
    .filter((d) => satisfies(access, d.need))
    .map(
      (d): Tool => ({
        ...d.tool,
        handler: async (args): Promise<ToolResult> => {
          try {
            await ensureScope(d.need)
            return { text: JSON.stringify(await d.run(args), null, 2) }
          } catch (e) {
            const text = e instanceof Error ? e.message : String(e)
            log.warn("todo.failed", { tool: d.tool.name, error: text })
            return { isError: true, text }
          }
        },
      }),
    )
}
