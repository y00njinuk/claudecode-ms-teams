import { describe, expect, test } from "bun:test"
import { MicrosoftTeamsChannel } from "../src/channel"
import { defaultConfig, mergeConfig, NOOP_LOG, type Config, type GraphMethod, type GraphResult, type TodoAccess } from "../src/core"
import {
  localDateTime,
  parseDueDate,
  parseReminder,
  resolveList,
  taskPatch,
  todoInstructions,
  todoTimeZone,
  todoTools,
  type TodoList,
} from "../src/todo"

const LISTS: TodoList[] = [
  { id: "L-default", displayName: "작업", wellknownListName: "defaultList", isShared: false },
  { id: "L-flag", displayName: "플래그가 지정된 메일", wellknownListName: "flaggedEmails", isShared: false },
  { id: "L-a", displayName: "업무", wellknownListName: "none", isShared: false },
  { id: "L-dup1", displayName: "중복", wellknownListName: "none", isShared: false },
  { id: "L-dup2", displayName: "중복", wellknownListName: "none", isShared: true },
]

const ok = (body: unknown, status = 200): GraphResult => ({ ok: true, status, body })
const fail = (status: number, message = "", error: "http" | "auth" | "network" | "ssl" | "badBody" = "http"): GraphResult => ({
  ok: false,
  status,
  body: message ? { error: { code: "x", message } } : null,
  error,
})

type Call = { method: GraphMethod; url: string; body?: any }

function harness(
  opts: {
    access?: TodoAccess
    scopes?: string
    tokenScope?: string | null
    refreshedScope?: string | null
    consent?: boolean
    hasToken?: boolean
    reloaded?: Config
    route?: (c: Call) => GraphResult | undefined
  } = {},
) {
  const cfg = mergeConfig(defaultConfig(), {
    azure: { tenantId: "t", clientId: "c", scopes: opts.scopes ?? "Chat.Read offline_access Tasks.ReadWrite" },
    todo: { timeZone: "Asia/Seoul" },
  })
  const calls: Call[] = []
  let scope = opts.tokenScope === undefined ? "Chat.Read Tasks.ReadWrite" : opts.tokenScope
  let refreshes = 0
  let reloads = 0
  const tools = todoTools(opts.access ?? "readwrite", {
    cfg,
    log: NOOP_LOG,
    auth: {
      hasRefreshToken: () => opts.hasToken ?? true,
      tokenScope: () => scope,
      todoConsentRequired: () => opts.consent ?? false,
      refreshWithTodoScopes: async () => {
        refreshes++
        if (opts.refreshedScope !== undefined) scope = opts.refreshedScope
        return "token"
      },
    },
    graph: {
      request: async (method, url, body) => {
        const c = { method, url, body }
        calls.push(c)
        const routed = opts.route?.(c)
        if (routed) return routed
        if (method === "GET" && url === "https://graph.microsoft.com/v1.0/me/todo/lists") return ok({ value: LISTS })
        return fail(500)
      },
    },
    reloadConfig: () => {
      reloads++
      return opts.reloaded ?? cfg
    },
  })
  const run = async (name: string, args: Record<string, unknown> = {}) => {
    const tool = tools.find((t) => t.name === name)
    if (!tool) throw new Error(`no tool ${name}`)
    const r = await tool.handler(args)
    return { ...r, json: r.isError ? null : JSON.parse(r.text) }
  }
  return {
    cfg,
    tools,
    calls,
    run,
    get refreshes() {
      return refreshes
    },
    get reloads() {
      return reloads
    },
  }
}

describe("기능 게이트", () => {
  const names = (access: TodoAccess) => harness({ access }).tools.map((t) => t.name)

  test("Tasks.ReadWrite 는 7개, Tasks.Read 는 조회 3개, 없으면 0개", () => {
    expect(names("readwrite")).toEqual([
      "todo_list_lists",
      "todo_list_tasks",
      "todo_get_task",
      "todo_create_task",
      "todo_update_task",
      "todo_delete_task",
      "todo_checklist",
    ])
    expect(names("read")).toEqual(["todo_list_lists", "todo_list_tasks", "todo_get_task"])
    expect(names("none")).toEqual([])
  })

  test("Teams 도구 4개는 그대로이고 이름이 겹치지 않는다", () => {
    const ch = new MicrosoftTeamsChannel({ cfg: defaultConfig(), log: NOOP_LOG, auth: {} as never, graph: {} as never, configPath: "", push: () => {} })
    const teams = ch.tools().map((t) => t.name)
    expect(teams).toEqual(["connect", "select_chat", "status", "disconnect"])
    expect(names("readwrite").some((n) => teams.includes(n))).toBe(false)
  })

  test("instructions 는 To Do 가 켜졌을 때만 추가되고, 삭제 확인 지침은 쓰기 권한일 때만 있다", () => {
    expect(todoInstructions("none")).toBe("")
    expect(todoInstructions("readwrite")).toContain("todo_list_lists")
    expect(todoInstructions("readwrite")).toContain("사용자에게 확인받은 뒤에만")
    expect(todoInstructions("read")).toContain("조회만 가능")
    expect(todoInstructions("read")).not.toContain("todo_delete_task")
  })

  test("todo_delete_task 설명에 복구 불가를 적는다", () => {
    expect(harness().tools.find((t) => t.name === "todo_delete_task")?.description).toContain("복구 불가")
  })
})

describe("목록 해석", () => {
  test("생략하면 defaultList, id · 표시 이름으로 찾는다", () => {
    expect(resolveList(LISTS, undefined).id).toBe("L-default")
    expect(resolveList(LISTS, "  ").id).toBe("L-default")
    expect(resolveList(LISTS, "L-a").id).toBe("L-a")
    expect(resolveList(LISTS, "업무").id).toBe("L-a")
  })

  test("같은 이름이 여럿이거나 없으면 오류", () => {
    expect(() => resolveList(LISTS, "중복")).toThrow("id 로 지정")
    expect(() => resolveList(LISTS, "없는목록")).toThrow("todo_list_lists")
    expect(() => resolveList(LISTS.slice(2), undefined)).toThrow("defaultList")
  })
})

describe("날짜 변환", () => {
  test("dueDate 는 YYYY-MM-DD 만 받고 그 날짜 00:00 으로 보낸다", () => {
    expect(parseDueDate("2026-11-25", "Asia/Seoul")).toEqual({ dateTime: "2026-11-25T00:00:00", timeZone: "Asia/Seoul" })
    expect(parseDueDate("", "Asia/Seoul")).toBeNull()
    for (const bad of ["2026-11-25T10:00", "2026/11/25", "26-11-25", "2026-02-30", "내일"]) {
      expect(() => parseDueDate(bad, "Asia/Seoul")).toThrow("YYYY-MM-DD")
    }
  })

  test("reminder 는 YYYY-MM-DDTHH:mm, 시각을 생략하면 09:00", () => {
    expect(parseReminder("2026-11-25T14:30", "Asia/Seoul")).toEqual({ dateTime: "2026-11-25T14:30:00", timeZone: "Asia/Seoul" })
    expect(parseReminder("2026-11-25 08:05", "Asia/Seoul")?.dateTime).toBe("2026-11-25T08:05:00")
    expect(parseReminder("2026-11-25", "Asia/Seoul")?.dateTime).toBe("2026-11-25T09:00:00")
    expect(parseReminder("", "Asia/Seoul")).toBeNull()
    for (const bad of ["2026-11-25T24:00", "2026-11-25T10:60", "2026-11-25T9:00", "2026-13-01"]) {
      expect(() => parseReminder(bad, "Asia/Seoul")).toThrow("YYYY-MM-DDTHH:mm")
    }
  })

  test("응답의 UTC 시각을 설정 시간대로 바꿔 보여준다 (Windows 이름이면 그대로)", () => {
    expect(localDateTime({ dateTime: "2026-11-24T15:00:00.0000000", timeZone: "UTC" }, "Asia/Seoul")).toEqual({
      value: "2026-11-25T00:00",
      timeZone: "Asia/Seoul",
    })
    expect(localDateTime({ dateTime: "2026-11-25T05:30:00.0000000", timeZone: "UTC" }, "Korea Standard Time")).toEqual({
      value: "2026-11-25T05:30",
      timeZone: "UTC",
    })
    expect(localDateTime({ dateTime: "2026-11-25T09:00:00", timeZone: "Korea Standard Time" }, "Asia/Seoul")).toEqual({
      value: "2026-11-25T09:00",
      timeZone: "Korea Standard Time",
    })
    expect(localDateTime(null, "Asia/Seoul")).toBeUndefined()
  })

  test("todo.timeZone 이 비어 있으면 시스템 시간대를 쓴다", () => {
    expect(todoTimeZone(mergeConfig(defaultConfig(), { todo: { timeZone: "Europe/London" } }))).toBe("Europe/London")
    expect(todoTimeZone(defaultConfig())).toBe(Intl.DateTimeFormat().resolvedOptions().timeZone)
  })
})

describe("patch 생성", () => {
  test("넘긴 필드만 todoTask 본문으로 바꾼다", () => {
    expect(
      taskPatch(
        { title: "보고서", body: "초안", importance: "high", status: "completed", dueDate: "2026-11-25", reminder: "2026-11-24T18:00" },
        "Asia/Seoul",
      ),
    ).toEqual({
      title: "보고서",
      body: { content: "초안", contentType: "text" },
      importance: "high",
      status: "completed",
      dueDateTime: { dateTime: "2026-11-25T00:00:00", timeZone: "Asia/Seoul" },
      reminderDateTime: { dateTime: "2026-11-24T18:00:00", timeZone: "Asia/Seoul" },
      isReminderOn: true,
    })
    expect(taskPatch({ list: "업무", taskId: "T1" }, "Asia/Seoul")).toEqual({})
  })

  test("빈 문자열은 기한 · 알림 해제, 알림 해제는 isReminderOn=false", () => {
    expect(taskPatch({ dueDate: "", reminder: "" }, "Asia/Seoul")).toEqual({ dueDateTime: null, reminderDateTime: null, isReminderOn: false })
  })

  test("잘못된 값은 거부한다", () => {
    expect(() => taskPatch({ importance: "urgent" }, "UTC")).toThrow("low | normal | high")
    expect(() => taskPatch({ status: "done" }, "UTC")).toThrow("notStarted")
    expect(() => taskPatch({ title: "  " }, "UTC")).toThrow("title")
    expect(() => taskPatch({ body: 3 }, "UTC")).toThrow("문자열")
  })
})

describe("도구 호출", () => {
  const TASK = {
    id: "T1",
    title: "보고서",
    status: "notStarted",
    importance: "normal",
    dueDateTime: { dateTime: "2026-11-24T15:00:00.0000000", timeZone: "UTC" },
    isReminderOn: true,
    reminderDateTime: { dateTime: "2026-11-24T09:00:00.0000000", timeZone: "UTC" },
    createdDateTime: "2026-11-23T01:00:00Z",
    lastModifiedDateTime: "2026-11-23T01:00:00Z",
    body: { content: "", contentType: "text" },
  }

  test("todo_list_lists 는 목록 요약을 돌려준다", async () => {
    const h = harness()
    const r = await h.run("todo_list_lists")
    expect(r.json[0]).toEqual({ id: "L-default", displayName: "작업", wellknownListName: "defaultList", isShared: false })
    expect(r.json).toHaveLength(LISTS.length)
  })

  test("todo_list_tasks 는 미완료 필터 · 최근순으로 조회하고 nextLink 를 top 까지 따라간다", async () => {
    const page2 = "https://graph.microsoft.com/v1.0/me/todo/lists/L-a/tasks?%24skip=2"
    const h = harness({
      route: (c) => {
        if (c.url.startsWith("https://graph.microsoft.com/v1.0/me/todo/lists/L-a/tasks?%24top")) {
          return ok({ value: [TASK, { ...TASK, id: "T2" }], "@odata.nextLink": page2 })
        }
        if (c.url === page2) return ok({ value: [{ ...TASK, id: "T3" }, { ...TASK, id: "T4" }], "@odata.nextLink": page2 + "x" })
      },
    })
    const r = await h.run("todo_list_tasks", { list: "업무", top: 3 })
    expect(r.json.listId).toBe("L-a")
    expect(r.json.tasks.map((t: { id: string }) => t.id)).toEqual(["T1", "T2", "T3"])
    expect(r.json.tasks[0]).toEqual({
      id: "T1",
      title: "보고서",
      status: "notStarted",
      importance: "normal",
      due: "2026-11-25",
      reminder: "2026-11-24T18:00 (Asia/Seoul)",
      created: "2026-11-23T01:00:00Z",
    })
    const first = h.calls.find((c) => c.url.includes("/tasks?"))!.url
    expect(first).toContain("%24top=3")
    expect(first).toContain(`%24orderby=${encodeURIComponent("createdDateTime desc")}`)
    expect(first).toContain(`%24filter=${encodeURIComponent("status ne 'completed'")}`)
    expect(h.calls.filter((c) => c.url.startsWith(page2))).toHaveLength(1)
  })

  test("todo_list_tasks 는 status=all 이면 필터 없이, Graph 밖 nextLink 는 따라가지 않는다", async () => {
    const h = harness({
      route: (c) => (c.url.includes("/tasks?") ? ok({ value: [TASK], "@odata.nextLink": "https://evil.example/next" }) : undefined),
    })
    const r = await h.run("todo_list_tasks", { status: "all" })
    expect(r.json.listId).toBe("L-default")
    expect(h.calls.some((c) => c.url.includes("%24filter"))).toBe(false)
    expect(h.calls.some((c) => c.url.includes("evil.example"))).toBe(false)
  })

  test("todo_list_tasks 는 top 범위 · status 값을 검사한다", async () => {
    const h = harness()
    expect((await h.run("todo_list_tasks", { top: 0 })).text).toContain("1–200")
    expect((await h.run("todo_list_tasks", { top: 201 })).isError).toBe(true)
    expect((await h.run("todo_list_tasks", { status: "done" })).text).toContain("open | completed | all")
  })

  test("todo_get_task 는 checklist 를 펼쳐 조회하고 id 를 URL 인코딩한다", async () => {
    const h = harness({
      route: (c) =>
        c.url.includes("/tasks/") ? ok({ ...TASK, body: { content: "메모" }, checklistItems: [{ id: "C1", displayName: "a", isChecked: false, createdDateTime: "x" }] }) : undefined,
    })
    const r = await h.run("todo_get_task", { taskId: "AQ/k=+x" })
    expect(r.json.body).toBe("메모")
    expect(r.json.checklist).toEqual([{ id: "C1", displayName: "a", isChecked: false }])
    expect(r.json.lastModified).toBe("2026-11-23T01:00:00Z")
    expect(h.calls.at(-1)?.url).toBe(
      `https://graph.microsoft.com/v1.0/me/todo/lists/L-default/tasks/${encodeURIComponent("AQ/k=+x")}?%24expand=checklistItems`,
    )
  })

  test("todo_create_task 는 task 를 만든 뒤 checklist 항목을 하나씩 추가한다", async () => {
    const h = harness({
      route: (c) => {
        if (c.method === "POST" && c.url.endsWith("/L-default/tasks")) return ok({ ...TASK, title: c.body.title }, 201)
        if (c.method === "POST" && c.url.endsWith("/tasks/T1/checklistItems")) return ok({ id: `C-${c.body.displayName}`, displayName: c.body.displayName, isChecked: false }, 201)
      },
    })
    const r = await h.run("todo_create_task", { title: "보고서", dueDate: "2026-11-25", checklist: ["초안", "검토"], status: "completed" })
    expect(r.isError).toBeUndefined()
    expect(r.json.listId).toBe("L-default")
    expect(r.json.task.checklist).toEqual([
      { id: "C-초안", displayName: "초안", isChecked: false },
      { id: "C-검토", displayName: "검토", isChecked: false },
    ])
    const create = h.calls.find((c) => c.method === "POST" && c.url.endsWith("/tasks"))!
    expect(create.body).toEqual({ title: "보고서", dueDateTime: { dateTime: "2026-11-25T00:00:00", timeZone: "Asia/Seoul" } })
    expect(h.calls.filter((c) => c.url.endsWith("/checklistItems")).map((c) => c.body)).toEqual([{ displayName: "초안" }, { displayName: "검토" }])
  })

  test("todo_create_task 는 checklist 추가가 실패하면 만든 task id 와 함께 알린다", async () => {
    const h = harness({
      route: (c) => {
        if (c.method === "POST" && c.url.endsWith("/tasks")) return ok(TASK, 201)
        if (c.url.endsWith("/checklistItems")) return c.body.displayName === "검토" ? fail(400, "bad item") : ok({ id: "C1", displayName: "초안", isChecked: false })
      },
    })
    const r = await h.run("todo_create_task", { title: "보고서", checklist: ["초안", "검토"] })
    expect(r.isError).toBe(true)
    expect(r.text).toContain("task id: T1")
    expect(r.text).toContain("추가된 항목 1개")
    expect(r.text).toContain("bad item")
  })

  test("todo_create_task 입력 검사: title 필수, checklist 는 문자열 배열", async () => {
    const h = harness()
    expect((await h.run("todo_create_task", {})).text).toContain("title")
    expect((await h.run("todo_create_task", { title: "a", checklist: "b" })).text).toContain("checklist")
    expect((await h.run("todo_create_task", { title: "a", checklist: ["", "b"] })).isError).toBe(true)
    expect(h.calls).toHaveLength(0)
  })

  test("todo_update_task 는 PATCH 로 넘긴 필드만 보내고, 수정할 필드가 없으면 거부한다", async () => {
    const h = harness({ route: (c) => (c.method === "PATCH" ? ok({ ...TASK, importance: "high", isReminderOn: false }) : undefined) })
    const r = await h.run("todo_update_task", { taskId: "T1", importance: "high", reminder: "" })
    expect(r.json.importance).toBe("high")
    expect(r.json.reminder).toBeUndefined()
    expect(h.calls.at(-1)).toEqual({
      method: "PATCH",
      url: "https://graph.microsoft.com/v1.0/me/todo/lists/L-default/tasks/T1",
      body: { importance: "high", reminderDateTime: null, isReminderOn: false },
    })
    expect((await h.run("todo_update_task", { taskId: "T1" })).text).toContain("수정할 필드가 없습니다")
    expect((await h.run("todo_update_task", { title: "x" })).text).toContain("taskId")
  })

  test("todo_delete_task 는 DELETE 후 삭제한 id 를 돌려준다", async () => {
    const h = harness({ route: (c) => (c.method === "DELETE" ? ok({}, 204) : undefined) })
    const r = await h.run("todo_delete_task", { list: "L-a", taskId: "T1" })
    expect(r.json).toEqual({ deleted: "T1" })
    expect(h.calls.at(-1)?.url).toBe("https://graph.microsoft.com/v1.0/me/todo/lists/L-a/tasks/T1")
  })

  test("todo_checklist add · update · delete", async () => {
    const h = harness({
      route: (c) => {
        if (!c.url.includes("/checklistItems")) return undefined
        if (c.method === "DELETE") return ok({}, 204)
        return ok({ id: "C1", displayName: c.body.displayName ?? "초안", isChecked: c.body.isChecked ?? false, createdDateTime: "x" })
      },
    })
    expect((await h.run("todo_checklist", { taskId: "T1", action: "add", displayName: "초안" })).json).toEqual({ id: "C1", displayName: "초안", isChecked: false })
    expect((await h.run("todo_checklist", { taskId: "T1", action: "update", itemId: "C1", isChecked: true })).json.isChecked).toBe(true)
    expect(h.calls.at(-1)).toMatchObject({ method: "PATCH", body: { isChecked: true } })
    expect(h.calls.at(-1)?.url).toEndWith("/tasks/T1/checklistItems/C1")
    expect((await h.run("todo_checklist", { taskId: "T1", action: "delete", itemId: "C1" })).json).toEqual({ deleted: "C1" })
  })

  test("todo_checklist 입력 검사", async () => {
    const h = harness()
    expect((await h.run("todo_checklist", { taskId: "T1", action: "add" })).text).toContain("displayName")
    expect((await h.run("todo_checklist", { taskId: "T1", action: "update", itemId: "C1" })).text).toContain("displayName 또는 isChecked")
    expect((await h.run("todo_checklist", { taskId: "T1", action: "delete" })).text).toContain("itemId")
    expect((await h.run("todo_checklist", { taskId: "T1", action: "toggle" })).text).toContain("add | update | delete")
    expect((await h.run("todo_checklist", { taskId: "T1", action: "update", itemId: "C1", isChecked: "yes" })).text).toContain("true 또는 false")
    expect(h.calls).toHaveLength(0)
  })
})

describe("오류 매핑", () => {
  const errorFor = async (result: GraphResult, opts: Parameters<typeof harness>[0] = {}, tool = "todo_list_lists", args = {}) => {
    const h = harness({ ...opts, route: () => result })
    const r = await h.run(tool, args)
    expect(r.isError).toBe(true)
    return r.text
  }

  test("403 은 Tasks 권한 · 재로그인 안내", async () => {
    const text = await errorFor(fail(403, "Access is denied"))
    expect(text).toContain("HTTP 403: Access is denied")
    expect(text).toContain("Tasks.ReadWrite")
    expect(text).toContain("relogin")
  })

  test("404 는 id 확인 안내", async () => {
    expect(await errorFor(fail(404))).toContain("todo_list_tasks")
  })

  test("400 은 Graph 메시지와 입력 형식, 날짜가 있는 요청이면 시간대 안내", async () => {
    const plain = await errorFor(fail(400, "Invalid request"))
    expect(plain).toContain("HTTP 400: Invalid request")
    expect(plain).not.toContain("todo.timeZone")
    const h = harness({ route: (c) => (c.method === "PATCH" ? fail(400, "TimeZone not supported") : undefined) })
    const dated = await h.run("todo_update_task", { taskId: "T1", dueDate: "2026-11-25" })
    expect(dated.text).toContain("todo.timeZone")
    expect(dated.text).toContain("Asia/Seoul")
  })

  test("429 · 5xx · 네트워크 · 인증서 · 본문 오류", async () => {
    expect(await errorFor(fail(429))).toContain("HTTP 429")
    expect(await errorFor(fail(503, "busy"))).toContain("HTTP 503: busy")
    expect(await errorFor(fail(0, "", "network"))).toContain("네트워크 오류")
    expect(await errorFor(fail(0, "", "ssl"))).toContain("NODE_EXTRA_CA_CERTS")
    expect(await errorFor(fail(200, "", "badBody"))).toContain("프록시")
  })

  test("토큰을 얻지 못하면 로그인 안내, To Do 동의가 없으면 동의 안내", async () => {
    expect(await errorFor(fail(0, "", "auth"))).toContain("토큰을 갱신하지 못했습니다")
    expect(await errorFor(fail(0, "", "auth"), { consent: true })).toContain("동의가 필요합니다")
  })

  test("저장된 토큰이 없으면 Graph 를 부르지 않고 로그인을 안내한다", async () => {
    const h = harness({ hasToken: false })
    const r = await h.run("todo_list_lists")
    expect(r.text).toContain("/microsoft-teams:connect")
    expect(h.calls).toHaveLength(0)
  })
})

describe("토큰 scope 불일치", () => {
  test("토큰에 Tasks 권한이 없으면 설정을 다시 읽고 한 번 갱신한 뒤 진행한다", async () => {
    const h = harness({ tokenScope: "Chat.Read User.Read", refreshedScope: "Chat.Read Tasks.ReadWrite" })
    const r = await h.run("todo_list_lists")
    expect(r.isError).toBeUndefined()
    expect(h.refreshes).toBe(1)
    expect(h.reloads).toBe(1)
  })

  test("갱신해도 권한이 없으면 재로그인을, 동의가 없으면 동의를 안내한다 (Graph 호출 없음)", async () => {
    const missing = harness({ tokenScope: "Chat.Read", refreshedScope: "Chat.Read" })
    const r = await missing.run("todo_list_lists")
    expect(r.text).toContain("권한이 없습니다")
    expect(r.text).toContain("relogin")
    expect(missing.refreshes).toBe(1)
    expect(missing.calls).toHaveLength(0)
    const consent = harness({ tokenScope: "Chat.Read", refreshedScope: "Chat.Read", consent: true })
    expect((await consent.run("todo_list_lists")).text).toContain("동의가 필요합니다")
  })

  test("조회 도구는 Tasks.Read 토큰으로 충분하고, 쓰기 도구는 Tasks.ReadWrite 가 필요하다", async () => {
    const h = harness({
      tokenScope: "Tasks.Read",
      refreshedScope: "Tasks.Read",
      route: (c) => (c.url.includes("/tasks/") ? ok({ id: "T1" }) : undefined),
    })
    expect((await h.run("todo_get_task", { taskId: "T1" })).isError).toBeUndefined()
    expect(h.refreshes).toBe(0)
    const w = await h.run("todo_delete_task", { taskId: "T1" })
    expect(w.text).toContain("Tasks.ReadWrite")
    expect(h.refreshes).toBe(1)
  })

  test("설정에서 Tasks scope 가 빠졌으면 갱신하지 않고 세션 재시작을 안내한다", async () => {
    const reloaded = mergeConfig(defaultConfig(), { azure: { tenantId: "t", clientId: "c", scopes: "Chat.Read offline_access" } })
    const h = harness({ tokenScope: "Chat.Read", reloaded })
    const r = await h.run("todo_list_lists")
    expect(r.text).toContain("세션을 다시 시작")
    expect(h.refreshes).toBe(0)
    // Auth 와 공유하는 cfg 객체를 in-place 로 갱신한다.
    expect(h.cfg.azure.scopes).toBe("Chat.Read offline_access")
  })

  test("토큰 scope 를 모르면(null) 그대로 진행한다", async () => {
    const h = harness({ tokenScope: null })
    expect((await h.run("todo_list_lists")).isError).toBeUndefined()
    expect(h.refreshes).toBe(0)
  })
})
