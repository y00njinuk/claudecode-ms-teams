import { describe, expect, test } from "bun:test"
import { MicrosoftTeamsChannel, type AuthPort, type GraphPort, type HookDecision } from "../src/channel"
import { defaultConfig, isInsideWorkspace, mergeConfig, NOOP_LOG, type ChatMsg, type Config, type DeviceCode, type TodoAccess } from "../src/core"
import { sendToChat } from "../src/send"

const CHAT_ID = "19:abc123@thread.v2"
const ME = { id: "me-1", displayName: "나" }

function configured(extra: object = {}): Config {
  return mergeConfig(defaultConfig(), {
    azure: { tenantId: "t", clientId: "c", scopes: "s" },
    // 테스트에서는 pollOnce 를 직접 호출하므로 백그라운드 폴링이 끼어들지 않게 간격을 크게 둔다.
    polling: { intervalMs: 60 * 60 * 1000 },
    ...extra,
  })
}

function setup(opts: { cfg?: Config; hasToken?: boolean; seed?: number | null; todoAccess?: TodoAccess; tokenScope?: string | null; consent?: boolean } = {}) {
  let hasToken = opts.hasToken ?? true
  let device: DeviceCode | null = null
  let deviceFlowStarted = 0
  // Teams 처럼 발신 · 수신 메시지가 한 시간순 id 를 공유한다 (코드 없는 답장은 그보다 먼저 알린 요청에만 적용된다).
  let idSeq = 1000
  let clock = 1_000_000
  let me = ME
  let failNextSend = false
  const sent: Array<{ chatId: string; html: string; id: string }> = []
  const inbox: ChatMsg[] = []
  const pushes: Array<{ content: string; meta: Record<string, string> }> = []
  const activeChanges: boolean[] = []

  const auth: AuthPort = {
    pendingDeviceCode: () => device,
    awaitDeviceCode: async () => {},
    hasRefreshToken: () => hasToken,
    startDeviceFlow: () => {
      deviceFlowStarted++
      device = { verificationUri: "https://microsoft.com/devicelogin", userCode: "ABCD-1234", expiresAt: Date.now() + 60_000 }
    },
    getAccessToken: async () => (hasToken ? "token" : null),
    bootstrap: () => {},
    tokenScope: () => (opts.tokenScope === undefined ? "Chat.Read" : opts.tokenScope),
    todoConsentRequired: () => opts.consent ?? false,
  }
  const graph: GraphPort = {
    me: async () => me,
    sendMessage: async (chatId, html) => {
      if (failNextSend) {
        failNextSend = false
        throw new Error("network reset")
      }
      const id = String(++idSeq)
      sent.push({ chatId, html, id })
      return id
    },
    latestMessageId: async () => (opts.seed === undefined ? 500 : opts.seed),
    messagesSince: async (_chatId, cursor) => inbox.filter((m) => Number(m.id) > cursor),
    listChats: async () => [{ id: CHAT_ID, topic: "배포방", chatType: "group", lastActivity: "2026-11-08T10:00:00Z" }],
  }
  const channel = new MicrosoftTeamsChannel({
    cfg: opts.cfg ?? configured(),
    log: NOOP_LOG,
    auth,
    graph,
    configPath: "/tmp/cfg.json",
    sessionId: "abcdef1234567890",
    push: (content, meta) => pushes.push({ content, meta }),
    onActiveChange: (a) => activeChanges.push(a),
    todoAccess: opts.todoAccess,
    now: () => clock,
  })
  const receive = (text: string, from = ME.id, id = String(++idSeq)) => {
    inbox.push({ id, messageType: "message", from: { user: { id: from, displayName: from === ME.id ? "나" : "동료" } }, body: { contentType: "html", content: `<p>${text}</p>` } })
    return id
  }
  // PermissionRequest hook 요청을 흉내낸다: 알림이 나갈 때까지 기다린 뒤 알림 본문·코드·결정 Promise 를 돌려준다.
  const ask = async (body: Parameters<MicrosoftTeamsChannel["handlePermission"]>[0], signal?: AbortSignal) => {
    const before = sent.length
    // 실제 hook 입력에는 항상 cwd 가 있다.
    const decided: Promise<HookDecision | null> = channel.handlePermission({ cwd: "/w/proj", ...body }, signal)
    await new Promise((r) => setTimeout(r, 0))
    const html = sent[before]?.html ?? ""
    const code = /<code>([pq][0-9a-z]{5})<\/code>/.exec(html)?.[1] ?? ""
    return { html, code, decided }
  }
  return {
    channel,
    sent,
    pushes,
    activeChanges,
    receive,
    ask,
    get deviceFlowStarted() {
      return deviceFlowStarted
    },
    setToken: (v: boolean) => (hasToken = v),
    setMe: (v: typeof ME) => (me = v),
    advance: (ms: number) => (clock += ms),
    failNextSend: () => (failNextSend = true),
  }
}

async function connected(opts: Parameters<typeof setup>[0] = {}) {
  const s = setup(opts)
  await s.channel.connect()
  const r = await s.channel.selectChat("1")
  expect(r.isError).toBeUndefined()
  return s
}

describe("connect / select_chat", () => {
  test("azure 설정이 비어 있으면 설정 파일 경로를 안내한다", async () => {
    const s = setup({ cfg: defaultConfig() })
    const r = await s.channel.connect()
    expect(r.isError).toBe(true)
    expect(r.text).toContain("/tmp/cfg.json")
    expect(s.deviceFlowStarted).toBe(0)
  })

  test("토큰이 없으면 device flow 를 시작하고 코드를 도구 결과로만 보여준다", async () => {
    const s = setup({ hasToken: false })
    const r = await s.channel.connect()
    expect(s.deviceFlowStarted).toBe(1)
    expect(r.text).toContain("ABCD-1234")
    expect(r.text).toContain("https://microsoft.com/devicelogin")
    expect(s.pushes).toHaveLength(0)
  })

  test("로그인 완료 이벤트는 채팅방 목록을 system 채널 메시지로 보낸다", async () => {
    const s = setup({ hasToken: false })
    await s.channel.connect()
    s.setToken(true)
    await s.channel.onLoginComplete()
    expect(s.pushes).toHaveLength(1)
    expect(s.pushes[0].meta).toEqual({ kind: "system" })
    expect(s.pushes[0].content).toContain(CHAT_ID)
  })

  test("connect 전에는 인증 이벤트를 채널로 보내지 않는다 (lazy connect)", async () => {
    const s = setup()
    s.channel.onReauthRequired()
    await s.channel.onLoginComplete()
    s.channel.onLoginFailed("x")
    expect(s.pushes).toHaveLength(0)
  })

  test("토큰 폐기 시 device flow 를 시작하지 않고 재로그인 안내를 한 번만 보낸다", async () => {
    const s = await connected()
    s.channel.onReauthRequired()
    s.channel.onReauthRequired()
    expect(s.deviceFlowStarted).toBe(0)
    expect(s.pushes).toHaveLength(1)
    expect(s.pushes[0].meta).toEqual({ kind: "system" })
    expect(s.pushes[0].content).toContain("connect 도구를 직접 호출하지 마라")
  })

  test("재로그인한 계정이 기존 계정과 다르면 연동을 해제하고 그 계정의 답신을 받지 않는다", async () => {
    const s = await connected()
    s.setMe({ id: "intruder", displayName: "동료" })
    await s.channel.onLoginComplete()
    expect(s.channel.active).toBe(false)
    expect(s.pushes.at(-1)?.content).toContain("다른 계정(동료)")
    expect(s.activeChanges).toEqual([true, false])
  })

  test("connect({ relogin: true }) 는 토큰이 있어도 device flow 를 시작하고 기존 연동 · 발신을 유지한다", async () => {
    const s = await connected()
    const r = await s.channel.connect({ relogin: true })
    expect(s.deviceFlowStarted).toBe(1)
    expect(r.text).toContain("ABCD-1234")
    expect(r.text).toContain("기존 로그인을 계속 사용")
    expect(s.channel.active).toBe(true)
    // 로그인 진행 중에도 Stop 발신은 계속되고 코드는 가린다.
    await s.channel.handleStop({ session_id: "abcdef1234567890", cwd: "/w/proj", text: "코드는 ABCD-1234" })
    expect(s.sent.at(-1)?.html).toContain("[로그인 코드 숨김]")
    expect(s.sent.at(-1)?.html).not.toContain("ABCD-1234")
  })

  test("relogin 으로 다른 계정이 로그인하면 연동을 해제한다", async () => {
    const s = await connected()
    await s.channel.connect({ relogin: true })
    s.setMe({ id: "intruder", displayName: "동료" })
    await s.channel.onLoginComplete()
    expect(s.channel.active).toBe(false)
    expect(s.pushes.at(-1)?.content).toContain("다른 계정(동료)")
  })

  test("connect 도구는 relogin 인자를 boolean true 일 때만 적용한다", async () => {
    const s = setup()
    const tool = s.channel.tools().find((t) => t.name === "connect")!
    expect(Object.keys((tool.inputSchema as { properties: object }).properties)).toEqual(["relogin"])
    await tool.handler({ relogin: "true" })
    expect(s.deviceFlowStarted).toBe(0)
    await tool.handler({ relogin: true })
    expect(s.deviceFlowStarted).toBe(1)
  })

  test("토큰이 있으면 목록을 반환하고 번호로 채팅방을 고른다", async () => {
    const s = setup()
    const list = await s.channel.connect()
    expect(list.text).toContain("1. 배포방")
    const r = await s.channel.selectChat("1")
    expect(r.text).toContain('"배포방" 연동 완료')
    expect(r.text).toContain("키워드 없이 모두 이 세션에 요청으로 전달")
    expect(r.text).toContain("--dangerously-load-development-channels")
    expect(s.channel.active).toBe(true)
    expect(s.activeChanges).toEqual([true])
  })

  test("chatId 직접 입력도 받는다", async () => {
    const s = setup()
    const r = await s.channel.selectChat(`이걸로 ${CHAT_ID} 해줘`)
    expect(r.isError).toBeUndefined()
    expect(s.channel.active).toBe(true)
  })

  test("커서 seed 조회가 실패하면 연동하지 않는다", async () => {
    const s = setup({ seed: null })
    await s.channel.connect()
    const r = await s.channel.selectChat("1")
    expect(r.isError).toBe(true)
    expect(s.channel.active).toBe(false)
  })

  test("잘못된 번호·형식은 거부한다", async () => {
    const s = setup()
    await s.channel.connect()
    expect((await s.channel.selectChat("7")).isError).toBe(true)
    expect((await s.channel.selectChat("아무거나")).isError).toBe(true)
  })

  test("disconnect 는 비활성화를 알린다", async () => {
    const s = await connected()
    s.channel.disconnect()
    expect(s.channel.active).toBe(false)
    expect(s.activeChanges).toEqual([true, false])
  })
})

describe("status 의 To Do 표시", () => {
  const todoLine = (s: ReturnType<typeof setup>) => s.channel.status().text.split("\n").find((l) => l.startsWith("To Do:"))
  const withTasks = configured({ azure: { tenantId: "t", clientId: "c", scopes: "Chat.Read Tasks.ReadWrite" } })

  test("scope 가 없으면 미사용만 표시한다", () => {
    expect(todoLine(setup())).toBe("To Do: 미사용 (scope 없음)")
  })

  test("사용 중이면 토큰의 Tasks 권한을 함께 보여준다", () => {
    expect(todoLine(setup({ cfg: withTasks, todoAccess: "readwrite", tokenScope: "Chat.Read Tasks.ReadWrite" }))).toBe(
      "To Do: 사용 (Tasks.ReadWrite) · 토큰 권한 Tasks.ReadWrite",
    )
    expect(todoLine(setup({ cfg: withTasks, todoAccess: "readwrite", tokenScope: "Chat.Read" }))).toContain("Tasks 권한 없음")
    expect(todoLine(setup({ cfg: withTasks, todoAccess: "readwrite", consent: true }))).toContain("동의 필요")
    expect(todoLine(setup({ cfg: withTasks, todoAccess: "readwrite", hasToken: false }))).toContain("로그인 필요")
  })

  test("시작 후 설정이 바뀌었으면 세션 재시작이 필요하다고 알린다", () => {
    expect(todoLine(setup({ cfg: withTasks }))).toBe("To Do: 미사용 (scope 없음) — 설정은 사용 (Tasks.ReadWrite), 세션을 다시 시작해야 적용")
    const read = configured({ azure: { tenantId: "t", clientId: "c", scopes: "Tasks.Read" } })
    expect(todoLine(setup({ cfg: read, todoAccess: "read", tokenScope: "Tasks.Read" }))).toBe("To Do: 읽기 전용 (Tasks.Read) · 토큰 권한 Tasks.Read")
  })
})

describe("Stop 발신", () => {
  test("채팅방 미선택이면 보내지 않는다", async () => {
    const s = setup()
    await s.channel.handleStop({ session_id: "s1", text: "완료", cwd: "/w/proj" })
    expect(s.sent).toHaveLength(0)
  })

  test("마지막 응답을 프로젝트·세션 헤더와 함께 보내고, 직후 중복 발화는 한 번만 보낸다", async () => {
    const s = await connected()
    await s.channel.handleStop({ session_id: "1234abcd-ffff", text: "빌드 **성공**", cwd: "/w/my-proj" })
    await s.channel.handleStop({ session_id: "1234abcd-ffff", text: "빌드 **성공**", cwd: "/w/my-proj" })
    expect(s.sent).toHaveLength(1)
    expect(s.sent[0].chatId).toBe(CHAT_ID)
    expect(s.sent[0].html).toStartWith("<b>[my-proj · 1234abcd]</b><br>")
    expect(s.sent[0].html).toContain("<b>성공</b>")
  })

  test("다음 턴에 같은 답을 하면 다시 보낸다", async () => {
    const s = await connected()
    await s.channel.handleStop({ text: "빌드 성공했습니다." })
    s.advance(60_000)
    await s.channel.handleStop({ text: "빌드 성공했습니다." })
    expect(s.sent).toHaveLength(2)
  })

  test("로그인 진행 중이면 응답에 섞인 device code 를 가리고 보낸다", async () => {
    const s = await connected()
    s.setToken(false)
    await s.channel.connect()
    await s.channel.handleStop({ text: "https://microsoft.com/devicelogin 에서 코드 ABCD-1234 를 입력하세요" })
    expect(s.sent.at(-1)?.html).not.toContain("ABCD-1234")
    expect(s.sent.at(-1)?.html).toContain("[로그인 코드 숨김]")
  })

  test("빈 응답은 보내지 않는다", async () => {
    const s = await connected()
    await s.channel.handleStop({ text: "   " })
    expect(s.sent).toHaveLength(0)
  })
})

describe("Teams 메시지 수신 (자동 전달)", () => {
  test("본인 메시지는 키워드 없이 그대로 세션에 push 하고, 다른 사용자의 메시지는 무시한다", async () => {
    const s = await connected()
    const mid = s.receive("테스트 돌려줘")
    s.receive("동료 요청", "other-user")
    await s.channel.pollOnce()
    await s.channel.pollOnce()
    expect(s.pushes).toEqual([{ content: "테스트 돌려줘", meta: { message_id: mid, sender: "나", chat: "배포방" } }])
  })

  test("예전 호출 키워드(#claude · 가재야 등)는 떼고 전달하고, 본문이 없으면 전달하지 않는다", async () => {
    const s = await connected()
    s.receive("#claude 빌드 돌려줘")
    s.receive("가재야, 로그 봐줘")
    s.receive("#agent")
    s.receive("   ")
    await s.channel.pollOnce()
    expect(s.pushes.map((p) => p.content)).toEqual(["빌드 돌려줘", "로그 봐줘"])
  })

  test("access.allowedUserIds 에 있는 사용자의 메시지는 전달한다", async () => {
    const s = await connected({ cfg: configured({ access: { allowedUserIds: ["other-user"] } }) })
    s.receive("동료 요청", "other-user")
    await s.channel.pollOnce()
    expect(s.pushes.map((p) => p.content)).toEqual(["동료 요청"])
  })

  test("메시지 하나의 처리 실패가 같은 배치의 나머지를 잃게 하지 않는다", async () => {
    const s = await connected()
    await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    s.receive("#allow")
    s.receive("다음 요청")
    s.failNextSend()
    await s.channel.pollOnce()
    expect(s.pushes.map((p) => p.content)).toEqual(["다음 요청"])
  })

  test("플러그인이 보낸 메시지는 다시 소비하지 않는다", async () => {
    const s = await connected()
    await s.channel.handleStop({ text: "로 시작하는 응답" })
    const ownId = s.sent[0].id
    s.receive("로 시작하는 응답", ME.id, String(Number(ownId)))
    await s.channel.pollOnce()
    expect(s.pushes).toHaveLength(0)
  })

  test("다른 세션(같은 계정)이 보낸 응답·알림·확인 메시지는 주입하지 않는다 (세션 간 응답 루프 방지)", async () => {
    // 다른 세션이 이 채팅방에 보낼 수 있는 모든 종류의 메시지를 만든다.
    const other = await connected()
    await other.channel.handleStop({ cwd: "/w/other", text: "다른 세션의 응답" })
    await other.ask({ tool_name: "Bash", tool_input: { command: "ls" }, agent_id: "a1", agent_type: "Explore" })
    await other.ask({ tool_name: "Edit", tool_input: { file_path: "/w/other/a.ts" } })
    other.receive("#allow")
    await other.channel.pollOnce()
    await other.ask({ tool_name: "AskUserQuestion", tool_input: { questions: [{ question: "몇 개?", options: [{ label: "하나" }] }, { question: "어디?" }] } })
    other.receive("9")
    other.receive("1")
    await other.channel.pollOnce()
    other.receive("#skip")
    other.receive("#allow-reset")
    await other.channel.pollOnce()
    const kinds = other.sent.map((m) => /^<b>\[([^·]+)/.exec(m.html)?.[1].trim())
    expect(kinds).toEqual([
      "other",
      "⚠️ Microsoft Teams: 권한 요청",
      "⚠️ Microsoft Teams: 권한 요청",
      "⚠️ 대상 지정 필요",
      "❓ Microsoft Teams: 세션 질문",
      "❌ 답변 미반영",
      "📝 답변 기록됨",
      "✅ 전달됨",
      "🔒 자동 승인 해제",
    ])

    // 세션 없이 보내는 CLI 발신(claudecode-ms-teams send)도 같은 헤더를 쓴다.
    const cli: string[] = []
    const viaCli = await sendToChat(
      {
        cfg: configured(),
        configPath: "",
        auth: { hasRefreshToken: () => true },
        graph: { sendMessage: async (_chatId, html) => (cli.push(html), "1") },
      },
      { chatId: CHAT_ID, title: "claude-nightly", text: "야간 작업 보고" },
    )
    expect(viaCli.code).toBe(0)

    const s = await connected()
    await s.ask({ tool_name: "Bash", tool_input: { command: "pwd" } })
    for (const html of [...other.sent.map((m) => m.html), ...cli]) s.receive(html.replace(/^<p>|<\/p>$/g, ""))
    await s.channel.pollOnce()
    expect(s.pushes).toHaveLength(0)
    expect(s.sent).toHaveLength(1)
  })
})

describe("권한 요청 (PermissionRequest hook)", () => {
  const SUGGESTIONS = [
    { type: "addRules", rules: [{ toolName: "Bash", ruleContent: "npm test" }], behavior: "allow", destination: "localSettings" },
  ]

  test("채팅방 미선택이면 알림 없이 결정 없음으로 끝낸다", async () => {
    const s = setup()
    expect(await s.channel.handlePermission({ tool_name: "Bash", tool_input: { command: "ls" } })).toBeNull()
    expect(s.sent).toHaveLength(0)
  })

  test("코드 없는 답장 안내와 함께 알리고, #allow 를 hook 결정으로 돌려준다", async () => {
    const s = await connected()
    const a = await s.ask({ cwd: "/w/proj", tool_name: "Bash", tool_input: { command: "ls -la <dir>", description: "목록 보기" } })
    expect(a.code).toMatch(/^p[0-9a-z]{5}$/)
    expect(a.html).toStartWith(`<b>[⚠️ Microsoft Teams: 권한 요청 · abcdef12]</b> · <code>${a.code}</code><br>`)
    expect(a.html).toContain("ls -la &lt;dir&gt;")
    expect(a.html).toContain("<b>설명</b>: 목록 보기")
    expect(a.html).toContain("/w/proj")
    expect(a.html).toContain("<li><code>#allow</code></li>")
    expect(a.html).toContain("<li><code>#allow-always</code> — 승인하고 이 세션에서 <code>Bash</code> 요청은 계속 자동 승인</li>")
    expect(a.html).toContain("<li><code>#reject</code></li>")
    expect(a.html).not.toContain(`#allow ${a.code}`)
    expect(a.html).not.toContain("Workspace 외부 경로")
    s.receive("#Allow")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "allow" })
    expect(s.sent.at(-1)?.html).toBe("<b>[✅ 전달됨 · abcdef12]</b> allow · Bash: ls -la &lt;dir&gt;")
    expect(s.pushes).toHaveLength(0)
  })

  test("#reject 는 거부로 돌려준다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Write", tool_input: { file_path: "/w/proj/a.txt", content: "x" } })
    s.receive("#reject")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "deny", message: expect.stringContaining("#reject") })
    expect(s.sent.at(-1)?.html).toBe("<b>[✅ 전달됨 · abcdef12]</b> reject · Write: /w/proj/a.txt")
  })

  test("대기 중인 요청이 여럿이면 코드 없는 답장은 적용하지 않고 코드를 붙이라고 안내한다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "rm -rf build" } })
    const b = await s.ask({ tool_name: "Edit", tool_input: { file_path: "/w/proj/a.ts" } })
    s.receive("#allow")
    await s.channel.pollOnce()
    const notice = s.sent.at(-1)?.html ?? ""
    expect(notice).toStartWith("<b>[⚠️ 대상 지정 필요 · abcdef12]</b> 대기 중인 요청이 2건입니다.")
    expect(notice).toContain(`<li><code>#allow ${a.code}</code> — Bash: rm -rf build</li>`)
    expect(notice).toContain(`<li><code>#allow ${b.code}</code> — Edit: /w/proj/a.ts</li>`)
    s.receive(`#reject ${a.code.toUpperCase()}`)
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ behavior: "deny" })
    s.receive("#allow")
    await s.channel.pollOnce()
    expect(await b.decided).toEqual({ behavior: "allow" })
  })

  test("#allow-always 는 승인하고 이후 같은 도구의 요청을 묻지 않고 자동 승인한다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "npm test" }, permission_suggestions: SUGGESTIONS })
    expect(a.html).toContain("— 승인하고 이 세션에서 <code>Bash</code> 요청은 계속 자동 승인 (Claude Code 규칙도 저장: Bash(npm test))")
    s.receive("#allow-always")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "allow", updatedPermissions: SUGGESTIONS })
    expect(s.sent.at(-1)?.html).toBe(
      "<b>[✅ 전달됨 · abcdef12]</b> allow-always · Bash: npm test<br>이 세션에서 <code>Bash</code> 요청은 계속 자동 승인합니다. 해제: <code>#allow-reset</code>",
    )
    const sentBefore = s.sent.length
    // 턴이 바뀌고 sub-agent 가 요청해도 계속 적용된다.
    await s.channel.handleStop({ text: "끝" })
    expect(await s.channel.handlePermission({ tool_name: "Bash", tool_input: { command: "rm -rf dist" } })).toEqual({ behavior: "allow" })
    expect(await s.channel.handlePermission({ tool_name: "Bash", tool_input: { command: "ls" }, agent_id: "a1" })).toEqual({ behavior: "allow" })
    expect(s.sent.length).toBe(sentBefore + 1)
    expect(s.channel.status().text).toContain("자동 승인 도구: Bash")
    // 다른 도구는 계속 묻는다.
    const edit = await s.ask({ tool_name: "Edit", tool_input: { file_path: "/w/proj/a.ts" } })
    expect(edit.code).toMatch(/^p/)
  })

  test("제안된 규칙이 없어도 #allow-always 는 세션 자동 승인을 켠다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "mcp__github__create_issue", tool_input: { title: "x" } })
    expect(a.html).not.toContain("Claude Code 규칙도 저장")
    s.receive("#allow-always")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "allow" })
    expect(await s.channel.handlePermission({ tool_name: "mcp__github__create_issue", tool_input: { title: "y" } })).toEqual({ behavior: "allow" })
  })

  test("코드 없는 #allow-always 는 대기 중인 요청이 모두 같은 도구면 함께 승인하고, 도구가 섞여 있으면 코드를 요구한다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" }, agent_id: "a1" })
    const b = await s.ask({ tool_name: "Bash", tool_input: { command: "pwd" }, agent_id: "a2" })
    s.receive("#allow-always")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "allow" })
    expect(await b.decided).toEqual({ behavior: "allow" })
    expect(s.sent.at(-1)?.html).toBe(
      "<b>[✅ 전달됨 · abcdef12]</b> allow-always · <code>Bash</code> 대기 요청 2건 승인<br>이 세션에서 <code>Bash</code> 요청은 계속 자동 승인합니다. 해제: <code>#allow-reset</code>",
    )

    const t = await connected()
    const c = await t.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    const d = await t.ask({ tool_name: "Edit", tool_input: { file_path: "/w/proj/a.ts" } })
    t.receive("#allow-always")
    await t.channel.pollOnce()
    expect(t.sent.at(-1)?.html).toContain(`<code>#allow-always ${c.code}</code>`)
    t.receive(`#allow-always ${d.code}`)
    await t.channel.pollOnce()
    expect(await d.decided).toEqual({ behavior: "allow" })
    expect(t.channel.status().text).toContain("자동 승인 도구: Edit")
  })

  test("자동 승인 중이어도 workspace 밖 경로를 건드리는 요청은 다시 묻는다", async () => {
    const s = await connected()
    const a = await s.ask({ cwd: "/w/proj", tool_name: "Read", tool_input: { file_path: "/w/proj/a.ts" } })
    s.receive("#allow-always")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "allow" })
    expect(await s.channel.handlePermission({ tool_name: "Read", tool_input: { file_path: "src/b.ts" } })).toEqual({ behavior: "allow" })
    const outside = await s.ask({ tool_name: "Read", tool_input: { file_path: "/etc/passwd" } })
    expect(outside.html).toContain("Workspace 외부 경로")
  })

  test("#allow-reset 은 자동 승인을 끄고, 연동 해제도 자동 승인을 지운다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    s.receive("#allow-always")
    await s.channel.pollOnce()
    await a.decided
    s.receive("#allow-reset")
    await s.channel.pollOnce()
    expect(s.sent.at(-1)?.html).toBe("<b>[🔒 자동 승인 해제 · abcdef12]</b> 다시 승인을 묻습니다: <code>Bash</code>")
    expect((await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })).code).toMatch(/^p/)
    s.receive("#allow-reset")
    await s.channel.pollOnce()
    expect(s.sent.at(-1)?.html).toBe("<b>[🔒 자동 승인 해제 · abcdef12]</b> 자동 승인 중인 도구가 없습니다.")

    const t = await connected()
    const b = await t.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    t.receive("#allow-always")
    await t.channel.pollOnce()
    await b.decided
    t.channel.disconnect()
    await t.channel.selectChat("1")
    expect((await t.ask({ tool_name: "Bash", tool_input: { command: "ls" } })).code).toMatch(/^p/)
  })

  test("plan 승인(ExitPlanMode)은 원래 입력을 updatedInput 으로 돌려주고, 자동 승인에도 같은 결정을 쓴다", async () => {
    const s = await connected()
    const input = { plan: "# 계획\n1. **테스트** 추가", planFilePath: "/p.md" }
    const a = await s.ask({ tool_name: "ExitPlanMode", tool_input: input })
    expect(a.html).toContain("계획 승인 요청")
    expect(a.html).toContain("<b>테스트</b> 추가")
    expect(a.html).toContain("<li><code>#allow-always</code> — 승인하고 이 세션의 이후 계획도 계속 자동 승인</li>")
    s.receive("#allow-always")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "allow", updatedInput: input })
    expect(s.sent.at(-1)?.html).toContain("allow-always · 계획 승인")
    const next = { plan: "# 다음 계획" }
    expect(await s.channel.handlePermission({ tool_name: "ExitPlanMode", tool_input: next })).toEqual({ behavior: "allow", updatedInput: next })
  })

  test("모르는 코드·형식 오류·다른 사용자의 답장은 후속 메시지 없이 무시하고 요청은 계속 대기한다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    s.receive("#allow pzzzzz")
    s.receive("#allow 지금바로")
    s.receive(`#allow ${a.code} 빨리`)
    s.receive("#allow", "other-user")
    await s.channel.pollOnce()
    expect(s.sent).toHaveLength(1)
    expect(s.pushes).toHaveLength(0)
    s.receive(`#allow ${a.code}`)
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "allow" })
  })

  test("대기 중인 요청이 없으면 답장 verb 는 조용히 무시하고 세션에도 넣지 않는다", async () => {
    const s = await connected()
    s.receive("#allow")
    s.receive("#skip")
    s.receive("#answer yes")
    await s.channel.pollOnce()
    expect(s.sent).toHaveLength(0)
    expect(s.pushes).toHaveLength(0)
  })

  test("권한 요청이 대기 중이어도 일반 메시지는 세션 요청으로 전달한다", async () => {
    const s = await connected()
    await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    s.receive("그거 말고 다른 거 해줘")
    await s.channel.pollOnce()
    expect(s.pushes.map((p) => p.content)).toEqual(["그거 말고 다른 거 해줘"])
  })

  test("workspace 밖 경로는 알림과 확인 메시지에서 강조한다", async () => {
    const s = await connected()
    const a = await s.ask({ cwd: "/w/proj", tool_name: "Read", tool_input: { file_path: "/etc/passwd" } })
    expect(a.html).toContain("Workspace 외부 경로")
    expect(a.html).toContain("<li>/etc/passwd</li>")
    s.receive("#allow")
    await s.channel.pollOnce()
    expect(s.sent.at(-1)?.html).toContain("workspace 외부 경로 승인")
  })

  test("hook 연결이 끊기면 대기를 풀고, 이후 답장은 무시한다", async () => {
    const s = await connected()
    const abort = new AbortController()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } }, abort.signal)
    abort.abort()
    expect(await a.decided).toBeNull()
    s.receive("#allow")
    s.receive(`#allow ${a.code}`)
    await s.channel.pollOnce()
    expect(s.sent).toHaveLength(1)
  })

  test("턴이 끝나면(Stop) main 의 대기는 폐기하고 sub-agent 의 대기는 유지한다", async () => {
    const s = await connected()
    const main = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    const sub = await s.ask({ tool_name: "Bash", tool_input: { command: "pwd" }, agent_id: "a1", agent_type: "Explore" })
    expect(sub.html).toStartWith("<b>[⚠️ Microsoft Teams: 권한 요청 · abcdef12]</b>")
    expect(sub.html).toContain("sub-agent <b>Explore</b>")
    await s.channel.handleStop({ text: "끝" })
    expect(await main.decided).toBeNull()
    s.receive("#allow")
    await s.channel.pollOnce()
    expect(await sub.decided).toEqual({ behavior: "allow" })
    expect(s.sent.at(-1)?.html).toContain("Bash: pwd")
  })

  test("sub-agent 가 끝나면(SubagentStop) 그 agent 의 대기만 폐기한다", async () => {
    const s = await connected()
    const a1 = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" }, agent_id: "a1" })
    const a2 = await s.ask({ tool_name: "Bash", tool_input: { command: "pwd" }, agent_id: "a2" })
    s.channel.handleSubagentStop({ agent_id: "a1" })
    expect(await a1.decided).toBeNull()
    s.receive("#allow")
    await s.channel.pollOnce()
    expect(await a2.decided).toEqual({ behavior: "allow" })
  })

  test("응답이 없는 턴 종료(StopFailure)도 main 의 대기를 폐기한다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    await s.channel.handleStop({ text: "" })
    expect(await a.decided).toBeNull()
    expect(s.sent).toHaveLength(1)
  })

  test("로그인 진행 중이면 알림에 섞인 device code 를 가린다", async () => {
    const s = await connected()
    s.setToken(false)
    await s.channel.connect()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "echo ABCD-1234" } })
    const q = await s.ask({ tool_name: "AskUserQuestion", tool_input: { questions: [{ question: "코드 ABCD-1234 를 입력했나요?", options: [{ label: "네" }] }] } })
    for (const html of [a.html, q.html]) {
      expect(html).not.toContain("ABCD-1234")
      expect(html).toContain("[로그인 코드 숨김]")
    }
  })

  test("알림 발신에 실패하면 hook 을 붙잡아 두지 않는다", async () => {
    const s = await connected()
    s.failNextSend()
    expect(await s.channel.handlePermission({ tool_name: "Bash", tool_input: { command: "ls" } })).toBeNull()
  })

  test("24시간이 지난 요청은 만료된다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    s.advance(24 * 60 * 60 * 1000 + 1)
    s.receive("#allow")
    await s.channel.pollOnce()
    expect(await a.decided).toBeNull()
    expect(s.sent).toHaveLength(1)
  })

  test("연동을 해제하면 대기 중인 hook 을 모두 풀어준다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    s.channel.disconnect()
    expect(await a.decided).toBeNull()
  })
})

describe("질문 응답 (AskUserQuestion)", () => {
  const DEPLOY = {
    header: "배포 승인",
    question: "production 에 배포할까요?",
    multiSelect: false,
    options: [{ label: "Yes", description: "지금 배포" }, { label: "No", description: "취소" }],
  }
  const TARGETS = {
    header: "배포 대상",
    question: "어느 환경에 배포할까요?",
    multiSelect: true,
    options: [{ label: "prod" }, { label: "staging" }, { label: "dev" }],
  }
  const askQuestion = (s: Awaited<ReturnType<typeof connected>>, questions: object[], extra: object = {}) =>
    s.ask({ tool_name: "AskUserQuestion", tool_input: { questions }, ...extra })

  test("선택지와 코드 없는 답장 안내를 알린다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [TARGETS])
    expect(a.code).toMatch(/^q[0-9a-z]{5}$/)
    expect(a.html).toStartWith(`<b>[❓ Microsoft Teams: 세션 질문 · abcdef12]</b> · <code>${a.code}</code><br>`)
    expect(a.html).toContain("<b>배포 대상</b> <i>(복수 선택)</i>")
    expect(a.html).toContain("<ol><li><b>prod</b></li>")
    expect(a.html).toContain("<li><code>&lt;선택지 또는 번호&gt;[, &lt;선택지2&gt;]</code> — 그대로 답장</li>")
    expect(a.html).toContain("<li><code>text: &lt;자유 텍스트&gt;</code>")
    expect(a.html).toContain("<li><code>#skip</code></li>")
    expect(a.html).not.toContain(`#answer`)
  })

  test("그냥 보낸 답장은 질문의 답이 되고 label 은 대소문자를 무시해 선택지 원문으로 돌려준다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY])
    s.receive("yes")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({
      behavior: "allow",
      updatedInput: { questions: [DEPLOY], answers: { "production 에 배포할까요?": "Yes" } },
    })
    expect(s.sent.at(-1)?.html).toBe("<b>[✅ 전달됨 · abcdef12]</b> answer · 배포 승인: Yes")
    expect(s.pushes).toHaveLength(0)
  })

  test("#answer 도 코드 없이 받고, 선택지 번호와 선택지에 없는 답(자유 텍스트)을 해석한다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY])
    s.receive("#Answer 2")
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ updatedInput: { answers: { "production 에 배포할까요?": "No" } } })
    const b = await askQuestion(s, [DEPLOY])
    s.receive("내일 오전에, 다시 물어봐")
    await s.channel.pollOnce()
    expect(await b.decided).toMatchObject({ updatedInput: { answers: { "production 에 배포할까요?": "내일 오전에, 다시 물어봐" } } })
  })

  test("범위를 벗어난 번호·빈 답은 이유를 알리고 질문을 계속 대기시킨다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [TARGETS])
    s.receive("7")
    s.receive("prod, 0")
    s.receive(",")
    await s.channel.pollOnce()
    const notices = s.sent.slice(1).map((m) => m.html)
    expect(notices).toHaveLength(3)
    expect(notices[0]).toBe(
      "<b>[❌ 답변 미반영 · abcdef12]</b> 배포 대상 — 선택지 번호는 1–3 범위여야 합니다 (숫자 자체가 답이면 text: 를 붙이세요)",
    )
    expect(notices[2]).toContain("답이 비어 있습니다")
    s.receive("text: 7")
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ updatedInput: { answers: { "어느 환경에 배포할까요?": "7" } } })
  })

  test("선택지가 없는 질문의 숫자 답은 그대로 전달한다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [{ question: "재시도 횟수는?" }])
    s.receive("7")
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ updatedInput: { answers: { "재시도 횟수는?": "7" } } })
  })

  test("번호는 알림에 보인 순서를 따르고, 전각 표기 label 도 원문으로 돌려준다", async () => {
    const s = await connected()
    const question = { question: "어느 안?", options: [{ label: "" }, { label: "Plan（A）" }, { label: "Plan B" }] }
    const a = await askQuestion(s, [question])
    s.receive("3")
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ updatedInput: { answers: { "어느 안?": "Plan B" } } })
    const b = await askQuestion(s, [question])
    s.receive("plan(a)")
    await s.channel.pollOnce()
    expect(await b.decided).toMatchObject({ updatedInput: { answers: { "어느 안?": "Plan（A）" } } })
  })

  test("복수 선택은 쉼표로 여러 선택지를 받는다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [TARGETS])
    s.receive("staging, 3")
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ updatedInput: { answers: { "어느 환경에 배포할까요?": "staging, dev" } } })
  })

  test("text: 뒤는 쉼표·개행을 포함해 그대로 자유 텍스트로 전달한다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [TARGETS])
    s.receive("text: Feat: 알림 큐, 비활성화<br>본문")
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ updatedInput: { answers: { "어느 환경에 배포할까요?": "Feat: 알림 큐, 비활성화\n본문" } } })
  })

  test("질문이 여러 개면 순서대로 하나씩 받아 모두 채워졌을 때 한 번에 돌려준다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY, TARGETS])
    expect(a.html).toContain("<b>1. 배포 승인</b>")
    expect(a.html).toContain("<b>2. 배포 대상</b> <i>(복수 선택)</i>")
    expect(a.html).toContain("— 질문 순서대로 하나씩 답장</li>")
    s.receive("yes")
    await s.channel.pollOnce()
    expect(s.sent.at(-1)?.html).toBe("<b>[📝 답변 기록됨 · abcdef12]</b> 배포 승인: Yes · 남은 질문 1개 — 다음: <b>배포 대상</b> (건너뛰기: <code>#skip</code>)")
    s.receive("prod, dev")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({
      behavior: "allow",
      updatedInput: {
        questions: [DEPLOY, TARGETS],
        answers: { "production 에 배포할까요?": "Yes", "어느 환경에 배포할까요?": "prod, dev" },
      },
    })
    expect(s.sent.at(-1)?.html).toBe("<b>[✅ 전달됨 · abcdef12]</b> answer · 배포 승인: Yes · 배포 대상: prod, dev")
  })

  test("코드와 sub-code 를 붙이면 순서와 무관하게 특정 질문에 답한다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY, TARGETS])
    s.receive(`#answer ${a.code}-2 staging`)
    s.receive(`#answer ${a.code}-3 yes`)
    await s.channel.pollOnce()
    expect(s.sent.slice(1).map((m) => m.html)).toEqual([
      "<b>[📝 답변 기록됨 · abcdef12]</b> 배포 대상: staging · 남은 질문 1개 — 다음: <b>배포 승인</b> (건너뛰기: <code>#skip</code>)",
    ])
    s.receive(`#answer ${a.code.toUpperCase()} no`)
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ updatedInput: { answers: { "production 에 배포할까요?": "No", "어느 환경에 배포할까요?": "staging" } } })
  })

  test("#skip 은 질문을 거부로 돌려준다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY, TARGETS])
    s.receive("#skip")
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "deny", message: expect.stringContaining("#skip") })
    expect(s.sent.at(-1)?.html).toBe("<b>[✅ 전달됨 · abcdef12]</b> skip · 배포 승인 · 배포 대상")
  })

  test("sub-agent 의 질문은 #answer 로만 답하고, 그냥 보낸 메시지는 세션 요청으로 전달한다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY], { agent_id: "a1", agent_type: "Plan" })
    expect(a.html).toContain("<li><code>#answer &lt;선택지 또는 번호&gt;</code> — 그대로 답장</li>")
    s.receive("새 요청")
    await s.channel.pollOnce()
    expect(s.pushes.map((p) => p.content)).toEqual(["새 요청"])
    s.receive("#answer yes")
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ updatedInput: { answers: { "production 에 배포할까요?": "Yes" } } })
  })

  test("질문이 여럿 대기 중이면 코드 없는 #answer 는 코드를 붙이라고 안내한다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY], { agent_id: "a1" })
    const b = await askQuestion(s, [TARGETS], { agent_id: "a2" })
    s.receive("#answer yes")
    await s.channel.pollOnce()
    const notice = s.sent.at(-1)?.html ?? ""
    expect(notice).toContain(`<li><code>#answer ${a.code} &lt;답&gt;</code> — 배포 승인: production 에 배포할까요?</li>`)
    expect(notice).toContain(`<li><code>#answer ${b.code} &lt;답&gt;</code> — 배포 대상: 어느 환경에 배포할까요?</li>`)
    s.receive(`#answer ${b.code} dev`)
    await s.channel.pollOnce()
    expect(await b.decided).toMatchObject({ updatedInput: { answers: { "어느 환경에 배포할까요?": "dev" } } })
  })

  test("터미널에서 먼저 답하면(PostToolUse) 대기를 폐기해 이후 Teams 답장이 다른 질문으로 간다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY])
    s.channel.handleAnswered({ tool_input: { questions: [DEPLOY], answers: { "production 에 배포할까요?": "No" } } })
    expect(await a.decided).toBeNull()
    s.receive("yes")
    await s.channel.pollOnce()
    expect(s.pushes.map((p) => p.content)).toEqual(["yes"])
    expect(s.sent).toHaveLength(1)
  })

  test("권한 verb 로는 질문에 답할 수 없다", async () => {
    const s = await connected()
    await askQuestion(s, [DEPLOY])
    s.receive("#allow")
    await s.channel.pollOnce()
    expect(s.sent).toHaveLength(1)
  })

  test("질문은 #allow-always 대상이 아니어서 매번 묻는다", async () => {
    const s = await connected()
    const a = await askQuestion(s, [DEPLOY])
    expect(a.html).not.toContain("#allow-always")
    s.receive("#allow-always")
    await s.channel.pollOnce()
    expect(s.sent).toHaveLength(1)
    s.receive("yes")
    await s.channel.pollOnce()
    expect(await a.decided).toMatchObject({ behavior: "allow" })
    expect((await askQuestion(s, [DEPLOY])).code).toMatch(/^q/)
  })

  test("질문 형식이 어긋나면 중계하지 않는다", async () => {
    const s = await connected()
    expect(await s.channel.handlePermission({ tool_name: "AskUserQuestion", tool_input: { questions: [{ header: "x" }] } })).toBeNull()
    expect(s.sent).toHaveLength(0)
  })
})

describe("리뷰 회귀 테스트", () => {
  test("코드 없는 답장은 그 답장보다 먼저 알린 요청에만 적용한다 (재전송된 #allow 가 새 요청을 승인하지 않는다)", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    s.receive("#allow")
    s.receive("#allow")
    const b = await s.ask({ tool_name: "Bash", tool_input: { command: "rm -rf ~/work" } })
    await s.channel.pollOnce()
    expect(await a.decided).toEqual({ behavior: "allow" })
    let bSettled = false
    void b.decided.then(() => (bSettled = true))
    await new Promise((r) => setTimeout(r, 0))
    expect(bSettled).toBe(false)
    expect(s.sent.at(-1)?.html).toContain("allow · Bash: ls")
  })

  test("질문이 오기 전에 보낸 메시지는 질문의 답이 아니라 세션 요청으로 전달한다", async () => {
    const s = await connected()
    s.receive("테스트도 돌려줘")
    const q = await s.ask({ tool_name: "AskUserQuestion", tool_input: { questions: [{ question: "배포할까요?", options: [{ label: "Yes" }] }] } })
    await s.channel.pollOnce()
    expect(s.pushes.map((p) => p.content)).toEqual(["테스트도 돌려줘"])
    s.receive("yes")
    await s.channel.pollOnce()
    expect(await q.decided).toMatchObject({ updatedInput: { answers: { "배포할까요?": "Yes" } } })
  })

  test("코드 없는 #allow-always 는 여러 요청 중 하나를 골라 규칙을 저장하지 않고, workspace 밖 요청은 승인하지 않는다", async () => {
    const s = await connected()
    const SUGGEST = [{ type: "addRules", rules: [{ toolName: "Read", ruleContent: "//etc/**" }], behavior: "allow", destination: "localSettings" }]
    const outside = await s.ask({ tool_name: "Read", tool_input: { file_path: "/etc/hosts" }, permission_suggestions: SUGGEST })
    const inside = await s.ask({ tool_name: "Read", tool_input: { file_path: "/w/proj/a.ts" } })
    s.receive("#allow-always")
    await s.channel.pollOnce()
    expect(await inside.decided).toEqual({ behavior: "allow" })
    let outsideSettled = false
    void outside.decided.then(() => (outsideSettled = true))
    await new Promise((r) => setTimeout(r, 0))
    expect(outsideSettled).toBe(false)
    expect(s.sent.at(-1)?.html).toContain("대기 요청 1건 승인 (workspace 밖 · 설정 경로 1건은 계속 대기)")
  })

  test("자동 승인 중이어도 설정 · git 경로, ~ 경로, Glob 의 workspace 밖 패턴은 다시 묻는다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Edit", tool_input: { file_path: "/w/proj/a.ts" } })
    s.receive("#allow-always")
    await s.channel.pollOnce()
    await a.decided
    expect(await s.channel.handlePermission({ cwd: "/w/proj", tool_name: "Edit", tool_input: { file_path: "src/b.ts" } })).toEqual({ behavior: "allow" })
    for (const file_path of ["/w/proj/.claude/settings.json", "/w/proj/.git/hooks/pre-commit", ".mcp.json", "~/.ssh/config"]) {
      const r = await s.ask({ tool_name: "Edit", tool_input: { file_path } })
      expect(r.code).toMatch(/^p/)
      expect(r.html).toContain("자동 승인 중인 도구지만")
    }
    const g = await s.ask({ tool_name: "Glob", tool_input: { pattern: "/etc/**" } })
    expect(g.html).toContain("Workspace 외부 경로")
    expect(isInsideWorkspace("~/x", "/w/proj")).toBe(false)
  })

  test("workspace 를 모르면 자동 승인하지 않는다", async () => {
    const s = await connected()
    const first = s.channel.handlePermission({ tool_name: "Bash", tool_input: { command: "ls" } })
    await new Promise((r) => setTimeout(r, 0))
    s.receive("#allow-always")
    await s.channel.pollOnce()
    expect(await first).toEqual({ behavior: "allow" })
    const second = s.channel.handlePermission({ tool_name: "Bash", tool_input: { command: "pwd" } })
    await new Promise((r) => setTimeout(r, 0))
    expect(s.sent.at(-1)?.html).toContain("workspace 를 알 수 없음")
    s.channel.disconnect()
    expect(await second).toBeNull()
  })

  test("확인 · 대상 지정 메시지도 로그인 코드를 가린다", async () => {
    const s = await connected()
    s.setToken(false)
    await s.channel.connect()
    await s.ask({ tool_name: "Bash", tool_input: { command: "echo ABCD-1234" } })
    await s.ask({ tool_name: "Bash", tool_input: { command: "echo 2" } })
    s.receive("#allow")
    await s.channel.pollOnce()
    expect(s.sent.at(-1)?.html).toContain("대상 지정 필요")
    for (const m of s.sent) expect(m.html).not.toContain("ABCD-1234")
  })

  test("질문 대기 중 명령 오타(#skp)는 답으로 쓰지 않는다", async () => {
    const s = await connected()
    const q = await s.ask({ tool_name: "AskUserQuestion", tool_input: { questions: [{ question: "배포할까요?" }] } })
    s.receive("#skp")
    await s.channel.pollOnce()
    expect(s.sent).toHaveLength(1)
    s.receive("#skip")
    await s.channel.pollOnce()
    expect(await q.decided).toMatchObject({ behavior: "deny" })
  })

  test("다른 채팅방으로 바꾸면 이전 대기 요청과 자동 승인을 이어받지 않는다", async () => {
    const s = await connected()
    const a = await s.ask({ tool_name: "Bash", tool_input: { command: "ls" } })
    s.receive("#allow-always")
    await s.channel.pollOnce()
    await a.decided
    const b = await s.ask({ tool_name: "Edit", tool_input: { file_path: "/w/proj/a.ts" } })
    await s.channel.selectChat(CHAT_ID)
    expect(await b.decided).toBeNull()
    expect(s.channel.status().text).toContain("자동 승인 도구: -")
  })

  test("세션에 넘기는 본문과 text: 답은 NFKC 로 바꾸지 않는다", async () => {
    const s = await connected()
    s.receive("ㅇㅇ ①번 해줘")
    s.receive("＃claude　x² 계산")
    await s.channel.pollOnce()
    expect(s.pushes.map((p) => p.content)).toEqual(["ㅇㅇ ①번 해줘", "x² 계산"])
    const q = await s.ask({ tool_name: "AskUserQuestion", tool_input: { questions: [{ question: "어때?", options: [{ label: "좋아" }] }] } })
    s.receive("#answer text: ㅇㅇ ①")
    await s.channel.pollOnce()
    expect(await q.decided).toMatchObject({ updatedInput: { answers: { "어때?": "ㅇㅇ ①" } } })
  })

  test("제목에 전각 대괄호가 있어도 다른 세션이 응답을 주입하지 않는다", async () => {
    const other = await connected()
    await other.channel.handleStop({ cwd: "/w/proj］x", text: "응답" })
    const s = await connected()
    s.receive(other.sent[0].html)
    await s.channel.pollOnce()
    expect(s.pushes).toHaveLength(0)
  })
})
