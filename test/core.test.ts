import { afterEach, beforeEach, describe, expect, test } from "bun:test"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"
import {
  Auth,
  azureConfigured,
  defaultConfig,
  delayFor,
  detectTrigger,
  Graph,
  htmlToText,
  mdToHtml,
  mergeConfig,
  NOOP_LOG,
  generateReplyCode,
  isInsideWorkspace,
  isPluginMessage,
  cleanInbound,
  payloadVerbatim,
  parseReplyCommand,
  renderHeader,
  renderOutbound,
  shortSessionId,
  todoAccess,
  TRIGGER_RE,
  type Config,
} from "../src/core"

describe("detectTrigger", () => {
  test("#claude 키워드 뒤 payload 를 추출한다", () => {
    const t = detectTrigger("#claude 빌드 돌려줘")
    expect(t).toEqual({ matched: true, keyword: "#claude", matchedSpan: "#claude ", payload: "빌드 돌려줘" })
  })

  test("전각 ＃ 과 NBSP 를 정규화해 매칭한다", () => {
    const t = detectTrigger("＃claude 테스트")
    expect(t.matched && t.keyword).toBe("#claude")
    expect(t.matched && t.payload).toBe("테스트")
  })

  test("대소문자 무시 + 구분자 연속 소비", () => {
    const t = detectTrigger("#Claude:: 안녕")
    expect(t.matched && t.keyword).toBe("#claude")
    expect(t.matched && t.payload).toBe("안녕")
  })

  test("한글 호출어도 트리거로 인식한다", () => {
    expect(detectTrigger("가재야 로그 봐줘").matched).toBe(true)
    expect(detectTrigger("막둥아, 커밋해줘").matched).toBe(true)
  })

  test("키워드 뒤에 경계 문자가 없거나 문장 중간이면 매칭하지 않는다", () => {
    expect(detectTrigger("#claudex 해줘").matched).toBe(false)
    expect(detectTrigger("이거 #claude 해줘").matched).toBe(false)
    expect(detectTrigger("그냥 잡담").matched).toBe(false)
  })

  test("#allow-always · #allow-reset 은 #allow 보다 우선 매칭된다", () => {
    const t = detectTrigger("#allow-always pa7k9m")
    expect(t.matched && t.keyword).toBe("#allow-always")
    const r = detectTrigger("#Allow-Reset")
    expect(r.matched && [r.keyword, r.payload]).toEqual(["#allow-reset", ""])
  })

  test("opencode-ms-teams 스펙의 키워드와 예시를 그대로 인식한다", () => {
    const payload = (raw: string) => {
      const t = detectTrigger(raw)
      return t.matched ? t.payload : null
    }
    expect(payload("#ask 카산드라 compaction 전략 알려줘")).toBe("카산드라 compaction 전략 알려줘")
    expect(payload("막둥아, 이 로그 좀 봐줘")).toBe("이 로그 좀 봐줘")
    expect(payload("가재야!! 배포 상태?")).toBe("배포 상태?")
    expect(payload("#AI\n다음 코드를 리뷰해줘")).toBe("다음 코드를 리뷰해줘")
    expect(payload("#Question 이건 뭐야")).toBe("이건 뭐야")
    expect(payload("＃agent 확인해줘")).toBe("확인해줘")
    expect(payload("#ask #agent 질문")).toBe("#agent 질문")
    expect(payload("#agent")).toBe("")
    expect(payload("그냥 잡담이야")).toBeNull()
    expect(payload("막둥아니야")).toBeNull()
    expect(payload("#asking 뭐해")).toBeNull()
  })

  test("답장 verb (#answer / #skip) 도 최상위 키워드로 매칭된다", () => {
    const answer = detectTrigger("#Answer q3x9m2 yes")
    expect(answer.matched && [answer.keyword, answer.payload]).toEqual(["#answer", "q3x9m2 yes"])
    const skip = detectTrigger("#skip q3x9m2")
    expect(skip.matched && skip.keyword).toBe("#skip")
  })

  test("TRIGGER_RE alternation 은 길이 내림차순이다", () => {
    const alts = /#\(\?:([^)]*)\)/.exec(TRIGGER_RE.source)![1].split("|")
    const sorted = [...alts].sort((a, b) => b.length - a.length)
    expect(alts.map((a) => a.length)).toEqual(sorted.map((a) => a.length))
  })
})

describe("parseReplyCommand", () => {
  test("권한 verb 는 코드 없이 받고, 코드를 붙이면 대상을 지정한다 (대문자는 소문자로)", () => {
    expect(parseReplyCommand("#allow", "")).toEqual({ verb: "allow", code: null })
    expect(parseReplyCommand("#allow-always", "")).toEqual({ verb: "allow-always", code: null })
    expect(parseReplyCommand("#skip", "")).toEqual({ verb: "skip", code: null })
    expect(parseReplyCommand("#allow", "pa7k9m")).toEqual({ verb: "allow", code: "pa7k9m" })
    expect(parseReplyCommand("#reject", "PA7K9M")).toEqual({ verb: "reject", code: "pa7k9m" })
    expect(parseReplyCommand("#allow-always", "p00000")).toEqual({ verb: "allow-always", code: "p00000" })
    expect(parseReplyCommand("#skip", "q3x9m2")).toEqual({ verb: "skip", code: "q3x9m2" })
    expect(parseReplyCommand("#allow-reset", "")).toEqual({ verb: "allow-reset" })
  })

  test("#answer 는 코드 없이 답만 받고, 자유 텍스트는 text: 로 받는다", () => {
    expect(parseReplyCommand("#answer", "yes")).toEqual({ verb: "answer", code: null, subIndex: null, value: "yes", freeText: false })
    expect(parseReplyCommand("#answer", "Text : Feat: 큐, 비활성화\n\n본문")).toEqual({
      verb: "answer",
      code: null,
      subIndex: null,
      value: "Feat: 큐, 비활성화\n\n본문",
      freeText: true,
    })
  })

  test("#answer 의 첫 단어는 대기 중인 질문 코드일 때만 코드로 본다", () => {
    const known = new Set(["q3x9m2", "qa9m2k", "q7bnp3"])
    expect(parseReplyCommand("#answer", "q3x9m2 yes", known)).toEqual({ verb: "answer", code: "q3x9m2", subIndex: null, value: "yes", freeText: false })
    expect(parseReplyCommand("#answer", "QA9M2K-2 staging, dev", known)).toEqual({
      verb: "answer",
      code: "qa9m2k",
      subIndex: 2,
      value: "staging, dev",
      freeText: false,
    })
    expect(parseReplyCommand("#answer", "q7bnp3 text: 7", known)).toEqual({ verb: "answer", code: "q7bnp3", subIndex: null, value: "7", freeText: true })
    // 코드 형식(p/q + 5자)과 겹치는 답도 그대로 답으로 받는다.
    expect(parseReplyCommand("#answer", "python", known)).toEqual({ verb: "answer", code: null, subIndex: null, value: "python", freeText: false })
    expect(parseReplyCommand("#answer", "q3x9m2 yes")).toEqual({ verb: "answer", code: null, subIndex: null, value: "q3x9m2 yes", freeText: false })
  })

  test("코드 형식 불일치·추가 텍스트·빈 답·비 verb 는 거부한다", () => {
    expect(parseReplyCommand("#allow", "abcde")).toBeNull()
    expect(parseReplyCommand("#allow", "pa7k9")).toBeNull()
    expect(parseReplyCommand("#allow", "xa7k9m")).toBeNull()
    expect(parseReplyCommand("#allow", "pa7k9m 지금")).toBeNull()
    expect(parseReplyCommand("#allow", "지금 바로")).toBeNull()
    expect(parseReplyCommand("#allow", "pa7k9m-1")).toBeNull()
    expect(parseReplyCommand("#skip", "q3x9m2-1")).toBeNull()
    expect(parseReplyCommand("#allow-reset", "pa7k9m")).toBeNull()
    expect(parseReplyCommand("#answer", "")).toBeNull()
    expect(parseReplyCommand("#answer", "text:  ")).toBeNull()
    expect(parseReplyCommand("#answer", "q3x9m2", new Set(["q3x9m2"]))).toBeNull()
    expect(parseReplyCommand("#answer", "q3x9m2 text:  ", new Set(["q3x9m2"]))).toBeNull()
    expect(parseReplyCommand("#claude", "pa7k9m")).toBeNull()
  })
})

describe("isPluginMessage", () => {
  test("플러그인이 보내는 헤더([제목 · 세션 8자])로 시작하는 메시지를 알아본다", () => {
    const echo = (html: string) => isPluginMessage(htmlToText(html))
    expect(echo(renderOutbound("my-proj", "1234abcd", "#claude 로 시작하는 응답"))).toBe(true)
    expect(echo(renderOutbound("Claude Code", "-", "응답"))).toBe(true)
    expect(echo(`${renderHeader("✅ 전달됨", "1234abcd")} allow`)).toBe(true)
    expect(echo("<b>[OpenCode · 6a1b2c3d]</b><br>응답")).toBe(true)
    // Teams 가 숫자 엔티티로 돌려줘도 알아본다.
    expect(echo("&#91;proj &#183; 1234abcd&#x5D; 응답")).toBe(true)
    // Teams 가 이모지를 <emoji> 태그로 바꿔 돌려줘도 헤더 형식은 남는다.
    expect(echo('<p><b>[<emoji alt="✅"></emoji> 전달됨 · 1234abcd]</b> allow</p>')).toBe(true)
  })

  test("제목에 대괄호·개행(전각 포함)이 있어도 헤더 형식을 유지한다", () => {
    expect(isPluginMessage(htmlToText(renderOutbound("a]b[c\nd", "1234abcd", "x")))).toBe(true)
    expect(isPluginMessage(htmlToText(renderOutbound("proj］x﹈y", "1234abcd", "x")))).toBe(true)
  })

  test("사용자가 쓴 일반 메시지는 플러그인 메시지로 보지 않는다", () => {
    expect(isPluginMessage("테스트 돌려줘")).toBe(false)
    expect(isPluginMessage("[긴급 · 서버] 로그 봐줘")).toBe(false)
    expect(isPluginMessage("[TODO] 정리 · 1234abcd]")).toBe(false)
    expect(isPluginMessage("앞에 글 [proj · 1234abcd]")).toBe(false)
    expect(isPluginMessage("[회의록 · 20261120] 요약해줘")).toBe(false)
    expect(isPluginMessage("[TODO · refactor] 정리해줘")).toBe(false)
  })

  test("shortSessionId 는 헤더 형식에 맞는 8자 또는 - 를 돌려준다", () => {
    expect(shortSessionId("1234abcd-ffff-4444")).toBe("1234abcd")
    expect(shortSessionId("abc")).toBe("-")
    expect(shortSessionId("")).toBe("-")
    expect(shortSessionId("한글세션아이디입니다")).toBe("-")
    expect(shortSessionId("ABCDEF12-x")).toBe("abcdef12")
    // 숫자만이면 사용자 메시지(날짜 등)와 겹치므로 쓰지 않는다.
    expect(shortSessionId("12345678-ffff")).toBe("-")
    expect(shortSessionId("ses_long_opencode_id")).toBe("-")
  })

  test("payloadVerbatim 은 키워드만 떼고 본문은 원문 문자 그대로 돌려준다", () => {
    const strip = (raw: string) => {
      const t = detectTrigger(raw)
      return t.matched ? payloadVerbatim(raw, t) : null
    }
    expect(strip("#claude ㅇㅇ ①")).toBe("ㅇㅇ ①")
    expect(strip("＃claude, x²")).toBe("x²")
    expect(strip("가재야!! 배포 상태?")).toBe("배포 상태?")
    expect(strip("#answer text: ㅇㅇ")).toBe("text: ㅇㅇ")
    expect(cleanInbound("\u200B a\u00A0b\r\nc\u3000 ")).toBe("a b\nc")
  })

  test("htmlToText 는 숫자 엔티티를 풀되 &amp; 는 한 번만 푼다", () => {
    expect(htmlToText("&#91;a &#183; b&#x5D;")).toBe("[a · b]")
    expect(htmlToText("&amp;#183;")).toBe("&#183;")
  })
})

describe("generateReplyCode", () => {
  test("접두 + base36 5자를 만들고, 이미 쓰는 코드는 피한다", () => {
    expect(generateReplyCode("p", new Set())).toMatch(/^p[0-9a-z]{5}$/)
    expect(generateReplyCode("q", new Set())).toMatch(/^q[0-9a-z]{5}$/)
    const seq = [0, 0, 0, 0, 0, 0.5, 0.5, 0.5, 0.5, 0.5]
    expect(generateReplyCode("p", new Set(["p00000"]), () => seq.shift() ?? 0)).toBe("piiiii")
    expect(generateReplyCode("p", new Set(["p00000"]), () => 0)).toBeNull()
  })
})

describe("isInsideWorkspace", () => {
  test("상대 경로와 .. 를 정규화해 workspace 하위인지 판정한다", () => {
    expect(isInsideWorkspace("/w/proj/src/a.ts", "/w/proj")).toBe(true)
    expect(isInsideWorkspace("src/a.ts", "/w/proj")).toBe(true)
    expect(isInsideWorkspace("/w/proj", "/w/proj")).toBe(true)
    expect(isInsideWorkspace("/w/proj/..data/a", "/w/proj")).toBe(true)
    expect(isInsideWorkspace("/w/proj/../other/a.ts", "/w/proj")).toBe(false)
    expect(isInsideWorkspace("/etc/passwd", "/w/proj")).toBe(false)
    expect(isInsideWorkspace("/w/proj-2/a.ts", "/w/proj")).toBe(false)
    expect(isInsideWorkspace("", "/w/proj")).toBe(false)
    expect(isInsideWorkspace("~/.ssh/id_rsa", "/w/proj")).toBe(false)
  })
})

describe("config", () => {
  test("access.allowedUserIds 는 문자열만 남기고 없으면 빈 배열", () => {
    const base = defaultConfig()
    expect(mergeConfig(base, {}).access.allowedUserIds).toEqual([])
    expect(mergeConfig(base, { access: { allowedUserIds: ["u1", 3, "", "u2"] } }).access.allowedUserIds).toEqual(["u1", "u2"])
    expect(mergeConfig(base, { access: { allowedUserIds: "u1" } }).access.allowedUserIds).toEqual([])
  })

  test("부분 설정은 기본값 위에 병합된다", () => {
    const cfg = mergeConfig(defaultConfig(), { azure: { tenantId: "t" }, polling: { intervalMs: 1000 } })
    expect(cfg.azure).toEqual({ tenantId: "t", clientId: "", scopes: "" })
    expect(cfg.polling.intervalMs).toBe(1000)
    expect(cfg.polling.backoff1Ms).toBe(defaultConfig().polling.backoff1Ms)
    expect(azureConfigured(cfg)).toBe(false)
    expect(azureConfigured(mergeConfig(cfg, { azure: { tenantId: "t", clientId: "c", scopes: "s" } }))).toBe(true)
  })

  test("todo 섹션은 timeZone 문자열만 받고 없으면 빈 값(시스템 시간대)이다", () => {
    expect(defaultConfig().todo).toEqual({ timeZone: "" })
    expect(mergeConfig(defaultConfig(), { todo: { timeZone: " Asia/Seoul " } }).todo).toEqual({ timeZone: "Asia/Seoul" })
    expect(mergeConfig(defaultConfig(), { todo: { timeZone: 9 } }).todo).toEqual({ timeZone: "" })
    expect(mergeConfig(defaultConfig(), { todo: "Asia/Seoul" }).todo).toEqual({ timeZone: "" })
  })

  test("todoAccess 는 Tasks.ReadWrite > Tasks.Read > 없음 순으로 판정한다 (대소문자 · resource 접두 무시)", () => {
    expect(todoAccess("Chat.Read offline_access Tasks.ReadWrite User.Read")).toBe("readwrite")
    expect(todoAccess("Tasks.Read Tasks.ReadWrite")).toBe("readwrite")
    expect(todoAccess("Chat.Read tasks.read")).toBe("read")
    expect(todoAccess("https://graph.microsoft.com/Tasks.ReadWrite openid")).toBe("readwrite")
    expect(todoAccess("Chat.Read Tasks.ReadWrite.Shared")).toBe("none")
    expect(todoAccess("Chat.Read User.Read")).toBe("none")
    expect(todoAccess("")).toBe("none")
    expect(todoAccess(null)).toBe("none")
  })

  test("delayFor 는 경과 시간에 따라 1x/3x/6x 로 늘어난다", () => {
    const cfg = defaultConfig()
    expect(delayFor(0, cfg)).toBe(5000)
    expect(delayFor(cfg.polling.backoff1Ms, cfg)).toBe(15000)
    expect(delayFor(cfg.polling.backoff2Ms, cfg)).toBe(30000)
  })
})

describe("rendering", () => {
  test("코드 블록은 class 없는 <pre><code> 로 감싼다", () => {
    const html = mdToHtml("```ts\nconst a = 1 < 2\n```")
    expect(html).toContain("<pre style=")
    expect(html).toContain("<code>const a = 1 &lt; 2</code>")
    expect(html).not.toContain("class=")
  })

  test("표·볼드·인라인 코드·개행을 변환한다", () => {
    const html = mdToHtml("**굵게** `x`\n| a | b |\n|---|---:|\n| 1 | 2 |")
    expect(html).toContain("<b>굵게</b>")
    expect(html).toContain("<table")
    expect(html).toContain('<td align="right">2</td>')
  })

  test("renderOutbound 는 헤더를 붙이고 최대 크기를 넘으면 자른다", () => {
    const html = renderOutbound("proj", "abcd1234", "가".repeat(20_000), 5_000)
    expect(html.startsWith("<b>[proj · abcd1234]</b><br>")).toBe(true)
    expect(Buffer.byteLength(html, "utf8")).toBeLessThanOrEqual(5_000)
    expect(html.endsWith("…(생략)")).toBe(true)
  })

  test("htmlToText 는 태그를 벗기고 엔티티를 한 번만 디코드한다", () => {
    expect(htmlToText("<p>#claude a&amp;lt;b</p><br>x")).toBe("#claude a&lt;b\n\nx")
  })

  test("렌더 결과처럼 생긴 원문 HTML 도 escape 한다", () => {
    const html = mdToHtml('<pre style="x"><code><a href="https://evil">클릭</a></code></pre>\n`<table>` 와 `</table>` 사용')
    expect(html).not.toContain("<a ")
    expect(html).not.toContain("<table>")
    expect(html).toContain("&lt;a href=")
    expect(html).toContain("&lt;table&gt;")
  })

  test("코드 블록 안의 백틱·볼드 기호는 인라인 변환하지 않는다", () => {
    const html = mdToHtml("앞 `a`\n```\n**x** `y`\n```\n뒤 **b**")
    expect(html).toContain("<code>**x** `y`</code>")
    expect(html).toContain("<b>b</b>")
  })
})

describe("Auth / Graph (fetch stub)", () => {
  const realFetch = globalThis.fetch
  let tmp: string

  beforeEach(() => {
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), "ccmt-auth-"))
  })

  afterEach(() => {
    globalThis.fetch = realFetch
    fs.rmSync(tmp, { recursive: true, force: true })
  })

  function cfgWithTokens(tokens: object | null): Config {
    const cfg = mergeConfig(defaultConfig(), {
      azure: { tenantId: "t", clientId: "c", scopes: "s offline_access" },
      tokens: { file: path.join(tmp, "tokens.json") },
    })
    if (tokens) fs.writeFileSync(cfg.tokens.file, JSON.stringify(tokens))
    return cfg
  }

  const stubFetch = (handler: (url: string) => Response) => {
    const urls: string[] = []
    globalThis.fetch = (async (input: string | URL | Request) => {
      urls.push(String(input))
      return handler(String(input))
    }) as typeof fetch
    return urls
  }

  test("Graph 가 HTML 본문을 받아도 throw 하지 않고 null 을 돌려준다", async () => {
    const cfg = cfgWithTokens({ access_token: "a", refresh_token: "r", expires_at: Date.now() + 3600_000 })
    stubFetch(() => new Response("<html>proxy</html>", { status: 200, headers: { "Content-Type": "text/html" } }))
    const graph = new Graph(new Auth(cfg, NOOP_LOG), NOOP_LOG)
    expect(await graph.sendMessage("19:x@thread.v2", "<b>hi</b>")).toBeNull()
    expect(await graph.latestMessageId("19:x@thread.v2")).toBeNull()
    expect(await graph.messagesSince("19:x@thread.v2", 0)).toBeNull()
  })

  test("Graph.request 는 실패 상태 코드와 Graph 오류 본문을 돌려주고, call 경로는 여전히 null 이다", async () => {
    const cfg = cfgWithTokens({ access_token: "a", refresh_token: "r", expires_at: Date.now() + 3600_000 })
    const status = { code: 403 }
    stubFetch(() => Response.json({ error: { code: "Forbidden", message: "Access denied" } }, { status: status.code }))
    const graph = new Graph(new Auth(cfg, NOOP_LOG), NOOP_LOG)
    for (const code of [400, 403, 404]) {
      status.code = code
      expect(await graph.request("GET", "https://graph.microsoft.com/v1.0/me/todo/lists")).toEqual({
        ok: false,
        status: code,
        body: { error: { code: "Forbidden", message: "Access denied" } },
        error: "http",
      })
    }
    expect(await graph.me()).toBeNull()
  })

  test("Graph.request 는 401 이면 토큰을 갱신해 한 번 재시도하고, 204 는 빈 본문으로 성공한다", async () => {
    const cfg = cfgWithTokens({ access_token: "a", refresh_token: "r", expires_at: Date.now() + 3600_000 })
    let graphCalls = 0
    const urls = stubFetch((url) => {
      if (url.includes("/oauth2/")) return Response.json({ access_token: "b", expires_in: 3600 })
      graphCalls++
      return graphCalls === 1 ? new Response("", { status: 401 }) : new Response(null, { status: 204 })
    })
    const auth = new Auth(cfg, NOOP_LOG)
    const graph = new Graph(auth, NOOP_LOG)
    expect(await graph.request("DELETE", "https://graph.microsoft.com/v1.0/x")).toEqual({ ok: true, status: 204, body: {} })
    expect(urls.filter((u) => u.includes("/oauth2/"))).toHaveLength(1)
    expect(graphCalls).toBe(2)
    auth.dispose()
  })

  test("Graph.request 는 401 재시도에도 토큰을 얻지 못하면 auth 오류다", async () => {
    const cfg = cfgWithTokens({ access_token: "a", refresh_token: "r", expires_at: Date.now() + 3600_000 })
    stubFetch((url) => (url.includes("/oauth2/") ? Response.json({ error: "temporarily_unavailable" }, { status: 503 }) : new Response("", { status: 401 })))
    const auth = new Auth(cfg, NOOP_LOG)
    const r = await new Graph(auth, NOOP_LOG).request("GET", "https://graph.microsoft.com/v1.0/x")
    expect(r).toMatchObject({ ok: false, status: 401, error: "auth" })
    auth.dispose()
  })

  test("Graph.request 는 429 를 Retry-After 만큼 기다려 한 번만 재시도한다", async () => {
    const cfg = cfgWithTokens({ access_token: "a", refresh_token: "r", expires_at: Date.now() + 3600_000 })
    let calls = 0
    stubFetch(() => {
      calls++
      return new Response("", { status: 429, headers: { "Retry-After": "0" } })
    })
    const graph = new Graph(new Auth(cfg, NOOP_LOG), NOOP_LOG)
    expect(await graph.request("GET", "https://graph.microsoft.com/v1.0/x")).toMatchObject({ ok: false, status: 429, error: "http" })
    expect(calls).toBe(2)
  })

  test("Graph.request 는 네트워크 오류와 인증서 오류를 구분한다", async () => {
    const cfg = cfgWithTokens({ access_token: "a", refresh_token: "r", expires_at: Date.now() + 3600_000 })
    const graph = new Graph(new Auth(cfg, NOOP_LOG), NOOP_LOG)
    globalThis.fetch = (async () => {
      throw new TypeError("fetch failed", { cause: new Error("ECONNRESET") })
    }) as unknown as typeof fetch
    expect(await graph.request("GET", "https://graph.microsoft.com/v1.0/x")).toMatchObject({ ok: false, status: 0, error: "network" })
    globalThis.fetch = (async () => {
      throw new TypeError("fetch failed", { cause: new Error("UNABLE_TO_VERIFY_LEAF_SIGNATURE") })
    }) as unknown as typeof fetch
    expect(await graph.request("GET", "https://graph.microsoft.com/v1.0/x")).toMatchObject({ ok: false, status: 0, error: "ssl" })
  })

  test("invalid_grant 이면 device flow 없이 재로그인 이벤트를 한 번 내고 이후 토큰 요청을 멈춘다", async () => {
    const cfg = cfgWithTokens({ access_token: "a", refresh_token: "r", expires_at: Date.now() - 1000 })
    const urls = stubFetch(() => Response.json({ error: "invalid_grant" }, { status: 400 }))
    let reauth = 0
    const auth = new Auth(cfg, NOOP_LOG, { onReauthRequired: () => reauth++ })
    expect(await auth.getAccessToken()).toBeNull()
    expect(await auth.getAccessToken()).toBeNull()
    expect(await auth.forceRefresh()).toBeNull()
    expect(reauth).toBe(1)
    expect(urls).toHaveLength(1)
    expect(urls.some((u) => u.includes("devicecode"))).toBe(false)
    expect(auth.hasRefreshToken()).toBe(false)
    expect(auth.pendingDeviceCode()).toBeNull()
    auth.dispose()
  })

  test("refresh 응답에 refresh_token 이 없으면 기존 값을 유지하고 tmp 파일을 남기지 않는다", async () => {
    const cfg = cfgWithTokens({ access_token: "old", refresh_token: "keep-me", expires_at: Date.now() - 1000 })
    stubFetch(() => Response.json({ access_token: "new", expires_in: 3600 }))
    const auth = new Auth(cfg, NOOP_LOG)
    expect(await auth.getAccessToken()).toBe("new")
    const saved = JSON.parse(fs.readFileSync(cfg.tokens.file, "utf8"))
    expect(saved.refresh_token).toBe("keep-me")
    expect(saved.access_token).toBe("new")
    expect(fs.readdirSync(tmp)).toEqual(["tokens.json"])
    expect((fs.statSync(cfg.tokens.file).mode & 0o777).toString(8)).toBe("600")
    auth.dispose()
  })
})
