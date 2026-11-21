import { describe, expect, test } from "bun:test"
import { defaultConfig, htmlToText, isPluginMessage, mergeConfig, type Config } from "../src/core"
import { parseSendArgs, SEND_EXIT, sendToChat, type SendPorts } from "../src/send"

const CHAT_ID = "19:abc123@thread.v2"
const MID = "1700000000001"

function configured(): Config {
  return mergeConfig(defaultConfig(), { azure: { tenantId: "t", clientId: "c", scopes: "s" } })
}

function fakes(opts: { hasToken?: boolean; sendFails?: boolean; revokeOnSend?: boolean } = {}) {
  let hasToken = opts.hasToken ?? true
  const sent: Array<{ chatId: string; html: string }> = []
  const ports: SendPorts = {
    auth: { hasRefreshToken: () => hasToken },
    graph: {
      sendMessage: async (chatId, html) => {
        if (opts.revokeOnSend) hasToken = false
        if (opts.sendFails || opts.revokeOnSend) return null
        sent.push({ chatId, html })
        return MID
      },
    },
  }
  return { ports, sent }
}

const send = (ports: SendPorts, text: string, cfg = configured()) =>
  sendToChat({ ...ports, cfg, configPath: "/cfg/claudecode-ms-teams.json" }, { chatId: CHAT_ID, title: "claude-nightly", text })

describe("parseSendArgs", () => {
  test("--chat · --file · --title 을 읽고, --file - 와 생략은 stdin 으로 본다", () => {
    expect(parseSendArgs(["--chat", CHAT_ID, "--file", "r.md", "--title", "nightly"])).toEqual({ chatId: CHAT_ID, file: "r.md", title: "nightly" })
    expect(parseSendArgs(["--file", "-", "--chat", CHAT_ID])).toEqual({ chatId: CHAT_ID, file: null, title: null })
    expect(parseSendArgs(["--chat", "19:x_y@unq.gbl.spaces"])).toEqual({ chatId: "19:x_y@unq.gbl.spaces", file: null, title: null })
  })

  test("chatId 가 없거나 형식 전체가 맞지 않으면, 모르는 옵션 · 값 없는 옵션은 오류", () => {
    for (const argv of [
      [],
      ["--chat"],
      ["--chat", "19:abc"],
      ["--chat", `x${CHAT_ID}`],
      ["--chat", `${CHAT_ID} extra`],
      ["--chat", CHAT_ID, "--file"],
      ["--chat", CHAT_ID, "--file", ""],
      ["--chat", CHAT_ID, "--verbose"],
      ["--chat", CHAT_ID, "r.md"],
    ]) {
      expect(parseSendArgs(argv)).toHaveProperty("error")
    }
  })
})

describe("sendToChat", () => {
  test("플러그인 헤더(세션 자리 -)를 붙여 보내고 메시지 id 와 함께 성공을 돌려준다", async () => {
    const { ports, sent } = fakes()
    const r = await send(ports, "## 야간 작업\n- 빌드 **성공**")
    expect(r.code).toBe(SEND_EXIT.ok)
    expect(r.message).toContain(MID)
    expect(sent).toHaveLength(1)
    expect(sent[0].chatId).toBe(CHAT_ID)
    expect(sent[0].html).toStartWith("<b>[claude-nightly · -]</b><br>")
    expect(sent[0].html).toContain("<b>성공</b>")
    // 같은 채팅방에 연동된 세션이 이 메시지를 사용자 요청으로 주입하지 않는다.
    expect(isPluginMessage(htmlToText(sent[0].html))).toBe(true)
  })

  test("빈 본문은 보내지 않고 사용법 오류로 끝낸다", async () => {
    const { ports, sent } = fakes()
    expect((await send(ports, " \n\t")).code).toBe(SEND_EXIT.usage)
    expect(sent).toHaveLength(0)
  })

  test("azure 설정이나 로그인 토큰이 없으면 보내지 않고 로그인 필요로 끝낸다 (device flow 를 시작하지 않는다)", async () => {
    const noAzure = fakes()
    const r1 = await send(noAzure.ports, "보고", defaultConfig())
    expect(r1.code).toBe(SEND_EXIT.login)
    expect(r1.message).toContain("/cfg/claudecode-ms-teams.json")
    expect(noAzure.sent).toHaveLength(0)

    const noToken = fakes({ hasToken: false })
    const r2 = await send(noToken.ports, "보고")
    expect(r2.code).toBe(SEND_EXIT.login)
    expect(r2.message).toContain("/microsoft-teams:connect")
    expect(noToken.sent).toHaveLength(0)
  })

  test("발신 중 refresh token 이 폐기되면 로그인 필요, 그 외 발신 실패는 실패로 구분한다", async () => {
    const revoked = await send(fakes({ revokeOnSend: true }).ports, "보고")
    expect(revoked.code).toBe(SEND_EXIT.login)
    expect(revoked.message).toContain("/microsoft-teams:connect")
    expect((await send(fakes({ sendFails: true }).ports, "보고")).code).toBe(SEND_EXIT.failed)
  })

  test("Teams 메시지 크기를 넘는 본문은 잘라 보내고 잘렸다고 알린다", async () => {
    const { ports, sent } = fakes()
    const r = await send(ports, "가".repeat(30_000))
    expect(r.code).toBe(SEND_EXIT.ok)
    expect(sent[0].html).toEndWith("…(생략)")
    expect(r.message).toContain("잘")
  })

  test("본문이 원래 …(생략) 으로 끝나도 잘리지 않았으면 잘렸다고 알리지 않는다", async () => {
    const { ports } = fakes()
    const r = await send(ports, "이하 …(생략)")
    expect(r.code).toBe(SEND_EXIT.ok)
    expect(r.message).not.toContain("잘")
  })
})
