import { describe, expect, test } from "bun:test"
import { PassThrough } from "node:stream"
import { NOOP_LOG } from "../src/core"
import { McpServer, type Tool } from "../src/mcp"

function harness(tools: Tool[] = [], onNotification?: (m: string, p: Record<string, unknown>) => void) {
  const input = new PassThrough()
  const output = new PassThrough()
  const server = new McpServer({
    name: "microsoft-teams",
    version: "9.9.9",
    instructions: "지시문",
    capabilities: { tools: {}, experimental: { "claude/channel": {}, "claude/channel/permission": {} } },
    tools,
    input,
    output,
    log: NOOP_LOG,
    onNotification,
  })
  server.start()
  const lines: any[] = []
  let pending = ""
  output.on("data", (d) => {
    pending += String(d)
    let nl
    while ((nl = pending.indexOf("\n")) >= 0) {
      lines.push(JSON.parse(pending.slice(0, nl)))
      pending = pending.slice(nl + 1)
    }
  })
  const send = (m: unknown) => input.write(JSON.stringify(m) + "\n")
  const waitFor = async (pred: (l: any[]) => boolean) => {
    for (let i = 0; i < 100 && !pred(lines); i++) await new Promise((r) => setTimeout(r, 5))
    return lines
  }
  return { server, input, send, lines, waitFor }
}

const echoTool: Tool = {
  name: "echo",
  description: "echo",
  inputSchema: { type: "object", properties: { v: { type: "string" } } },
  handler: async (args) => ({ text: `echo:${String(args.v)}` }),
}

describe("McpServer", () => {
  test("initialize 는 지원 버전을 그대로 돌려주고 채널 capability 와 instructions 를 알린다", async () => {
    const h = harness()
    h.send({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-06-18", capabilities: {} } })
    const [res] = await h.waitFor((l) => l.length >= 1)
    expect(res.id).toBe(1)
    expect(res.result.protocolVersion).toBe("2025-06-18")
    expect(res.result.capabilities.experimental["claude/channel"]).toEqual({})
    expect(res.result.capabilities.experimental["claude/channel/permission"]).toEqual({})
    expect(res.result.instructions).toBe("지시문")
    expect(res.result.serverInfo).toEqual({ name: "microsoft-teams", version: "9.9.9" })
  })

  test("모르는 프로토콜 버전이면 최신 지원 버전으로 응답한다", async () => {
    const h = harness()
    h.send({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2099-01-01" } })
    const [res] = await h.waitFor((l) => l.length >= 1)
    expect(res.result.protocolVersion).toBe("2025-11-25")
  })

  test("tools/list · tools/call · 도구 예외는 isError 로 감싼다", async () => {
    const boom: Tool = { ...echoTool, name: "boom", handler: async () => Promise.reject(new Error("x")) }
    const h = harness([echoTool, boom])
    h.send({ jsonrpc: "2.0", id: 1, method: "tools/list" })
    h.send({ jsonrpc: "2.0", id: 2, method: "tools/call", params: { name: "echo", arguments: { v: "hi" } } })
    h.send({ jsonrpc: "2.0", id: 3, method: "tools/call", params: { name: "boom", arguments: {} } })
    const lines = await h.waitFor((l) => l.length >= 3)
    const byId = Object.fromEntries(lines.map((l) => [l.id, l]))
    expect(byId[1].result.tools.map((t: { name: string }) => t.name)).toEqual(["echo", "boom"])
    expect(byId[1].result.tools[0].handler).toBeUndefined()
    expect(byId[2].result).toEqual({ content: [{ type: "text", text: "echo:hi" }] })
    expect(byId[3].result.isError).toBe(true)
  })

  test("모르는 도구는 -32602, 모르는 메서드(server/discover 등)는 -32601", async () => {
    const h = harness([echoTool])
    h.send({ jsonrpc: "2.0", id: 1, method: "tools/call", params: { name: "nope" } })
    h.send({ jsonrpc: "2.0", id: 2, method: "server/discover", params: {} })
    const lines = await h.waitFor((l) => l.length >= 2)
    const byId = Object.fromEntries(lines.map((l) => [l.id, l]))
    expect(byId[1].error.code).toBe(-32602)
    expect(byId[2].error.code).toBe(-32601)
  })

  test("notification 은 응답 없이 핸들러로 전달된다", async () => {
    const got: Array<[string, unknown]> = []
    const h = harness([], (m, p) => got.push([m, p]))
    h.send({ jsonrpc: "2.0", method: "notifications/initialized" })
    h.send({
      jsonrpc: "2.0",
      method: "notifications/claude/channel/permission_request",
      params: { request_id: "abcde", tool_name: "Bash" },
    })
    h.send({ jsonrpc: "2.0", id: 9, method: "ping" })
    const lines = await h.waitFor((l) => l.length >= 1)
    expect(lines).toEqual([{ jsonrpc: "2.0", id: 9, result: {} }])
    expect(got.map(([m]) => m)).toEqual(["notifications/initialized", "notifications/claude/channel/permission_request"])
  })

  test("청크로 쪼개진 줄과 batch 배열을 처리한다", async () => {
    const h = harness()
    const line = JSON.stringify([
      { jsonrpc: "2.0", id: 1, method: "ping" },
      { jsonrpc: "2.0", id: 2, method: "ping" },
    ])
    h.input.write(line.slice(0, 10))
    h.input.write(line.slice(10) + "\n")
    const lines = await h.waitFor((l) => l.length >= 2)
    expect(lines.map((l) => l.id).sort()).toEqual([1, 2])
  })

  test("notify 는 id 없는 JSON-RPC notification 을 쓴다", async () => {
    const h = harness()
    h.server.notify("notifications/claude/channel", { content: "hi", meta: { sender: "나" } })
    const [n] = await h.waitFor((l) => l.length >= 1)
    expect(n).toEqual({ jsonrpc: "2.0", method: "notifications/claude/channel", params: { content: "hi", meta: { sender: "나" } } })
  })

  test("stdin 이 닫히면 onClose 핸들러를 호출한다", async () => {
    const h = harness()
    let closed = false
    h.server.onClose(() => (closed = true))
    h.input.end()
    for (let i = 0; i < 50 && !closed; i++) await new Promise((r) => setTimeout(r, 5))
    expect(closed).toBe(true)
  })
})
