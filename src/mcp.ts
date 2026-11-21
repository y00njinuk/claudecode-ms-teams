// 최소 MCP stdio 서버 (newline-delimited JSON-RPC 2.0).
// 채널에 필요한 범위만 구현한다: initialize · ping · tools/list · tools/call · 서버 발신 notification.
import type { Readable, Writable } from "node:stream"
import type { Log } from "./core.js"

export const SUPPORTED_PROTOCOL_VERSIONS = ["2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05"]

export type ToolResult = { text: string; isError?: boolean }

export type Tool = {
  name: string
  description: string
  inputSchema: Record<string, unknown>
  handler: (args: Record<string, unknown>) => Promise<ToolResult>
}

export type McpServerOptions = {
  name: string
  version: string
  instructions: string
  capabilities: Record<string, unknown>
  tools: Tool[]
  onNotification?: (method: string, params: Record<string, unknown>) => void
  input: Readable
  output: Writable
  log: Log
}

type JsonRpcMessage = {
  jsonrpc?: string
  id?: string | number | null
  method?: string
  params?: Record<string, unknown>
}

const METHOD_NOT_FOUND = -32601
const INVALID_PARAMS = -32602
const PARSE_ERROR = -32700

export class McpServer {
  private buffer = ""
  private closeHandlers: Array<() => void> = []

  constructor(private opts: McpServerOptions) {}

  start(): void {
    this.opts.input.setEncoding?.("utf8")
    this.opts.input.on("data", (chunk: string | Buffer) => this.onData(String(chunk)))
    this.opts.input.on("end", () => this.closeHandlers.forEach((h) => h()))
  }

  onClose(handler: () => void): void {
    this.closeHandlers.push(handler)
  }

  notify(method: string, params: Record<string, unknown>): void {
    this.write({ jsonrpc: "2.0", method, params })
  }

  private write(msg: object): void {
    this.opts.output.write(JSON.stringify(msg) + "\n")
  }

  private onData(chunk: string): void {
    this.buffer += chunk
    let nl: number
    while ((nl = this.buffer.indexOf("\n")) >= 0) {
      const line = this.buffer.slice(0, nl).trim()
      this.buffer = this.buffer.slice(nl + 1)
      if (!line) continue
      let parsed: unknown
      try {
        parsed = JSON.parse(line)
      } catch {
        this.write({ jsonrpc: "2.0", id: null, error: { code: PARSE_ERROR, message: "Parse error" } })
        continue
      }
      for (const msg of Array.isArray(parsed) ? parsed : [parsed]) {
        void this.dispatch(msg as JsonRpcMessage)
      }
    }
  }

  private async dispatch(msg: JsonRpcMessage): Promise<void> {
    if (!msg || typeof msg !== "object" || typeof msg.method !== "string") return
    const params = msg.params && typeof msg.params === "object" ? msg.params : {}
    const isRequest = msg.id !== undefined && msg.id !== null
    if (!isRequest) {
      try {
        this.opts.onNotification?.(msg.method, params)
      } catch (e) {
        this.opts.log.error("mcp.notification.crash", { method: msg.method, error: String(e) })
      }
      return
    }
    const id = msg.id
    try {
      const result = await this.handleRequest(msg.method, params)
      this.write({ jsonrpc: "2.0", id, result })
    } catch (e) {
      const err = e as { code?: number; message?: string }
      this.write({
        jsonrpc: "2.0",
        id,
        error: { code: typeof err.code === "number" ? err.code : -32603, message: err.message ?? String(e) },
      })
    }
  }

  private async handleRequest(method: string, params: Record<string, unknown>): Promise<unknown> {
    switch (method) {
      case "initialize": {
        const requested = typeof params.protocolVersion === "string" ? params.protocolVersion : ""
        return {
          protocolVersion: SUPPORTED_PROTOCOL_VERSIONS.includes(requested) ? requested : SUPPORTED_PROTOCOL_VERSIONS[0],
          capabilities: this.opts.capabilities,
          serverInfo: { name: this.opts.name, version: this.opts.version },
          instructions: this.opts.instructions,
        }
      }
      case "ping":
        return {}
      case "tools/list":
        return {
          tools: this.opts.tools.map(({ name, description, inputSchema }) => ({ name, description, inputSchema })),
        }
      case "tools/call": {
        const tool = this.opts.tools.find((t) => t.name === params.name)
        if (!tool) throw { code: INVALID_PARAMS, message: `Unknown tool: ${String(params.name)}` }
        const args = params.arguments && typeof params.arguments === "object" ? (params.arguments as Record<string, unknown>) : {}
        let res: ToolResult
        try {
          res = await tool.handler(args)
        } catch (e) {
          this.opts.log.error("tool.crash", { tool: tool.name, error: String(e) })
          res = { text: `도구 실행 실패: ${String(e)}`, isError: true }
        }
        return { content: [{ type: "text", text: res.text }], ...(res.isError ? { isError: true } : {}) }
      }
      default:
        throw { code: METHOD_NOT_FOUND, message: `Method not found: ${method}` }
    }
  }
}
