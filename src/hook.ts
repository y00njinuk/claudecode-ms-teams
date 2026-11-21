#!/usr/bin/env node
// Claude Code command hook → 같은 세션의 MCP 서버로 전달하는 얇은 클라이언트.
//   stop      : Stop 의 last_assistant_message 를 Teams 로 발신. StopFailure 도 이 이벤트로 받아 턴 종료만 알린다
//   permission: PermissionRequest (권한 요청 · AskUserQuestion · plan 승인) 를 Teams 로 알리고,
//               Teams 답장이 올 때까지 기다렸다가 결정을 stdout 으로 돌려준다
//   answered  : PostToolUse(AskUserQuestion) — 터미널에서 먼저 답한 질문의 대기 상태를 정리
//   subagent-stop: SubagentStop — 끝난 sub-agent 가 남긴 대기 상태를 정리
// stdout 은 permission 결정 JSON 에만 쓴다. 그 외에는 아무것도 쓰지 않고 항상 exit 0 으로 끝낸다.
import * as fs from "node:fs"
import * as http from "node:http"
import { pathToFileURL } from "node:url"
import { findSession, TOKEN_HEADER, type RegistryEntry } from "./ipc.js"

const IPC_TIMEOUT_MS = 3000
const ASK_TOOL = "AskUserQuestion"
const EVENTS = new Set(["stop", "permission", "answered", "subagent-stop"])

export function candidatePids(env: NodeJS.ProcessEnv, ppid: number): number[] {
  const fromEnv = Number(env.CLAUDE_PID)
  return [...new Set([fromEnv, ppid].filter((p) => Number.isInteger(p) && p > 1))]
}

export function buildPayload(event: string, input: Record<string, unknown>): Record<string, unknown> | null {
  const base = { session_id: input.session_id, cwd: input.cwd }
  if (event === "stop") {
    // 응답이 비어 있어도 보낸다: 서버가 턴 종료를 알아야 main 의 대기 상태를 정리한다.
    return { ...base, text: typeof input.last_assistant_message === "string" ? input.last_assistant_message : "" }
  }
  if (event === "permission") {
    if (typeof input.tool_name !== "string" || !input.tool_name) return null
    return {
      ...base,
      tool_name: input.tool_name,
      tool_input: input.tool_input,
      permission_suggestions: input.permission_suggestions,
      agent_id: input.agent_id,
      agent_type: input.agent_type,
    }
  }
  if (event === "answered") {
    return input.tool_name === ASK_TOOL ? { ...base, tool_input: input.tool_input } : null
  }
  if (event === "subagent-stop") {
    return typeof input.agent_id === "string" && input.agent_id ? { agent_id: input.agent_id } : null
  }
  return null
}

// 서버가 돌려준 결정을 PermissionRequest hook 출력으로 감싼다. 결정이 없으면 null (터미널 다이얼로그에 맡긴다).
export function permissionOutput(responseBody: string): Record<string, unknown> | null {
  let parsed: { decision?: { behavior?: unknown } | null } | null
  try {
    parsed = JSON.parse(responseBody)
  } catch {
    return null
  }
  const decision = parsed?.decision
  if (!decision || (decision.behavior !== "allow" && decision.behavior !== "deny")) return null
  return { hookSpecificOutput: { hookEventName: "PermissionRequest", decision } }
}

// timeoutMs=0 이면 응답이 올 때까지 기다린다. fetch 는 헤더 대기 제한(5분)이 있어 장시간 대기에 쓸 수 없다.
function post(entry: RegistryEntry, route: string, payload: object, timeoutMs: number): Promise<{ status: number; body: string } | null> {
  return new Promise((resolve) => {
    const data = JSON.stringify(payload)
    const req = http.request(
      {
        host: "127.0.0.1",
        port: entry.port,
        path: `/${route}`,
        method: "POST",
        headers: { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(data), [TOKEN_HEADER]: entry.token },
      },
      (res) => {
        let body = ""
        res.setEncoding("utf8")
        res.on("data", (c) => (body += c))
        res.on("end", () => resolve({ status: res.statusCode ?? 0, body }))
        res.on("error", () => resolve(null))
      },
    )
    req.on("error", () => resolve(null))
    if (timeoutMs > 0) req.setTimeout(timeoutMs, () => req.destroy())
    req.end(data)
  })
}

function readStdin(): Promise<string> {
  return new Promise((resolve) => {
    let data = ""
    process.stdin.setEncoding("utf8")
    process.stdin.on("data", (c) => (data += c))
    process.stdin.on("end", () => resolve(data))
    process.stdin.on("error", () => resolve(data))
  })
}

async function main(): Promise<void> {
  const event = process.argv[2] ?? ""
  if (!EVENTS.has(event)) return
  // 터미널이 먼저 응답해 Claude Code 가 출력을 읽지 않더라도 EPIPE 로 죽지 않게 한다.
  process.stdout.on("error", () => {})
  // stdin 을 먼저 끝까지 읽어 Claude Code 쪽 쓰기가 EPIPE 를 받지 않게 한다.
  const raw = await readStdin()
  const entry = findSession(candidatePids(process.env, process.ppid))
  if (!entry?.active) return
  let input: Record<string, unknown>
  try {
    input = JSON.parse(raw)
  } catch {
    return
  }
  const payload = buildPayload(event, input)
  if (!payload) return
  const res = await post(entry, event, payload, event === "permission" ? 0 : IPC_TIMEOUT_MS)
  if (event !== "permission" || res?.status !== 200) return
  const out = permissionOutput(res.body)
  if (out) await new Promise<void>((done) => process.stdout.write(JSON.stringify(out), () => done()))
}

function invokedDirectly(): boolean {
  try {
    return Boolean(process.argv[1]) && import.meta.url === pathToFileURL(fs.realpathSync(process.argv[1])).href
  } catch {
    return false
  }
}

if (invokedDirectly()) {
  main()
    .catch(() => {})
    .finally(() => process.exit(0))
}
