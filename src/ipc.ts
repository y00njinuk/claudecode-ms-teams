// hook 프로세스 → 같은 Claude Code 세션의 MCP 서버로 가는 localhost IPC.
// 두 프로세스 모두 같은 `claude` 프로세스의 직계 자식이므로 그 PID 를 레지스트리 키로 쓴다.
// (session_id 는 /clear 시 바뀌지만 MCP 서버 프로세스는 유지되므로 키로 부적합)
import { randomBytes, timingSafeEqual } from "node:crypto"
import * as fs from "node:fs"
import * as http from "node:http"
import * as path from "node:path"
import { dataDir, type Log } from "./core.js"

export const TOKEN_HEADER = "x-claudecode-ms-teams-token"
const MAX_BODY_BYTES = 1024 * 1024
// Stop 요청은 Teams 발신이 끝날 때까지 응답을 잡아둔다. claude 는 Stop hook 이 끝나면 곧 종료하면서
// MCP 서버에 SIGINT → SIGTERM → SIGKILL 을 0.5초 안에 보내므로 (2.1.330 `-p` 실측) 먼저 응답하면 마지막 응답이 사라진다.
// hooks.json 의 Stop timeout 보다 짧아야 한다 (테스트로 강제).
export const STOP_WAIT_MS = 8_000

// 작업이 끝나거나 ms 가 지나면 resolve 한다. 시간이 지나도 작업은 취소하지 않고, 작업의 실패는 밖으로 던지지 않는다.
export function holdAtMost(task: Promise<unknown>, ms: number): Promise<void> {
  return new Promise((resolve) => {
    const timer = setTimeout(resolve, ms)
    timer.unref?.()
    const done = () => {
      clearTimeout(timer)
      resolve()
    }
    task.then(done, done)
  })
}

export type RegistryEntry = {
  serverPid: number
  claudePid: number
  port: number
  token: string
  // 채팅방이 선택된 상태. false 면 hook 은 HTTP 호출 없이 종료한다.
  active: boolean
}

export function runDir(): string {
  return path.join(dataDir(), "run")
}

export function registryPath(claudePid: number): string {
  return path.join(runDir(), `${claudePid}.json`)
}

export function isAlive(pid: number): boolean {
  if (!Number.isInteger(pid) || pid <= 0) return false
  try {
    process.kill(pid, 0)
    return true
  } catch (e) {
    return (e as NodeJS.ErrnoException).code === "EPERM"
  }
}

export function writeRegistry(entry: RegistryEntry): void {
  fs.mkdirSync(runDir(), { recursive: true, mode: 0o700 })
  const target = registryPath(entry.claudePid)
  const tmp = `${target}.${process.pid}.tmp`
  fs.writeFileSync(tmp, JSON.stringify(entry), { mode: 0o600 })
  fs.renameSync(tmp, target)
}

export function readRegistry(claudePid: number): RegistryEntry | null {
  try {
    const e = JSON.parse(fs.readFileSync(registryPath(claudePid), "utf8")) as RegistryEntry
    if (typeof e.port !== "number" || typeof e.token !== "string") return null
    return e
  } catch {
    return null
  }
}

// 다른 서버가 같은 claudePid 로 덮어쓴 파일은 지우지 않는다.
export function removeRegistry(claudePid: number, serverPid: number): void {
  const e = readRegistry(claudePid)
  if (e && e.serverPid !== serverPid) return
  try {
    fs.unlinkSync(registryPath(claudePid))
  } catch {}
}

export function sweepStaleRegistry(): void {
  let names: string[]
  try {
    names = fs.readdirSync(runDir())
  } catch {
    return
  }
  for (const name of names) {
    const m = /^(\d+)\.json$/.exec(name)
    if (!m) continue
    const e = readRegistry(Number(m[1]))
    if (!e || !isAlive(e.serverPid)) {
      try {
        fs.unlinkSync(path.join(runDir(), name))
      } catch {}
    }
  }
}

export function findSession(candidatePids: number[]): RegistryEntry | null {
  for (const pid of candidatePids) {
    if (!Number.isInteger(pid) || pid <= 1) continue
    const e = readRegistry(pid)
    if (e) return e
  }
  return null
}

// void 를 돌려주면 즉시 202 로 끝낸다. Promise 를 돌려주면 resolve 될 때까지 응답을 잡아두고 값을 JSON 으로 보낸다.
// signal 은 응답 전에 hook 쪽 연결이 끊기면 abort 된다 (hook timeout · 세션 종료).
export type IpcHandler = (body: Record<string, unknown>, signal: AbortSignal) => void | Promise<unknown>
export type IpcHandlers = Record<string, IpcHandler>

export type IpcServer = { port: number; token: string; close: () => void }

function tokenMatches(given: string | string[] | undefined, expected: string): boolean {
  if (typeof given !== "string") return false
  const a = Buffer.from(given)
  const b = Buffer.from(expected)
  return a.length === b.length && timingSafeEqual(a, b)
}

export function startIpcServer(handlers: IpcHandlers, log: Log): Promise<IpcServer> {
  const token = randomBytes(24).toString("hex")
  const server = http.createServer((req, res) => {
    const route = (req.url ?? "").replace(/^\//, "")
    const handler = Object.prototype.hasOwnProperty.call(handlers, route) ? handlers[route] : undefined
    if (req.method !== "POST" || !handler) {
      res.writeHead(404).end()
      return
    }
    if (!tokenMatches(req.headers[TOKEN_HEADER], token)) {
      res.writeHead(401).end()
      return
    }
    let size = 0
    const chunks: Buffer[] = []
    req.on("data", (c: Buffer) => {
      size += c.length
      if (size > MAX_BODY_BYTES) {
        res.writeHead(413).end()
        req.destroy()
        return
      }
      chunks.push(c)
    })
    req.on("end", () => {
      if (res.headersSent) return
      let body: Record<string, unknown>
      try {
        body = JSON.parse(Buffer.concat(chunks).toString("utf8"))
      } catch {
        res.writeHead(400).end()
        return
      }
      const abort = new AbortController()
      let pending: void | Promise<unknown> = undefined
      try {
        pending = handler(body, abort.signal)
      } catch (e) {
        log.error("ipc.handler.crash", { route, error: String(e) })
      }
      if (!(pending instanceof Promise)) {
        res.writeHead(202).end()
        return
      }
      res.on("close", () => {
        if (!res.writableEnded) abort.abort()
      })
      pending.then(
        (value) => {
          if (!res.destroyed) res.writeHead(200, { "Content-Type": "application/json" }).end(JSON.stringify(value ?? null))
        },
        (e) => {
          log.error("ipc.handler.crash", { route, error: String(e) })
          if (!res.destroyed) res.writeHead(500).end()
        },
      )
    })
  })
  return new Promise((resolve, reject) => {
    server.once("error", reject)
    server.listen(0, "127.0.0.1", () => {
      server.unref()
      const addr = server.address()
      const port = typeof addr === "object" && addr ? addr.port : 0
      resolve({ port, token, close: () => server.close() })
    })
  })
}
