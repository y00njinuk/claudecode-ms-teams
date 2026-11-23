// 배포본은 Bun 이 아니라 Node 로 실행된다. 런타임마다 동작이 다른 부분은 빌드 결과를 Node 로 돌려 확인한다.
import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { spawn, spawnSync } from "node:child_process"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"
import { pathToFileURL } from "node:url"

const ROOT = path.join(import.meta.dir, "..")
let outDir: string

beforeAll(() => {
  outDir = fs.mkdtempSync(path.join(os.tmpdir(), "ccmt-node-"))
  const tsc = spawnSync(path.join(ROOT, "node_modules/.bin/tsc"), ["-p", "tsconfig.build.json", "--outDir", outDir], {
    cwd: ROOT,
    encoding: "utf8",
  })
  if (tsc.status !== 0) throw new Error(`tsc 실패: ${tsc.stdout}${tsc.stderr}`)
})

afterAll(() => {
  fs.rmSync(outDir, { recursive: true, force: true })
})

describe("Node 런타임", () => {
  test("응답을 기다리던 hook 연결이 끊기면 IPC 핸들러의 signal 이 abort 된다", () => {
    const r = spawnSync("node", [path.join(import.meta.dir, "fixtures/ipc-abort.mjs"), outDir], { encoding: "utf8" })
    expect(r.status).toBe(0)
    expect(JSON.parse(r.stdout)).toEqual({ abortedBeforeDisconnect: false, abortedAfterDisconnect: true })
  })

  test("send CLI 는 사용법 오류면 2, 로그인 전이면 3 으로 끝난다", () => {
    const xdg = fs.mkdtempSync(path.join(os.tmpdir(), "ccmt-send-"))
    try {
      const send = (args: string[], input: string) =>
        spawnSync("node", [path.join(outDir, "cli.js"), "send", ...args], {
          input,
          encoding: "utf8",
          env: { ...process.env, XDG_CONFIG_HOME: xdg },
        })
      expect(send([], "보고").status).toBe(2)
      expect(send(["--chat", "19:abc@thread.v2"], "").status).toBe(2)
      const noLogin = send(["--chat", "19:abc@thread.v2"], "보고")
      expect(noLogin.status).toBe(3)
      expect(noLogin.stdout).toBe("")
      expect(noLogin.stderr).toContain("azure")
    } finally {
      fs.rmSync(xdg, { recursive: true, force: true })
    }
  })

  test("send CLI 는 Graph 429 의 재시도 대기 중에 종료하지 않고 결과를 종료 코드로 돌려준다", () => {
    const xdg = fs.mkdtempSync(path.join(os.tmpdir(), "ccmt-send429-"))
    try {
      const dir = path.join(xdg, "claudecode-ms-teams")
      fs.mkdirSync(dir, { recursive: true })
      fs.writeFileSync(path.join(dir, "claudecode-ms-teams.json"), JSON.stringify({ azure: { tenantId: "t", clientId: "c", scopes: "s" } }))
      const tokens = { access_token: "at", refresh_token: "rt", expires_at: Date.now() + 3_600_000 }
      fs.writeFileSync(path.join(dir, "tokens.json"), JSON.stringify(tokens), { mode: 0o600 })
      const report = path.join(xdg, "report.md")
      fs.writeFileSync(report, "## 야간 작업\n- 완료")
      const r = spawnSync(
        "node",
        ["--import", pathToFileURL(path.join(import.meta.dir, "fixtures/fetch-429.mjs")).href, path.join(outDir, "cli.js"), "send", "--chat", "19:abc@thread.v2", "--file", report],
        { encoding: "utf8", env: { ...process.env, XDG_CONFIG_HOME: xdg } },
      )
      expect(r.status).toBe(0)
      expect(r.stdout).toContain("1700000000042")
    } finally {
      fs.rmSync(xdg, { recursive: true, force: true })
    }
  })

  test("서버의 stop 라우트는 응답을 잡아두고(200), answered 는 즉시 202 로 끝낸다", async () => {
    const xdg = fs.mkdtempSync(path.join(os.tmpdir(), "ccmt-server-"))
    const server = spawn("node", [path.join(outDir, "server.js")], { env: { ...process.env, XDG_CONFIG_HOME: xdg }, stdio: ["pipe", "ignore", "ignore"] })
    try {
      // 서버는 부모(이 테스트 프로세스) PID 로 레지스트리를 쓴다.
      const registry = path.join(xdg, "claudecode-ms-teams", "run", `${process.pid}.json`)
      const deadline = Date.now() + 5_000
      while (!fs.existsSync(registry) && Date.now() < deadline) await new Promise((r) => setTimeout(r, 50))
      const { port, token } = JSON.parse(fs.readFileSync(registry, "utf8"))
      const post = (route: string, body: object) =>
        fetch(`http://127.0.0.1:${port}/${route}`, {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-claudecode-ms-teams-token": token },
          body: JSON.stringify(body),
        })
      expect((await post("stop", { session_id: "s", cwd: "/w", text: "" })).status).toBe(200)
      expect((await post("answered", { tool_input: { questions: [] } })).status).toBe(202)
    } finally {
      server.stdin.end()
      server.kill()
      fs.rmSync(xdg, { recursive: true, force: true })
    }
  })

  test("서버는 시작 시 azure.scopes 의 Tasks 권한으로 To Do 도구 · 지침 노출을 정한다 (Graph 요청 없음)", async () => {
    const listTools = async (scopes: string) => {
      const xdg = fs.mkdtempSync(path.join(os.tmpdir(), "ccmt-todo-"))
      const dir = path.join(xdg, "claudecode-ms-teams")
      fs.mkdirSync(dir, { recursive: true })
      fs.writeFileSync(path.join(dir, "claudecode-ms-teams.json"), JSON.stringify({ azure: { tenantId: "t", clientId: "c", scopes } }))
      const server = spawn("node", [path.join(outDir, "server.js")], { env: { ...process.env, XDG_CONFIG_HOME: xdg }, stdio: ["pipe", "pipe", "ignore"] })
      try {
        let out = ""
        server.stdout.on("data", (d) => (out += String(d)))
        server.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-06-18" } }) + "\n")
        server.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: 2, method: "tools/list" }) + "\n")
        const deadline = Date.now() + 5_000
        while (out.split("\n").filter(Boolean).length < 2 && Date.now() < deadline) await new Promise((r) => setTimeout(r, 20))
        const [init, list] = out.split("\n").filter(Boolean).map((l) => JSON.parse(l))
        return { instructions: init.result.instructions as string, names: (list.result.tools as Array<{ name: string }>).map((t) => t.name) }
      } finally {
        server.stdin.end()
        server.kill()
        fs.rmSync(xdg, { recursive: true, force: true })
      }
    }
    const teams = ["connect", "select_chat", "status", "disconnect"]
    const none = await listTools("Chat.Read offline_access")
    expect(none.names).toEqual(teams)
    expect(none.instructions).not.toContain("To Do")
    const read = await listTools("Chat.Read offline_access Tasks.Read")
    expect(read.names).toEqual([...teams, "todo_list_lists", "todo_list_tasks", "todo_get_task"])
    const rw = await listTools("Chat.Read offline_access Tasks.ReadWrite")
    expect(rw.names).toHaveLength(teams.length + 7)
    expect(rw.instructions).toContain("todo_list_lists")
  })
})
