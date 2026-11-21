import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { spawn } from "node:child_process"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"
import { NOOP_LOG } from "../src/core"
import { buildPayload, candidatePids, permissionOutput } from "../src/hook"
import {
  findSession,
  readRegistry,
  registryPath,
  removeRegistry,
  startIpcServer,
  sweepStaleRegistry,
  TOKEN_HEADER,
  writeRegistry,
} from "../src/ipc"

let tmp: string
const prevXdg = process.env.XDG_CONFIG_HOME

beforeAll(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), "ccmt-ipc-"))
  process.env.XDG_CONFIG_HOME = tmp
})

afterAll(() => {
  if (prevXdg === undefined) delete process.env.XDG_CONFIG_HOME
  else process.env.XDG_CONFIG_HOME = prevXdg
  fs.rmSync(tmp, { recursive: true, force: true })
})

describe("registry", () => {
  test("claude PID 로 쓰고 읽고, 다른 서버가 덮어쓴 파일은 지우지 않는다", () => {
    const entry = { serverPid: process.pid, claudePid: 4242, port: 1234, token: "t", active: true }
    writeRegistry(entry)
    expect(readRegistry(4242)).toEqual(entry)
    expect((fs.statSync(registryPath(4242)).mode & 0o777).toString(8)).toBe("600")
    expect(findSession([0, 1, 9999, 4242])).toEqual(entry)
    removeRegistry(4242, 777)
    expect(readRegistry(4242)).not.toBeNull()
    removeRegistry(4242, process.pid)
    expect(readRegistry(4242)).toBeNull()
  })

  test("죽은 서버의 레지스트리는 sweep 으로 정리한다", () => {
    writeRegistry({ serverPid: 2 ** 22 + 12345, claudePid: 5151, port: 1, token: "t", active: false })
    sweepStaleRegistry()
    expect(readRegistry(5151)).toBeNull()
  })
})

describe("IPC server", () => {
  test("토큰 없는 요청은 401, 모르는 경로는 404, 정상 요청은 202 후 핸들러 호출", async () => {
    const got: unknown[] = []
    const ipc = await startIpcServer({ stop: (b) => void got.push(b) }, NOOP_LOG)
    const url = `http://127.0.0.1:${ipc.port}`
    const post = (p: string, headers: Record<string, string>, body = "{}") =>
      fetch(`${url}/${p}`, { method: "POST", headers: { "Content-Type": "application/json", ...headers }, body })
    expect((await post("stop", {})).status).toBe(401)
    expect((await post("stop", { [TOKEN_HEADER]: "wrong" })).status).toBe(401)
    expect((await post("nope", { [TOKEN_HEADER]: ipc.token })).status).toBe(404)
    expect((await post("stop", { [TOKEN_HEADER]: ipc.token }, "not json")).status).toBe(400)
    expect((await post("stop", { [TOKEN_HEADER]: ipc.token }, JSON.stringify({ text: "hi" }))).status).toBe(202)
    expect(got).toEqual([{ text: "hi" }])
    ipc.close()
  })

  test("Promise 를 돌려주는 핸들러는 resolve 될 때까지 응답을 잡아두고 값을 JSON 으로 보낸다", async () => {
    let release!: (v: unknown) => void
    const ipc = await startIpcServer({ permission: () => new Promise((r) => (release = r)) }, NOOP_LOG)
    const pending = fetch(`http://127.0.0.1:${ipc.port}/permission`, {
      method: "POST",
      headers: { "Content-Type": "application/json", [TOKEN_HEADER]: ipc.token },
      body: "{}",
    })
    const early = await Promise.race([pending.then(() => "responded"), new Promise((r) => setTimeout(() => r("waiting"), 150))])
    expect(early).toBe("waiting")
    release({ decision: { behavior: "allow" } })
    const res = await pending
    expect(res.status).toBe(200)
    expect(await res.json()).toEqual({ decision: { behavior: "allow" } })
    ipc.close()
  })
})

describe("hook", () => {
  test("candidatePids 는 CLAUDE_PID 우선, 중복·무효 값 제거", () => {
    expect(candidatePids({ CLAUDE_PID: "300" }, 300)).toEqual([300])
    expect(candidatePids({ CLAUDE_PID: "300" }, 200)).toEqual([300, 200])
    expect(candidatePids({}, 200)).toEqual([200])
    expect(candidatePids({ CLAUDE_PID: "abc" }, 1)).toEqual([])
  })

  test("buildPayload: stop 은 빈 응답도 넘기고, permission 은 도구 정보를, answered 는 AskUserQuestion 만 넘긴다", () => {
    expect(buildPayload("stop", { session_id: "s", cwd: "/w", last_assistant_message: "완료" })).toEqual({
      session_id: "s",
      cwd: "/w",
      text: "완료",
    })
    // StopFailure 처럼 응답이 없어도 턴 종료는 알린다.
    expect(buildPayload("stop", { session_id: "s" })).toEqual({ session_id: "s", cwd: undefined, text: "" })
    expect(buildPayload("subagent-stop", { session_id: "s", agent_id: "a1", agent_type: "Explore" })).toEqual({ agent_id: "a1" })
    expect(buildPayload("subagent-stop", { session_id: "s" })).toBeNull()
    expect(
      buildPayload("permission", {
        session_id: "s",
        cwd: "/w",
        tool_name: "Bash",
        tool_input: { command: "ls" },
        permission_suggestions: [{ type: "addRules" }],
        agent_id: "a1",
        agent_type: "Explore",
        transcript_path: "/t.jsonl",
      }),
    ).toEqual({
      session_id: "s",
      cwd: "/w",
      tool_name: "Bash",
      tool_input: { command: "ls" },
      permission_suggestions: [{ type: "addRules" }],
      agent_id: "a1",
      agent_type: "Explore",
    })
    expect(buildPayload("permission", { session_id: "s" })).toBeNull()
    expect(buildPayload("answered", { tool_name: "AskUserQuestion", tool_input: { questions: [] } })).toEqual({
      session_id: undefined,
      cwd: undefined,
      tool_input: { questions: [] },
    })
    expect(buildPayload("answered", { tool_name: "Bash", tool_input: {} })).toBeNull()
    expect(buildPayload("other", {})).toBeNull()
  })

  test("permissionOutput: 결정이 있을 때만 PermissionRequest hook 출력을 만든다", () => {
    expect(permissionOutput(JSON.stringify({ decision: { behavior: "allow", updatedInput: { a: 1 } } }))).toEqual({
      hookSpecificOutput: { hookEventName: "PermissionRequest", decision: { behavior: "allow", updatedInput: { a: 1 } } },
    })
    expect(permissionOutput(JSON.stringify({ decision: null }))).toBeNull()
    expect(permissionOutput(JSON.stringify({ decision: { behavior: "ask" } }))).toBeNull()
    expect(permissionOutput("null")).toBeNull()
    expect(permissionOutput("<html>")).toBeNull()
  })

  function runHook(event: string, stdin: string, env: Record<string, string>): Promise<{ code: number; stdout: string }> {
    return new Promise((resolve) => {
      const child = spawn(process.execPath, [path.join(import.meta.dir, "../src/hook.ts"), event], {
        env: { ...process.env, XDG_CONFIG_HOME: tmp, ...env },
      })
      let stdout = ""
      child.stdout.on("data", (d) => (stdout += d))
      child.on("close", (code) => resolve({ code: code ?? -1, stdout }))
      child.stdin.end(stdin)
    })
  }

  test("활성 세션이면 hook 프로세스가 IPC 로 응답을 넘기고 stdout 없이 exit 0", async () => {
    const got: unknown[] = []
    const ipc = await startIpcServer({ stop: (b) => void got.push(b) }, NOOP_LOG)
    writeRegistry({ serverPid: process.pid, claudePid: 6161, port: ipc.port, token: ipc.token, active: true })
    const r = await runHook("stop", JSON.stringify({ session_id: "s1", cwd: "/w/p", last_assistant_message: "끝" }), { CLAUDE_PID: "6161" })
    expect(r).toEqual({ code: 0, stdout: "" })
    expect(got).toEqual([{ session_id: "s1", cwd: "/w/p", text: "끝" }])
    ipc.close()
  })

  test("비활성 세션·레지스트리 없음·깨진 입력이면 아무것도 보내지 않는다", async () => {
    const got: unknown[] = []
    const ipc = await startIpcServer({ stop: (b) => void got.push(b) }, NOOP_LOG)
    writeRegistry({ serverPid: process.pid, claudePid: 7171, port: ipc.port, token: ipc.token, active: false })
    const stdin = JSON.stringify({ last_assistant_message: "끝" })
    expect(await runHook("stop", stdin, { CLAUDE_PID: "7171" })).toEqual({ code: 0, stdout: "" })
    expect(await runHook("stop", stdin, { CLAUDE_PID: "8181" })).toEqual({ code: 0, stdout: "" })
    writeRegistry({ serverPid: process.pid, claudePid: 7171, port: ipc.port, token: ipc.token, active: true })
    expect(await runHook("stop", "{not json", { CLAUDE_PID: "7171" })).toEqual({ code: 0, stdout: "" })
    expect(got).toEqual([])
    ipc.close()
  })

  test("permission hook 은 서버의 결정을 기다렸다가 PermissionRequest 출력으로 stdout 에 쓴다", async () => {
    const got: unknown[] = []
    const decision = { behavior: "allow", updatedInput: { questions: [], answers: { "어디에?": "dev" } } }
    const ipc = await startIpcServer(
      {
        permission: (b) => {
          got.push(b)
          return new Promise((r) => setTimeout(() => r({ decision }), 200))
        },
      },
      NOOP_LOG,
    )
    writeRegistry({ serverPid: process.pid, claudePid: 9191, port: ipc.port, token: ipc.token, active: true })
    const stdin = JSON.stringify({ session_id: "s1", cwd: "/w/p", tool_name: "AskUserQuestion", tool_input: { questions: [] } })
    const r = await runHook("permission", stdin, { CLAUDE_PID: "9191" })
    expect(r.code).toBe(0)
    expect(JSON.parse(r.stdout)).toEqual({ hookSpecificOutput: { hookEventName: "PermissionRequest", decision } })
    expect(got).toEqual([{ session_id: "s1", cwd: "/w/p", tool_name: "AskUserQuestion", tool_input: { questions: [] } }])
    ipc.close()
  })

  test("permission hook 은 결정이 없거나 세션이 비활성이면 stdout 없이 끝난다 (터미널 다이얼로그 유지)", async () => {
    const ipc = await startIpcServer({ permission: async () => ({ decision: null }) }, NOOP_LOG)
    const stdin = JSON.stringify({ tool_name: "Bash", tool_input: { command: "ls" } })
    writeRegistry({ serverPid: process.pid, claudePid: 9292, port: ipc.port, token: ipc.token, active: true })
    expect(await runHook("permission", stdin, { CLAUDE_PID: "9292" })).toEqual({ code: 0, stdout: "" })
    writeRegistry({ serverPid: process.pid, claudePid: 9292, port: ipc.port, token: ipc.token, active: false })
    expect(await runHook("permission", stdin, { CLAUDE_PID: "9292" })).toEqual({ code: 0, stdout: "" })
    ipc.close()
  })
})
