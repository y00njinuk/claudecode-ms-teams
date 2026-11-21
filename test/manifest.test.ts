// 플러그인 매니페스트·hook·스킬이 코드 상수와 어긋나지 않는지 검사한다.
import { describe, expect, test } from "bun:test"
import * as fs from "node:fs"
import * as path from "node:path"
import { MicrosoftTeamsChannel, REPLY_TTL_MS } from "../src/channel"
import { defaultConfig, MARKETPLACE_NAME, NOOP_LOG, PLUGIN_NAME } from "../src/core"
import { ipcTimeoutMs } from "../src/hook"
import { STOP_WAIT_MS } from "../src/ipc"

const ROOT = path.join(import.meta.dir, "..")
const readJson = (p: string) => JSON.parse(fs.readFileSync(path.join(ROOT, p), "utf8"))

const pkg = readJson("package.json")
const plugin = readJson(".claude-plugin/plugin.json")
const market = readJson(".claude-plugin/marketplace.json")
const hooks = readJson("hooks/hooks.json")

describe("manifest", () => {
  test("plugin.json 버전은 package.json 버전과 같다", () => {
    expect(plugin.version).toBe(pkg.version)
  })

  test("플러그인·마켓플레이스 이름이 코드 상수와 같다", () => {
    expect(plugin.name).toBe(PLUGIN_NAME)
    expect(market.name).toBe(MARKETPLACE_NAME)
    expect(market.plugins).toEqual([expect.objectContaining({ name: PLUGIN_NAME, source: "./" })])
  })

  test("MCP 서버 이름이 채널 source 와 같고 dist/server.js 를 실행한다", () => {
    expect(Object.keys(plugin.mcpServers)).toEqual([PLUGIN_NAME])
    expect(plugin.mcpServers[PLUGIN_NAME]).toEqual({ command: "node", args: ["${CLAUDE_PLUGIN_ROOT}/dist/server.js"] })
  })

  test("hook 은 Stop · PermissionRequest · PostToolUse(AskUserQuestion) 에서 dist/hook.js 를 부른다", () => {
    const command = (event: string) => `node "\${CLAUDE_PLUGIN_ROOT}/dist/hook.js" ${event}`
    expect(hooks.hooks.Stop[0].hooks[0].command).toBe(command("stop"))
    expect(hooks.hooks.PermissionRequest[0].matcher).toBe("*")
    expect(hooks.hooks.PermissionRequest[0].hooks[0].command).toBe(command("permission"))
    expect(hooks.hooks.PostToolUse[0].matcher).toBe("AskUserQuestion")
    expect(hooks.hooks.PostToolUse[0].hooks[0].command).toBe(command("answered"))
    // StopFailure 는 응답 없이 턴 종료만 알리므로 Stop 과 같은 명령을 쓴다.
    expect(hooks.hooks.StopFailure[0].hooks[0].command).toBe(command("stop"))
    expect(hooks.hooks.SubagentStop[0].hooks[0].command).toBe(command("subagent-stop"))
    expect(Object.keys(hooks.hooks).sort()).toEqual(["PermissionRequest", "PostToolUse", "Stop", "StopFailure", "SubagentStop"])
  })

  test("PermissionRequest hook 은 답장 코드가 유효한 동안 기다린다", () => {
    expect(hooks.hooks.PermissionRequest[0].hooks[0].timeout * 1000).toBe(REPLY_TTL_MS)
  })

  test("Stop hook 은 서버가 발신을 기다리는 시간보다 오래, hook timeout 보다 짧게 응답을 기다린다", () => {
    expect(ipcTimeoutMs("stop")).toBeGreaterThan(STOP_WAIT_MS)
    for (const event of ["Stop", "StopFailure"]) {
      expect(ipcTimeoutMs("stop")).toBeLessThan(hooks.hooks[event][0].hooks[0].timeout * 1000)
    }
  })

  test("npm tarball 에 플러그인 구성 요소가 포함된다", () => {
    for (const f of [".claude-plugin", "hooks", "skills", "dist", "postinstall.mjs"]) expect(pkg.files).toContain(f)
    expect(pkg.bin["claudecode-ms-teams"]).toBe("dist/cli.js")
  })

  test("connect 스킬이 언급하는 도구가 실제로 존재한다", () => {
    const skill = fs.readFileSync(path.join(ROOT, "skills/connect/SKILL.md"), "utf8")
    const ch = new MicrosoftTeamsChannel({
      cfg: defaultConfig(),
      log: NOOP_LOG,
      auth: {} as never,
      graph: {} as never,
      configPath: "",
      push: () => {},
    })
    const names = ch.tools().map((t) => t.name)
    for (const name of ["connect", "select_chat", "status"]) {
      expect(names).toContain(name)
      expect(skill).toContain(`\`${name}\``)
    }
  })
})
