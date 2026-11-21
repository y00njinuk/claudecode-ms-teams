import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"
import { CONFIG_FILE_NAME, DATA_DIR_NAME, defaultConfig, MARKETPLACE_NAME, PLUGIN_ID } from "../src/core"
import { doctor, install, uninstall, type RunResult } from "../src/install"

let tmp: string
let pkgDir: string
const prevXdg = process.env.XDG_CONFIG_HOME

beforeAll(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), "ccmt-install-"))
  pkgDir = path.join(tmp, "pkg")
  fs.mkdirSync(pkgDir)
  process.env.XDG_CONFIG_HOME = path.join(tmp, "xdg")
})

afterAll(() => {
  if (prevXdg === undefined) delete process.env.XDG_CONFIG_HOME
  else process.env.XDG_CONFIG_HOME = prevXdg
  fs.rmSync(tmp, { recursive: true, force: true })
})

type State = {
  markets: Array<{ name: string; path: string }>
  plugins: Array<{ id: string; enabled: boolean; version?: string }>
  claudeMissing?: boolean
  updateFails?: boolean
}

function fakeClaude(state: State) {
  const calls: string[] = []
  const ok = (stdout = ""): RunResult => ({ status: 0, stdout, stderr: "" })
  const run = (args: string[]): RunResult => {
    calls.push(args.join(" "))
    if (state.claudeMissing) return { status: -1, stdout: "", stderr: "ENOENT" }
    const cmd = args.join(" ")
    if (cmd === "--version") return ok("2.1.319 (Claude Code)")
    if (cmd === "plugin marketplace list --json") return ok(JSON.stringify(state.markets.map((m) => ({ ...m, source: "directory" }))))
    if (cmd === "plugin list --json") return ok(JSON.stringify(state.plugins))
    if (args[0] === "plugin" && args[1] === "marketplace" && args[2] === "add") {
      state.markets.push({ name: MARKETPLACE_NAME, path: args[3] })
      return ok()
    }
    if (cmd === `plugin marketplace remove ${MARKETPLACE_NAME}`) {
      state.markets = state.markets.filter((m) => m.name !== MARKETPLACE_NAME)
      return ok()
    }
    if (cmd.startsWith(`plugin install ${PLUGIN_ID}`)) {
      state.plugins.push({ id: PLUGIN_ID, enabled: true })
      return ok()
    }
    if (cmd === `plugin enable ${PLUGIN_ID}`) {
      state.plugins = state.plugins.map((p) => ({ ...p, enabled: true }))
      return ok()
    }
    if (cmd === `plugin marketplace update ${MARKETPLACE_NAME}`) return ok()
    if (cmd === `plugin update ${PLUGIN_ID}`) {
      if (state.updateFails) return { status: 1, stdout: "", stderr: "network" }
      state.plugins = state.plugins.map((p) => ({ ...p, version: "9.9.9" }))
      return ok()
    }
    if (cmd === `plugin uninstall ${PLUGIN_ID}`) {
      state.plugins = []
      return ok()
    }
    return { status: 1, stdout: "", stderr: `unexpected: ${cmd}` }
  }
  return { run, calls }
}

const mutating = (calls: string[]) => calls.filter((c) => !c.endsWith("--json") && c !== "--version")

describe("install", () => {
  test("최초 설치: 설정 파일 생성 + 마켓플레이스 등록 + 플러그인 설치", () => {
    const { run, calls } = fakeClaude({ markets: [], plugins: [] })
    const r = install({ packageDir: pkgDir, run })
    expect(mutating(calls)).toEqual([`plugin marketplace add ${pkgDir} --scope user`, `plugin install ${PLUGIN_ID} --scope user`])
    expect(r.configCreated).toBe(true)
    const cfgPath = path.join(tmp, "xdg", DATA_DIR_NAME, CONFIG_FILE_NAME)
    expect(r.configPath).toBe(cfgPath)
    expect(Object.keys(JSON.parse(fs.readFileSync(cfgPath, "utf8")))).toEqual(Object.keys(defaultConfig()))
    expect(r.launchCommand).toBe(`claude --dangerously-load-development-channels plugin:${PLUGIN_ID}`)
  })

  test("이미 같은 경로로 설치돼 있으면 아무것도 바꾸지 않고 설정 파일도 덮어쓰지 않는다", () => {
    const cfgPath = path.join(tmp, "xdg", DATA_DIR_NAME, CONFIG_FILE_NAME)
    fs.writeFileSync(cfgPath, '{"azure":{"tenantId":"keep"}}')
    const { run, calls } = fakeClaude({ markets: [{ name: MARKETPLACE_NAME, path: pkgDir }], plugins: [{ id: PLUGIN_ID, enabled: true }] })
    const r = install({ packageDir: pkgDir, run })
    expect(mutating(calls)).toEqual([])
    expect(r.configCreated).toBe(false)
    expect(fs.readFileSync(cfgPath, "utf8")).toBe('{"azure":{"tenantId":"keep"}}')
  })

  test("패키지 경로가 바뀌면 마켓플레이스를 다시 등록하고, 비활성 플러그인은 활성화한다", () => {
    const { run, calls } = fakeClaude({
      markets: [{ name: MARKETPLACE_NAME, path: path.join(tmp, "old") }],
      plugins: [{ id: PLUGIN_ID, enabled: false }],
    })
    install({ packageDir: pkgDir, run })
    expect(mutating(calls)).toEqual([
      `plugin marketplace remove ${MARKETPLACE_NAME}`,
      `plugin marketplace add ${pkgDir} --scope user`,
      `plugin enable ${PLUGIN_ID}`,
    ])
  })

  test("npm 업그레이드 후에는 Claude Code 에 기록된 플러그인 버전을 새 버전으로 갱신한다", () => {
    fs.writeFileSync(path.join(pkgDir, "package.json"), JSON.stringify({ version: "9.9.9" }))
    const state: State = { markets: [{ name: MARKETPLACE_NAME, path: pkgDir }], plugins: [{ id: PLUGIN_ID, enabled: true, version: "0.1.0" }] }
    const first = fakeClaude(state)
    const r = install({ packageDir: pkgDir, run: first.run })
    expect(mutating(first.calls)).toEqual([`plugin marketplace update ${MARKETPLACE_NAME}`, `plugin update ${PLUGIN_ID}`])
    expect(r.steps).toEqual(["플러그인 갱신: 0.1.0 → 9.9.9"])
    const second = fakeClaude(state)
    install({ packageDir: pkgDir, run: second.run })
    expect(mutating(second.calls)).toEqual([])
  })

  test("버전 기록 갱신이 실패해도 설치는 실패로 처리하지 않는다", () => {
    const { run } = fakeClaude({
      markets: [{ name: MARKETPLACE_NAME, path: pkgDir }],
      plugins: [{ id: PLUGIN_ID, enabled: true, version: "0.1.0" }],
      updateFails: true,
    })
    const r = install({ packageDir: pkgDir, run })
    expect(r.steps).toEqual(["플러그인 버전 기록 갱신 실패 (동작에는 영향 없음): network"])
  })

  test("claude CLI 가 없으면 조치 가능한 오류를 던진다", () => {
    const { run } = fakeClaude({ markets: [], plugins: [], claudeMissing: true })
    expect(() => install({ packageDir: pkgDir, run })).toThrow("claude CLI")
  })
})

describe("uninstall / doctor", () => {
  test("uninstall 은 플러그인과 마켓플레이스를 제거하고, 기본값으로는 설정을 남긴다", () => {
    const { run, calls } = fakeClaude({ markets: [{ name: MARKETPLACE_NAME, path: pkgDir }], plugins: [{ id: PLUGIN_ID, enabled: true }] })
    const steps = uninstall({ run })
    expect(mutating(calls)).toEqual([`plugin uninstall ${PLUGIN_ID}`, `plugin marketplace remove ${MARKETPLACE_NAME}`])
    expect(steps).toHaveLength(2)
    expect(fs.existsSync(path.join(tmp, "xdg", DATA_DIR_NAME))).toBe(true)
  })

  test("doctor 는 azure 미설정과 미설치를 실패로 보고한다", () => {
    const { run } = fakeClaude({ markets: [], plugins: [] })
    const checks = doctor({ packageDir: pkgDir, run })
    const byLabel = Object.fromEntries(checks.map((c) => [c.label, c.ok]))
    expect(byLabel["설정 파일"]).toBe(false)
    expect(byLabel["마켓플레이스"]).toBe(false)
    expect(byLabel["플러그인"]).toBe(false)
    expect(byLabel["claude CLI"]).toBe(true)
  })

  test("doctor 는 기록된 플러그인 버전이 패키지 버전과 다르면 갱신 방법을 안내한다", () => {
    fs.writeFileSync(path.join(pkgDir, "package.json"), JSON.stringify({ version: "9.9.9" }))
    const { run } = fakeClaude({ markets: [{ name: MARKETPLACE_NAME, path: pkgDir }], plugins: [{ id: PLUGIN_ID, enabled: true, version: "0.1.0" }] })
    const plugin = doctor({ packageDir: pkgDir, run }).find((c) => c.label === "플러그인")
    expect(plugin?.ok).toBe(false)
    expect(plugin?.detail).toContain("0.1.0 ≠ 패키지 9.9.9")
  })

  test("uninstall --purge 는 데이터 디렉토리와 외부 경로의 토큰 파일만 지운다", () => {
    const outside = path.join(tmp, "elsewhere")
    fs.mkdirSync(outside)
    const tokensFile = path.join(outside, "tokens.json")
    const neighbor = path.join(outside, "keep.txt")
    fs.writeFileSync(tokensFile, "{}")
    fs.writeFileSync(neighbor, "x")
    fs.writeFileSync(path.join(tmp, "xdg", DATA_DIR_NAME, CONFIG_FILE_NAME), JSON.stringify({ tokens: { file: tokensFile } }))
    const { run } = fakeClaude({ markets: [], plugins: [] })
    uninstall({ run, purge: true })
    expect(fs.existsSync(path.join(tmp, "xdg", DATA_DIR_NAME))).toBe(false)
    expect(fs.existsSync(tokensFile)).toBe(false)
    expect(fs.existsSync(neighbor)).toBe(true)
    expect(fs.existsSync(pkgDir)).toBe(true)
  })
})
