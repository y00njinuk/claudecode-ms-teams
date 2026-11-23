// 설치/제거/진단. npm 글로벌 설치 경로를 Claude Code 로컬 마켓플레이스로 등록하고 플러그인을 설치한다.
// 로컬 경로 마켓플레이스의 상대 경로 플러그인은 제자리에서 로드되므로 (`readFromFolder`) npm 업그레이드가
// 다음 세션부터 바로 반영된다. 다만 Claude Code 가 기록해 둔 설치 버전은 그대로 남으므로 install 이 함께 갱신한다.
import { spawnSync } from "node:child_process"
import * as fs from "node:fs"
import * as path from "node:path"
import {
  Auth,
  azureConfigured,
  CHANNEL_LAUNCH_COMMAND,
  dataDir,
  ensureConfigFile,
  loadConfig,
  MARKETPLACE_NAME,
  NOOP_LOG,
  PLUGIN_ID,
  TODO_SCOPE_RW,
  todoAccess,
} from "./core.js"
import { readRegistry, runDir, isAlive } from "./ipc.js"
import { describeTodoAccess, describeTodoToken, todoTimeZone } from "./todo.js"

export type RunResult = { status: number; stdout: string; stderr: string }
export type Runner = (args: string[]) => RunResult

export const defaultRunner: Runner = (args) => {
  const r = spawnSync("claude", args, { encoding: "utf8" })
  if (r.error) return { status: -1, stdout: "", stderr: String(r.error.message) }
  return { status: r.status ?? -1, stdout: r.stdout ?? "", stderr: r.stderr ?? "" }
}

type MarketplaceRow = { name?: string; source?: string; path?: string; installLocation?: string }
type PluginRow = { id?: string; enabled?: boolean; installPath?: string; version?: string }

function packageVersion(packageDir: string): string {
  try {
    return String(JSON.parse(fs.readFileSync(path.join(packageDir, "package.json"), "utf8")).version ?? "")
  } catch {
    return ""
  }
}

function parseJsonArray<T>(text: string): T[] | null {
  try {
    const v = JSON.parse(text)
    return Array.isArray(v) ? (v as T[]) : null
  } catch {
    return null
  }
}

function samePath(a: string | undefined, b: string): boolean {
  if (!a) return false
  try {
    return fs.realpathSync(a) === fs.realpathSync(b)
  } catch {
    return path.resolve(a) === path.resolve(b)
  }
}

function marketplaceLocation(m: MarketplaceRow): string | undefined {
  return m.path ?? m.installLocation
}

function fail(step: string, r: RunResult): Error {
  return new Error(`${step} 실패 (exit ${r.status}): ${(r.stderr || r.stdout).trim()}`)
}

export type InstallResult = {
  configPath: string
  configCreated: boolean
  steps: string[]
  launchCommand: string
}

export function install({ packageDir, run = defaultRunner }: { packageDir: string; run?: Runner }): InstallResult {
  const cfg = ensureConfigFile()
  const steps: string[] = []
  if (run(["--version"]).status !== 0) {
    throw new Error("claude CLI 를 찾을 수 없습니다. Claude Code 설치 후 `claudecode-ms-teams install` 을 다시 실행하세요.")
  }

  const markets = parseJsonArray<MarketplaceRow>(run(["plugin", "marketplace", "list", "--json"]).stdout) ?? []
  const existing = markets.find((m) => m.name === MARKETPLACE_NAME)
  if (existing && !samePath(marketplaceLocation(existing), packageDir)) {
    const r = run(["plugin", "marketplace", "remove", MARKETPLACE_NAME])
    if (r.status !== 0) throw fail("기존 마켓플레이스 제거", r)
    steps.push(`마켓플레이스 경로 변경: ${marketplaceLocation(existing)} → ${packageDir}`)
  }
  if (!existing || !samePath(marketplaceLocation(existing), packageDir)) {
    const r = run(["plugin", "marketplace", "add", packageDir, "--scope", "user"])
    if (r.status !== 0) throw fail("마켓플레이스 등록", r)
    steps.push(`마켓플레이스 등록: ${MARKETPLACE_NAME} → ${packageDir}`)
  }

  const plugins = parseJsonArray<PluginRow>(run(["plugin", "list", "--json"]).stdout) ?? []
  const plugin = plugins.find((p) => p.id === PLUGIN_ID)
  if (!plugin) {
    const r = run(["plugin", "install", PLUGIN_ID, "--scope", "user"])
    if (r.status !== 0) throw fail("플러그인 설치", r)
    steps.push(`플러그인 설치: ${PLUGIN_ID}`)
  } else {
    if (plugin.enabled === false) {
      const r = run(["plugin", "enable", PLUGIN_ID])
      if (r.status !== 0) throw fail("플러그인 활성화", r)
      steps.push(`플러그인 활성화: ${PLUGIN_ID}`)
    }
    const version = packageVersion(packageDir)
    if (version && plugin.version && plugin.version !== version) {
      // 실패해도 플러그인은 패키지 경로에서 새 버전으로 로드된다. 기록만 옛 버전으로 남는다.
      const market = run(["plugin", "marketplace", "update", MARKETPLACE_NAME])
      const r = market.status === 0 ? run(["plugin", "update", PLUGIN_ID]) : market
      steps.push(
        r.status === 0
          ? `플러그인 갱신: ${plugin.version} → ${version}`
          : `플러그인 버전 기록 갱신 실패 (동작에는 영향 없음): ${(r.stderr || r.stdout).trim()}`,
      )
    } else if (plugin.enabled !== false) {
      steps.push(`플러그인 이미 설치됨: ${PLUGIN_ID}`)
    }
  }
  return { configPath: cfg.path, configCreated: cfg.created, steps, launchCommand: CHANNEL_LAUNCH_COMMAND }
}

export function uninstall({ run = defaultRunner, purge = false }: { run?: Runner; purge?: boolean } = {}): string[] {
  const steps: string[] = []
  const plugins = parseJsonArray<PluginRow>(run(["plugin", "list", "--json"]).stdout) ?? []
  if (plugins.some((p) => p.id === PLUGIN_ID)) {
    const r = run(["plugin", "uninstall", PLUGIN_ID])
    steps.push(r.status === 0 ? `플러그인 제거: ${PLUGIN_ID}` : `플러그인 제거 실패: ${(r.stderr || r.stdout).trim()}`)
  }
  const markets = parseJsonArray<MarketplaceRow>(run(["plugin", "marketplace", "list", "--json"]).stdout) ?? []
  if (markets.some((m) => m.name === MARKETPLACE_NAME)) {
    const r = run(["plugin", "marketplace", "remove", MARKETPLACE_NAME])
    steps.push(r.status === 0 ? `마켓플레이스 제거: ${MARKETPLACE_NAME}` : `마켓플레이스 제거 실패: ${(r.stderr || r.stdout).trim()}`)
  }
  if (purge) {
    // tokens.file 이 데이터 디렉토리 밖이어도 토큰은 지운다. 디렉토리 삭제는 데이터 디렉토리로만 한정한다.
    const tokensFile = loadConfig().tokens.file
    const dir = dataDir()
    fs.rmSync(dir, { recursive: true, force: true })
    fs.rmSync(tokensFile, { force: true })
    steps.push(`설정·토큰 삭제: ${dir}${path.dirname(tokensFile) === dir ? "" : `, ${tokensFile}`}`)
  }
  return steps
}

export type DoctorCheck = { ok: boolean; label: string; detail: string }

export function doctor({ packageDir, run = defaultRunner }: { packageDir: string; run?: Runner }): DoctorCheck[] {
  const checks: DoctorCheck[] = []
  const major = Number(process.versions.node.split(".")[0])
  checks.push({ ok: major >= 20, label: "Node.js", detail: process.versions.node })

  const cfg = loadConfig()
  const cfgFile = ensureConfigFile()
  checks.push({
    ok: azureConfigured(cfg),
    label: "설정 파일",
    detail: `${cfgFile.path} (${azureConfigured(cfg) ? "azure 설정됨" : "azure.tenantId/clientId/scopes 미설정"})`,
  })
  checks.push({
    ok: fs.existsSync(cfg.tokens.file),
    label: "로그인 토큰",
    detail: fs.existsSync(cfg.tokens.file) ? cfg.tokens.file : "없음 — /microsoft-teams:connect 로 로그인",
  })
  // To Do 는 선택 기능이라 꺼져 있거나 토큰 권한이 모자라도 실패로 보지 않는다.
  const todo = todoAccess(cfg.azure.scopes)
  checks.push({
    ok: true,
    label: "Microsoft To Do",
    detail:
      todo === "none"
        ? `${describeTodoAccess(todo)} — 쓰려면 azure.scopes 에 ${TODO_SCOPE_RW} 추가 (선택)`
        : `${describeTodoAccess(todo)} · 시간대 ${todoTimeZone(cfg)}` +
          (fs.existsSync(cfg.tokens.file) ? ` · ${describeTodoToken(new Auth(cfg, NOOP_LOG).tokenScope())}` : ""),
  })

  const version = run(["--version"])
  checks.push({ ok: version.status === 0, label: "claude CLI", detail: version.status === 0 ? version.stdout.trim() : "PATH 에 없음" })
  if (version.status === 0) {
    const markets = parseJsonArray<MarketplaceRow>(run(["plugin", "marketplace", "list", "--json"]).stdout) ?? []
    const m = markets.find((x) => x.name === MARKETPLACE_NAME)
    checks.push({
      ok: Boolean(m) && samePath(m ? marketplaceLocation(m) : undefined, packageDir),
      label: "마켓플레이스",
      detail: m ? `${marketplaceLocation(m)}${samePath(marketplaceLocation(m), packageDir) ? "" : ` (현재 패키지: ${packageDir})`}` : "미등록",
    })
    const plugins = parseJsonArray<PluginRow>(run(["plugin", "list", "--json"]).stdout) ?? []
    const p = plugins.find((x) => x.id === PLUGIN_ID)
    const version = packageVersion(packageDir)
    const stale = Boolean(p?.version && version && p.version !== version)
    checks.push({
      ok: Boolean(p?.enabled) && !stale,
      label: "플러그인",
      detail: !p
        ? "미설치"
        : !p.enabled
          ? "비활성화됨"
          : stale
            ? `기록된 버전 ${p.version} ≠ 패키지 ${version} — \`claudecode-ms-teams install\` 로 갱신`
            : `설치·활성화됨${p.version ? ` (${p.version})` : ""}`,
    })
  }

  let sessions: string[] = []
  try {
    sessions = fs.readdirSync(runDir()).filter((n) => /^\d+\.json$/.test(n))
  } catch {}
  const live = sessions
    .map((n) => readRegistry(Number(n.replace(".json", ""))))
    .filter((e) => e && isAlive(e.serverPid))
  checks.push({
    ok: true,
    label: "실행 중 세션",
    detail: `${live.length}개 (채팅방 연동 ${live.filter((e) => e?.active).length}개)`,
  })
  return checks
}
