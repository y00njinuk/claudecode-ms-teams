// claudecode-ms-teams 공용 코어: 설정 · 로그 · Azure AD 인증 · Microsoft Graph · 트리거 · Teams HTML 렌더링.
// 런타임 의존성 제로 — Node 내장 fetch 와 node: 빌트인만 사용한다.
import { createHash } from "node:crypto"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"

// ---------- 식별자 ----------

export const DATA_DIR_NAME = "claudecode-ms-teams"
export const CONFIG_FILE_NAME = "claudecode-ms-teams.json"
export const LOG_FILE_NAME = "claudecode-ms-teams.log"
export const TOKENS_FILE_NAME = "tokens.json"

export const PLUGIN_NAME = "microsoft-teams"
export const MARKETPLACE_NAME = "claudecode-ms-teams"
export const PLUGIN_ID = `${PLUGIN_NAME}@${MARKETPLACE_NAME}`
export const CHANNEL_LAUNCH_COMMAND = `claude --dangerously-load-development-channels plugin:${PLUGIN_ID}`
export const REQUIRED_SCOPES =
  "Chat.Create Chat.Read Chat.ReadBasic Chat.ReadWrite ChatMessage.Read ChatMessage.Send offline_access User.Read"

// Microsoft To Do 는 선택 기능이다. azure.scopes 에 이 scope 가 있을 때만 todo_* 도구를 노출한다 (REQUIRED_SCOPES 와 별개).
export const TODO_SCOPE_RW = "Tasks.ReadWrite"
export const TODO_SCOPE_R = "Tasks.Read"
export type TodoAccess = "readwrite" | "read" | "none"

// 토큰 응답의 scope 는 `https://graph.microsoft.com/Tasks.ReadWrite` 처럼 resource 가 붙어 올 수 있고, scope 는 대소문자를 구분하지 않는다.
function scopeName(scope: string): string {
  return scope.slice(scope.lastIndexOf("/") + 1).toLowerCase()
}

export function todoAccess(scopes: string | null | undefined): TodoAccess {
  const names = (scopes ?? "").split(/\s+/).filter(Boolean).map(scopeName)
  if (names.includes(TODO_SCOPE_RW.toLowerCase())) return "readwrite"
  if (names.includes(TODO_SCOPE_R.toLowerCase())) return "read"
  return "none"
}

// ---------- 트리거 ----------

// Microsoft Teams chat ID (group `19:...@thread.v2` / 1:1 `19:...@unq.gbl.spaces`).
export const CHAT_ID_RE = /(19:[A-Za-z0-9._~%!$&'()*+,;=:@-]+@(?:thread\.v2|unq\.gbl\.spaces))/

// 후행 경계 문자: ASCII 공백류 + 문장부호 + 유니코드 ellipsis + 전각 문장부호 + 전각 공백.
const TRIGGER_DELIMS = " \t\n,.!?~:;…，。！？　"
// 답장 verb 와 예전 호출 키워드. 호출 키워드는 이제 없어도 전달되며, 붙어 있으면 떼고 전달한다.
// alternation 은 longest-match 를 위해 길이 내림차순을 유지한다.
export const TRIGGER_RE =
  /^(?:#(?:allow-always|allow-reset|question|answer|claude|reject|agent|allow|skip|ask|ai)|가재야|막둥아)(?=[ \t\n,.!?~:;…，。！？　]|$)/iu

// 답장 verb 는 권한·질문 응답 전용이다. 매칭되면 세션에 주입하지 않는다.
export const REPLY_VERB_KEYWORDS = new Set(["#allow-always", "#allow-reset", "#allow", "#reject", "#answer", "#skip"])

// 답장 코드: 권한 요청은 p, 질문은 q 로 시작하는 base36 6자. 질문이 여러 개면 `-<n>` sub-code 를 붙인다.
// 코드는 대기 중인 요청이 여럿일 때 대상을 고르는 용도로만 쓴다.
export const REPLY_CODE_RE = /^([pq][0-9a-z]{5})(?:-([1-9][0-9]?))?$/i
const REPLY_CODE_CHARS = "0123456789abcdefghijklmnopqrstuvwxyz"
const REPLY_CODE_MAX_ATTEMPTS = 8

export type TriggerMatch =
  | { matched: false }
  | { matched: true; keyword: string; matchedSpan: string; payload: string }

// freeText=true 면 value 를 선택지와 대조하지 않고 그대로 답으로 쓴다 (`text:` 접두).
export type AnswerValue = { value: string; freeText: boolean }

export type ReplyCommand =
  // code 가 null 이면 대기 중인 요청이 하나일 때 그 요청에 적용한다.
  | { verb: "allow" | "allow-always" | "reject" | "skip"; code: string | null }
  | ({ verb: "answer"; code: string | null; subIndex: number | null } & AnswerValue)
  | { verb: "allow-reset" }

export function generateReplyCode(
  prefix: "p" | "q",
  existing: ReadonlySet<string>,
  random: () => number = Math.random,
): string | null {
  for (let attempt = 0; attempt < REPLY_CODE_MAX_ATTEMPTS; attempt++) {
    let code: string = prefix
    for (let i = 0; i < 5; i++) code += REPLY_CODE_CHARS[Math.floor(random() * REPLY_CODE_CHARS.length)]
    if (!existing.has(code)) return code
  }
  return null
}

// 정규화: 제로폭 문자 제거 → NFKC → CRLF 통일 → NBSP/전각공백을 공백으로 → trim.
// NFKC 는 전각 '＃'(U+FF03) → '#', 조합형 자모 → 완성형으로 흡수한다.
export function normalizeForTrigger(raw: string): string {
  let s = raw.replace(/[​﻿‎‏]/g, "")
  s = s.normalize("NFKC")
  s = s.replace(/\r\n?/g, "\n")
  s = s.replace(/[ 　]/g, " ")
  return s.trim()
}

// 세션에 넘기는 본문용 정리. 매칭용 normalizeForTrigger 와 달리 NFKC 를 하지 않는다
// (ㅇㅇ → 조합형 자모, ① → 1 처럼 사용자가 쓴 문자를 바꾸지 않기 위함).
export function cleanInbound(raw: string): string {
  return raw.replace(/[\u200B\uFEFF\u200E\u200F]/g, "").replace(/\r\n?/g, "\n").replace(/[\u00A0\u3000]/g, " ").trim()
}

// detectTrigger 의 payload 는 NFKC 된 문자열이다. NFKC 후 matchedSpan 과 같아지는 원문 앞부분을 떼어내 본문은 원문 그대로 돌려준다.
export function payloadVerbatim(text: string, match: Extract<TriggerMatch, { matched: true }>): string {
  const limit = Math.min(text.length, match.matchedSpan.length * 4 + 8)
  for (let k = 1; k <= limit; k++) {
    if (text.slice(0, k).normalize("NFKC") === match.matchedSpan) return text.slice(k).replace(/^[ \t\n]+/, "")
  }
  return match.payload
}

// offset 0 에서만 매칭. payload 는 키워드와 연속 후행 구분자를 소비한 나머지.
export function detectTrigger(raw: string): TriggerMatch {
  const normalized = normalizeForTrigger(raw)
  const m = TRIGGER_RE.exec(normalized)
  if (!m) return { matched: false }
  let end = m[0].length
  while (end < normalized.length && TRIGGER_DELIMS.includes(normalized[end])) end++
  return {
    matched: true,
    keyword: m[0].toLowerCase(),
    matchedSpan: normalized.slice(0, end),
    payload: normalized.slice(end).replace(/^[ \t\n]+/, ""),
  }
}

// `text:` 뒤는 쉼표·개행을 포함해 끝까지 자유 텍스트로 받는다. 빈 답이면 null.
export function parseAnswerValue(raw: string): AnswerValue | null {
  const trimmed = raw.trim()
  const text = /^text\s*:\s*(.*)$/is.exec(trimmed)
  const value = text ? text[1].trim() : trimmed
  return value ? { value, freeText: Boolean(text) } : null
}

// keyword 는 detectTrigger 가 돌려준 소문자 verb, payload 는 그 뒤의 나머지.
// `#answer` 의 첫 단어는 knownCodes(대기 중인 질문 코드)에 있을 때만 코드로 본다.
// 코드 형식(p/q + 5자)과 겹치는 답(python, pytest 등)을 코드로 오인하지 않기 위함이다.
export function parseReplyCommand(keyword: string, payload: string, knownCodes: ReadonlySet<string> = new Set()): ReplyCommand | null {
  if (!REPLY_VERB_KEYWORDS.has(keyword)) return null
  const trimmed = payload.trim()
  if (keyword === "#allow-reset") return trimmed ? null : { verb: "allow-reset" }
  const spaceIdx = trimmed.search(/\s/)
  const m = REPLY_CODE_RE.exec(spaceIdx < 0 ? trimmed : trimmed.slice(0, spaceIdx))
  const rest = spaceIdx < 0 ? "" : trimmed.slice(spaceIdx + 1)
  if (keyword !== "#answer") {
    const verb = keyword.slice(1) as "allow" | "allow-always" | "reject" | "skip"
    if (!trimmed) return { verb, code: null }
    if (!m || m[2] || rest) return null
    return { verb, code: m[1].toLowerCase() }
  }
  const code = m && knownCodes.has(m[1].toLowerCase()) ? m[1].toLowerCase() : null
  const answer = parseAnswerValue(code ? rest : trimmed)
  if (!answer) return null
  return { verb: "answer", code, subIndex: code && m?.[2] ? Number(m[2]) : null, ...answer }
}

// 플러그인이 보내는 모든 메시지는 `[<제목> · <세션 8자 또는 ->]` 헤더로 시작한다 (renderHeader).
// 같은 계정으로 발신하는 다른 세션의 응답·알림을 이 헤더로 알아보고 주입하지 않는다 —
// 키워드 없이 모든 메시지를 전달하므로, 걸러내지 않으면 세션끼리 응답을 주고받는 루프가 된다.
// 세션 자리는 영문자를 하나 이상 포함한 hex 8자로 제한해 `[회의록 · 20261120]` 같은 사용자 메시지와 겹치지 않게 한다.
const SESSION_TAG = "(?=[0-9]*[a-f])[0-9a-f]{8}"
const PLUGIN_HEADER_RE = new RegExp(`^\\[[^\\]\\n]+ · (?:${SESSION_TAG}|-)\\]`)

export function isPluginMessage(text: string): boolean {
  return PLUGIN_HEADER_RE.test(normalizeForTrigger(text))
}

// UUID 앞 8자. 헤더 판별 형식(SESSION_TAG)에 맞지 않으면 "-" 를 쓴다 (숫자만 나온 경우 포함).
export function shortSessionId(sessionId: string): string {
  const short = sessionId.slice(0, 8).toLowerCase()
  return new RegExp(`^${SESSION_TAG}$`).test(short) ? short : "-"
}

// symlink 는 풀지 않고 `..` 와 상대 경로만 정규화해 workspace 하위인지 본다.
// `~` 로 시작하는 경로는 도구가 홈 디렉토리로 펼칠 수 있으므로 workspace 밖으로 본다.
export function isInsideWorkspace(target: string, workspace: string): boolean {
  if (!target || !workspace || target.startsWith("~")) return false
  const ws = path.resolve(workspace)
  const rel = path.relative(ws, path.resolve(ws, target))
  return rel !== ".." && !rel.startsWith(`..${path.sep}`) && !path.isAbsolute(rel)
}

// ---------- 설정 ----------

export type Config = {
  azure: { tenantId: string; clientId: string; scopes: string }
  polling: { intervalMs: number; backoff1Ms: number; backoff2Ms: number }
  // 로그인한 본인 외에 세션을 조작할 수 있는 Entra 사용자 ID 목록.
  access: { allowedUserIds: string[] }
  log: { file: string; enabled: boolean }
  tokens: { file: string }
  // timeZone: To Do 기한 · 알림을 해석할 시간대 (IANA 또는 Windows 이름). 빈 문자열이면 시스템 시간대.
  todo: { timeZone: string }
}

export const DEFAULT_POLL_MS = 5000
export const DEFAULT_BACKOFF1_MS = 10 * 60 * 1000
export const DEFAULT_BACKOFF2_MS = 60 * 60 * 1000

// HOME 미설정 + /etc/passwd 에 uid 가 없으면 os.homedir() 가 libuv ENOENT 로 throw 한다.
export function homeDir(): string {
  try {
    return os.homedir()
  } catch {
    throw new Error("홈 디렉토리를 확인할 수 없습니다 (HOME 미설정 + /etc/passwd 에 uid 없음). XDG_CONFIG_HOME 을 설정하세요.")
  }
}

export function dataDir(): string {
  const xdg = process.env.XDG_CONFIG_HOME
  return path.join(xdg || path.join(homeDir(), ".config"), DATA_DIR_NAME)
}

export function configFilePath(): string {
  return path.join(dataDir(), CONFIG_FILE_NAME)
}

export function expandPath(p: string): string {
  if (!p) return p
  if (p === "~") return homeDir()
  if (p.startsWith("~/")) return path.join(homeDir(), p.slice(2))
  return p
}

export function defaultConfig(): Config {
  const dd = dataDir()
  return {
    azure: { tenantId: "", clientId: "", scopes: "" },
    polling: { intervalMs: DEFAULT_POLL_MS, backoff1Ms: DEFAULT_BACKOFF1_MS, backoff2Ms: DEFAULT_BACKOFF2_MS },
    access: { allowedUserIds: [] },
    log: { file: path.join(dd, LOG_FILE_NAME), enabled: false },
    tokens: { file: path.join(dd, TOKENS_FILE_NAME) },
    todo: { timeZone: "" },
  }
}

export function mergeConfig(base: Config, partial: unknown): Config {
  if (!partial || typeof partial !== "object") return base
  const p = partial as Record<string, unknown>
  const pick = <T extends object>(k: keyof Config, def: T): T => {
    const v = p[k as string]
    return v && typeof v === "object" && !Array.isArray(v) ? ({ ...def, ...(v as object) } as T) : def
  }
  const access = pick("access", base.access)
  const todo = pick("todo", base.todo)
  return {
    azure: pick("azure", base.azure),
    polling: pick("polling", base.polling),
    access: {
      allowedUserIds: Array.isArray(access.allowedUserIds)
        ? access.allowedUserIds.filter((x): x is string => typeof x === "string" && x.length > 0)
        : [],
    },
    log: pick("log", base.log),
    tokens: pick("tokens", base.tokens),
    todo: { timeZone: typeof todo.timeZone === "string" ? todo.timeZone.trim() : "" },
  }
}

export function azureConfigured(cfg: Config): boolean {
  return Boolean(cfg.azure.tenantId && cfg.azure.clientId && cfg.azure.scopes)
}

// 파일이 없으면 기본값으로 생성한다. 파싱 실패 시에는 기본값으로 계속 진행 (세션 방해 금지).
export function ensureConfigFile(): { path: string; created: boolean } {
  const cfgPath = configFilePath()
  if (fs.existsSync(cfgPath)) return { path: cfgPath, created: false }
  fs.mkdirSync(path.dirname(cfgPath), { recursive: true })
  fs.writeFileSync(cfgPath, JSON.stringify(defaultConfig(), null, 2) + "\n", "utf8")
  return { path: cfgPath, created: true }
}

export function loadConfig(): Config {
  const base = defaultConfig()
  let loaded = base
  try {
    loaded = mergeConfig(base, JSON.parse(fs.readFileSync(configFilePath(), "utf8")))
  } catch (err: unknown) {
    if ((err as NodeJS.ErrnoException | undefined)?.code === "ENOENT") {
      try {
        ensureConfigFile()
      } catch {
        // 생성 실패해도 기본값으로 동작
      }
    }
  }
  loaded.log.file = expandPath(loaded.log.file)
  loaded.tokens.file = expandPath(loaded.tokens.file)
  return loaded
}

// ---------- 로그 ----------

type L = (msg: string, data?: object) => void
export type Log = { info: L; warn: L; error: L; debug: L }

export const NOOP_LOG: Log = { info: () => {}, warn: () => {}, error: () => {}, debug: () => {} }

// enabled=false 면 파일을 만들지 않는다. true 면 서버 기동 시 truncate 후 append.
export function makeLogger(cfg: Config, tag: string): Log {
  if (!cfg.log.enabled) return NOOP_LOG
  try {
    fs.mkdirSync(path.dirname(cfg.log.file), { recursive: true })
    fs.writeFileSync(cfg.log.file, "")
  } catch {
    // truncate 실패해도 append 는 시도
  }
  const write = (level: string, msg: string, data?: object) => {
    try {
      fs.appendFileSync(
        cfg.log.file,
        `${new Date().toISOString()} [${level.toUpperCase()}] ${tag} ${msg}${data ? " " + JSON.stringify(data) : ""}\n`,
      )
    } catch {
      // 로그 실패는 무시
    }
  }
  return {
    info: (m, d) => write("info", m, d),
    warn: (m, d) => write("warn", m, d),
    error: (m, d) => write("error", m, d),
    debug: (m, d) => write("debug", m, d),
  }
}

// ---------- 유틸 ----------

// unref: MCP 서버는 stdin 이 프로세스를 유지하므로 폴링 타이머가 종료를 막지 않게 한다.
export const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms).unref?.())
export const fingerprint = (s: string) => createHash("sha1").update(s).digest("hex")

export function delayFor(elapsedMs: number, cfg: Config): number {
  const p = cfg.polling
  if (elapsedMs < p.backoff1Ms) return p.intervalMs
  if (elapsedMs < p.backoff2Ms) return p.intervalMs * 3
  return p.intervalMs * 6
}

// ---------- 인증 (Azure AD device code flow) ----------

type TokenSet = { access_token: string; refresh_token: string; expires_at?: number; expires_in?: number; scope?: string }
type TokenError = { error?: string; suberror?: string; error_codes?: unknown; error_description?: string }

// 동의(consent)가 없는 scope 를 refresh 요청에 넣었을 때 Azure AD 가 주는 오류 (AADSTS65001).
// refresh token 자체는 유효하므로 폐기(revoked)로 보면 안 된다.
export function isConsentRequired(body: TokenError): boolean {
  if (body.error !== "invalid_grant") return false
  return (
    body.suberror === "consent_required" ||
    (Array.isArray(body.error_codes) && body.error_codes.includes(65001)) ||
    /\bAADSTS65001\b/.test(body.error_description ?? "")
  )
}

export function withoutTodoScopes(scopes: string): string {
  return scopes
    .split(/\s+/)
    .filter((s) => s && !scopeName(s).startsWith("tasks."))
    .join(" ")
}

export type DeviceCode = { verificationUri: string; userCode: string; expiresAt: number }

export type AuthEvents = {
  onLoginComplete?: () => void
  onLoginFailed?: (reason: string) => void
  // refresh token 폐기(invalid_grant). device flow 는 자동 시작하지 않는다 —
  // 무인 상태에서 코드가 응답에 실려 Teams 로 새어나가면 채팅방 멤버가 대신 로그인할 수 있다.
  onReauthRequired?: () => void
}

const REFRESH_LEAD_MS = 5 * 60 * 1000
const REFRESH_RETRY_MS = 30 * 1000
const REFRESH_MIN_MS = 15 * 1000

export class Auth {
  private refreshing: Promise<string | null> | null = null
  private deviceFlowRunning = false
  private pending: DeviceCode | null = null
  private devicePhase1: Promise<void> | null = null
  private refreshTimer: ReturnType<typeof setTimeout> | null = null
  private disposed = false
  private revoked = false
  // To Do scope 에 동의가 없어 refresh 를 To Do scope 없이 하는 중. To Do 도구가 refreshWithTodoScopes() 로 다시 시도한다.
  private todoConsentMissing = false

  constructor(
    private cfg: Config,
    private log: Log,
    private events: AuthEvents = {},
  ) {}

  pendingDeviceCode(): DeviceCode | null {
    if (this.pending && Date.now() < this.pending.expiresAt) return this.pending
    return null
  }

  awaitDeviceCode(): Promise<void> {
    return this.devicePhase1 ?? Promise.resolve()
  }

  hasRefreshToken(): boolean {
    return !this.revoked && Boolean(this.load()?.refresh_token)
  }

  // 저장된 토큰의 scope. 여러 세션이 tokens.json 을 공유하므로 설정이 아니라 파일을 본다. 모르면 null.
  tokenScope(): string | null {
    const scope = this.load()?.scope
    return typeof scope === "string" ? scope : null
  }

  todoConsentRequired(): boolean {
    return this.todoConsentMissing
  }

  // 설정의 scope 전체로 refresh 한다 (다른 세션이 예전 scope 로 토큰을 덮어썼거나 그 사이 관리자 동의가 난 경우).
  // 진행 중인 refresh 는 To Do scope 없이 보낸 것일 수 있으므로 끝나기를 기다린 뒤 새로 요청한다.
  async refreshWithTodoScopes(): Promise<string | null> {
    if (this.refreshing) await this.refreshing
    this.todoConsentMissing = false
    return this.forceRefresh()
  }

  private tokenUrl() {
    return `https://login.microsoftonline.com/${this.cfg.azure.tenantId}/oauth2/v2.0/token`
  }

  private load(): TokenSet | null {
    try {
      return JSON.parse(fs.readFileSync(this.cfg.tokens.file, "utf8")) as TokenSet
    } catch {
      return null
    }
  }

  // 여러 세션의 서버가 같은 파일을 읽고 쓰므로 tmp + rename 으로 원자적으로 교체한다.
  private save(t: TokenSet) {
    const file = this.cfg.tokens.file
    fs.mkdirSync(path.dirname(file), { recursive: true })
    const tmp = `${file}.${process.pid}.tmp`
    fs.writeFileSync(tmp, JSON.stringify(t, null, 2), { mode: 0o600 })
    try {
      fs.chmodSync(tmp, 0o600)
    } catch {}
    fs.renameSync(tmp, file)
  }

  dispose() {
    this.disposed = true
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer)
      this.refreshTimer = null
    }
  }

  private scheduleRefreshFromToken(t: TokenSet | null): void {
    if (this.disposed) return
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer)
      this.refreshTimer = null
    }
    if (!t?.refresh_token) return
    const expiresAt = t.expires_at ?? Date.now() + (t.expires_in ?? 3600) * 1000
    const delay = Math.max(REFRESH_MIN_MS, expiresAt - Date.now() - REFRESH_LEAD_MS)
    this.refreshTimer = setTimeout(() => {
      this.refreshTimer = null
      void this.backgroundRefresh()
    }, delay)
    this.refreshTimer.unref?.()
    this.log.debug("token.refresh.scheduled", { inMs: delay })
  }

  // 저장된 refresh token 이 있으면 background refresh 를 예약한다. 없으면 아무것도 하지 않는다.
  bootstrap(): void {
    const saved = this.load()
    if (saved?.refresh_token) this.scheduleRefreshFromToken(saved)
  }

  private async backgroundRefresh(): Promise<void> {
    if (this.disposed) return
    const cur = this.load()
    if (!cur?.refresh_token) return
    const next = await this.refresh(cur)
    if (this.disposed || this.revoked) return
    if (next) {
      this.scheduleRefreshFromToken(this.load())
    } else {
      this.refreshTimer = setTimeout(() => {
        this.refreshTimer = null
        void this.backgroundRefresh()
      }, REFRESH_RETRY_MS)
      this.refreshTimer.unref?.()
    }
  }

  // 토큰이 없으면 null. device flow 는 startDeviceFlow() 로만 명시적으로 시작한다 (lazy connect).
  async getAccessToken(): Promise<string | null> {
    if (this.revoked) return null
    const t = this.load()
    if (!t?.refresh_token) return null
    if (t.expires_at && Date.now() < t.expires_at - 60_000) return t.access_token
    const at = await this.refresh(t)
    if (at) this.scheduleRefreshFromToken(this.load())
    return at
  }

  async forceRefresh(): Promise<string | null> {
    if (this.revoked) return null
    const t = this.load()
    if (!t?.refresh_token) return null
    const at = await this.refresh(t)
    if (at) this.scheduleRefreshFromToken(this.load())
    return at
  }

  private refresh(t: TokenSet): Promise<string | null> {
    if (this.refreshing) return this.refreshing
    const scopes = this.todoConsentMissing ? withoutTodoScopes(this.cfg.azure.scopes) : this.cfg.azure.scopes
    this.refreshing = this.requestRefresh(t, scopes).finally(() => {
      this.refreshing = null
    })
    return this.refreshing
  }

  private async requestRefresh(t: TokenSet, scopes: string): Promise<string | null> {
    try {
      const res = await fetch(this.tokenUrl(), {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "refresh_token",
          client_id: this.cfg.azure.clientId,
          refresh_token: t.refresh_token,
          scope: scopes,
        }),
      })
      const body = (await res.json()) as TokenSet & TokenError
      if (!res.ok) {
        // To Do scope 를 설정에 추가했지만 동의가 없으면 Teams 까지 멈추지 않도록 To Do scope 를 빼고 다시 받는다.
        // 동의 없음의 오류 형태(AADSTS65001 · 70000 등)를 모두 알 수 없으므로 Tasks 를 요청한 invalid_grant 는
        // Tasks 없이 한 번 다시 받아 보고, 그것도 실패해야 폐기로 본다.
        if (body.error === "invalid_grant" && todoAccess(scopes) !== "none") {
          this.log.warn("token.todoScopeRejected", {
            consentRequired: isConsentRequired(body),
            suberror: body.suberror,
            errorCodes: body.error_codes,
          })
          this.todoConsentMissing = true
          return this.requestRefresh(t, withoutTodoScopes(scopes))
        }
        if (body.error === "invalid_grant") {
          this.log.warn("token.revoked")
          if (!this.revoked) {
            this.revoked = true
            this.events.onReauthRequired?.()
          }
        } else {
          this.log.error("token.refresh.failed", { error: body.error })
        }
        return null
      }
      const nt: TokenSet = {
        ...body,
        refresh_token: body.refresh_token ?? t.refresh_token,
        expires_at: Date.now() + (body.expires_in ?? 3600) * 1000,
      }
      this.save(nt)
      this.log.debug("token.refreshed")
      return nt.access_token
    } catch (e) {
      this.log.error("token.refresh.network", { error: String(e) })
      return null
    }
  }

  startDeviceFlow(): void {
    if (this.deviceFlowRunning || this.disposed) return
    this.deviceFlowRunning = true
    let resolvePhase1!: () => void
    this.devicePhase1 = new Promise<void>((r) => {
      resolvePhase1 = r
    })

    void (async () => {
      try {
        const res = await fetch(`https://login.microsoftonline.com/${this.cfg.azure.tenantId}/oauth2/v2.0/devicecode`, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({ client_id: this.cfg.azure.clientId, scope: this.cfg.azure.scopes }),
        })
        const d = (await res.json()) as {
          device_code: string
          user_code: string
          verification_uri: string
          interval?: number
          expires_in?: number
          error?: string
          error_description?: string
        }
        if (!res.ok) {
          this.log.error("devicecode.request.failed", { error: d.error })
          resolvePhase1()
          this.events.onLoginFailed?.(d.error_description ?? d.error ?? "unknown")
          return
        }
        const expiresInMs = (d.expires_in ?? 900) * 1000
        this.pending = { verificationUri: d.verification_uri, userCode: d.user_code, expiresAt: Date.now() + expiresInMs }
        this.log.info("devicecode.obtained", { verificationUri: d.verification_uri })
        resolvePhase1()
        const interval = (d.interval ?? 5) * 1000
        const deadline = Date.now() + expiresInMs
        while (Date.now() < deadline && !this.disposed) {
          await sleep(interval)
          const tr = await fetch(this.tokenUrl(), {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({
              grant_type: "urn:ietf:params:oauth:grant-type:device_code",
              client_id: this.cfg.azure.clientId,
              device_code: d.device_code,
            }),
          })
          const tb = (await tr.json()) as TokenSet & { error?: string }
          if (tr.ok) {
            this.save({ ...tb, expires_at: Date.now() + (tb.expires_in ?? 3600) * 1000 })
            this.pending = null
            this.revoked = false
            this.todoConsentMissing = false
            this.scheduleRefreshFromToken(this.load())
            this.log.info("devicecode.login.complete")
            this.events.onLoginComplete?.()
            return
          }
          if (tb.error === "authorization_pending" || tb.error === "slow_down") continue
          this.log.error("devicecode.polling.failed", { error: tb.error })
          this.pending = null
          this.events.onLoginFailed?.(tb.error ?? "unknown")
          return
        }
        this.log.warn("devicecode.expired")
        this.pending = null
      } catch (e) {
        this.log.error("devicecode.network", { error: String(e) })
        this.pending = null
        resolvePhase1()
        this.events.onLoginFailed?.(String(e))
      } finally {
        this.deviceFlowRunning = false
      }
    })()
  }
}

// ---------- Teams HTML 렌더링 ----------

const CODE_FONT_STYLE = "font-family:Consolas,'Courier New',monospace"
// class="language-*" 를 붙이면 Teams 가 line number 를 붙이므로 class 없는 <pre><code> 를 쓴다.
export const CODE_BLOCK_STYLE = `${CODE_FONT_STYLE};background:#f4f4f4;padding:8px;border-left:3px solid #ccc`
export const INLINE_CODE_STYLE = `${CODE_FONT_STYLE};background:#f4f4f4;padding:1px 4px;border-radius:3px`
export const MAX_HTML_BYTES = 25_000

export function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
}

const TABLE_SEP_RE = /^\s*\|?\s*:?-{3,}:?(\s*\|\s*:?-{3,}:?)+\s*\|?\s*$/
const TABLE_ROW_RE = /^\s*\|.*\|\s*$/
type Align = "left" | "center" | "right"

function splitRow(row: string): string[] {
  return row.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim())
}

function parseAlign(sep: string): Align[] {
  return splitRow(sep).map((c) => {
    const left = c.startsWith(":")
    const right = c.endsWith(":")
    if (left && right) return "center"
    if (right) return "right"
    return "left"
  })
}

// 인라인 패턴은 줄바꿈과 블록 placeholder 를 넘지 않는다.
function inlineMd(t: string): string {
  return escapeHtml(t)
    .replace(/`([^`\n\u0000]+)`/g, `<span style="${INLINE_CODE_STYLE}">$1</span>`)
    .replace(/\*\*([^*\n\u0000]+)\*\*/g, "<b>$1</b>")
}

function renderTable(header: string, sep: string, rows: string[]): string {
  const aligns = parseAlign(sep)
  const th = splitRow(header)
    .map((c, i) => `<th align="${aligns[i] ?? "left"}">${inlineMd(c)}</th>`)
    .join("")
  const trs = rows
    .map((r) => `<tr>${splitRow(r).map((c, i) => `<td align="${aligns[i] ?? "left"}">${inlineMd(c)}</td>`).join("")}</tr>`)
    .join("")
  return `<table border="1" style="border-collapse:collapse"><thead><tr>${th}</tr></thead><tbody>${trs}</tbody></table>`
}

export function renderCodeBlock(code: string): string {
  return `<pre style="${CODE_BLOCK_STYLE}"><code>${escapeHtml(code)}</code></pre>`
}

// 렌더된 블록(코드·표)은 placeholder 로 빼두고 나머지 텍스트만 escape 한 뒤 되돌린다.
// 렌더 결과처럼 생긴 원문 HTML 을 분할 기준으로 쓰면 escape 없이 통과해 버린다.
const BLOCK_MARK = "\u0000"
const BLOCK_RE = /\u0000(\d+)\u0000/g

// 코드펜스는 라인 시작의 백틱 3개 이상만 인식한다. 닫는 펜스가 없으면 끝까지 코드 블록으로 처리.
export function mdToHtml(md: string): string {
  const blocks: string[] = []
  const hold = (html: string) => `${BLOCK_MARK}${blocks.push(html) - 1}${BLOCK_MARK}`
  const lines = md.replaceAll(BLOCK_MARK, "").split("\n")
  const out: string[] = []
  const fenceRe = /^(\s*)(`{3,})\s*([A-Za-z0-9_+.-]*)\s*$/
  let i = 0
  while (i < lines.length) {
    const m = lines[i].match(fenceRe)
    if (m) {
      const fence = m[2]
      const inner: string[] = []
      let j = i + 1
      while (j < lines.length) {
        const closeM = lines[j].match(/^(\s*)(`{3,})\s*$/)
        if (closeM && closeM[2].length >= fence.length) break
        inner.push(lines[j++])
      }
      out.push(hold(renderCodeBlock(inner.join("\n"))))
      i = j < lines.length ? j + 1 : j
      continue
    }
    if (i + 1 < lines.length && TABLE_ROW_RE.test(lines[i]) && TABLE_SEP_RE.test(lines[i + 1])) {
      const rows: string[] = []
      let j = i + 2
      while (j < lines.length && TABLE_ROW_RE.test(lines[j])) rows.push(lines[j++])
      out.push(hold(renderTable(lines[i], lines[i + 1], rows)))
      i = j
      continue
    }
    out.push(lines[i++])
  }
  return inlineMd(out.join("\n"))
    .replace(/\n/g, "<br>")
    .replace(BLOCK_RE, (_, idx: string) => blocks[Number(idx)] ?? "")
}

// 대괄호·개행은 isPluginMessage 가 헤더를 알아보는 데 방해되므로 제목에서 뺀다.
// 수신 쪽이 NFKC 로 정규화하므로 전각 ］ 등도 NFKC 후에 걸러야 한다.
export function renderHeader(label: string, sessionShort: string): string {
  const safe = label.normalize("NFKC").replace(/[[\]\r\n]+/g, " ").trim() || "Claude Code"
  return `<b>[${escapeHtml(safe)} · ${escapeHtml(sessionShort)}]</b>`
}

export function renderOutbound(title: string, sessionShort: string, text: string, maxBytes = MAX_HTML_BYTES): string {
  const header = `${renderHeader(title, sessionShort)}<br>`
  let body = text
  let html = header + mdToHtml(body)
  while (Buffer.byteLength(html, "utf8") > maxBytes && body.length > 100) {
    body = body.slice(0, Math.floor(body.length * 0.8))
    html = header + mdToHtml(body) + "<br>…(생략)"
  }
  return html
}

function decodeCodePoint(cp: number): string | null {
  return cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : null
}

export function htmlToText(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#(\d{1,7});/g, (m, d: string) => decodeCodePoint(Number(d)) ?? m)
    .replace(/&#x([0-9a-f]{1,6});/gi, (m, h: string) => decodeCodePoint(parseInt(h, 16)) ?? m)
    .replace(/&amp;/g, "&") // 이중 디코드 방지를 위해 반드시 마지막
    .trim()
}

// ---------- Microsoft Graph ----------

export type ChatMsg = {
  id: string
  messageType: string
  deletedDateTime?: string | null
  from?: { user?: { id?: string; displayName?: string } | null } | null
  body: { contentType: string; content: string }
}

export type ChatSummary = { id: string; topic: string; chatType: string; lastActivity: string | null }

export type Me = { id: string; displayName: string }

export const SSL_HINT = "사내망 프록시 인증서 문제로 보임. NODE_EXTRA_CA_CERTS=<사내 CA 번들 경로> 설정 후 Claude Code 재시작"

export type GraphMethod = "GET" | "POST" | "PATCH" | "DELETE"

// ok=false 의 error: auth(토큰 없음 · 401 재시도 실패) · network · ssl · badBody(본문 파싱 실패) · http(그 외 상태 코드).
// status 는 HTTP 응답이 없으면 0. body 는 파싱한 응답 본문 (Graph 오류면 `{ error: { code, message } }`).
export type GraphResult =
  | { ok: true; status: number; body: any }
  | { ok: false; status: number; body: any; error: "auth" | "network" | "ssl" | "badBody" | "http" }

export class Graph {
  constructor(
    private auth: Auth,
    private log: Log,
  ) {}

  private chatBase(chatId: string) {
    return `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(chatId)}`
  }

  // 실패 원인이 필요한 호출용. 401 은 forceRefresh 후 1회, 429 는 Retry-After 대기 후 1회 재시도한다.
  async request(method: GraphMethod, url: string, body?: object, retried = false): Promise<GraphResult> {
    const token = await this.auth.getAccessToken()
    if (!token) return { ok: false, status: 0, body: null, error: "auth" }
    let res: Response
    try {
      res = await fetch(url, {
        method,
        headers: { Authorization: `Bearer ${token}`, ...(body ? { "Content-Type": "application/json" } : {}) },
        body: body ? JSON.stringify(body) : undefined,
      })
    } catch (e) {
      const msg = String((e as { cause?: unknown })?.cause ?? e)
      this.log.warn("graph.network", { url, error: msg })
      if (/certificat|UNABLE_TO_VERIFY|self.signed/i.test(msg)) {
        this.log.error("graph.ssl.hint", { hint: SSL_HINT })
        return { ok: false, status: 0, body: null, error: "ssl" }
      }
      return { ok: false, status: 0, body: null, error: "network" }
    }
    if (res.status === 401 && !retried) {
      const t2 = await this.auth.forceRefresh()
      if (!t2) return { ok: false, status: 401, body: null, error: "auth" }
      return this.request(method, url, body, true)
    }
    if (res.status === 429) {
      const ra = Number(res.headers.get("Retry-After") ?? "10")
      this.log.warn("graph.throttled", { retryAfterSec: ra })
      if (!retried) {
        await sleep(ra * 1000)
        return this.request(method, url, body, true)
      }
    }
    if (!res.ok) {
      this.log.error("graph.error", { url, status: res.status })
      let errorBody: unknown = null
      try {
        errorBody = await res.json()
      } catch {}
      return { ok: false, status: res.status, body: errorBody, error: "http" }
    }
    if (res.status === 204) return { ok: true, status: 204, body: {} }
    // 프록시·캡티브 포털의 HTML 응답이나 본문 중간 끊김이 서버 프로세스를 죽이지 않도록 실패로 흡수한다.
    try {
      return { ok: true, status: res.status, body: await res.json() }
    } catch (e) {
      this.log.warn("graph.badBody", { url, error: String(e) })
      return { ok: false, status: res.status, body: null, error: "badBody" }
    }
  }

  // 실패는 모두 null (원인은 로그에만 남는다).
  private async call(method: "GET" | "POST", url: string, body?: object): Promise<any | null> {
    const r = await this.request(method, url, body)
    return r.ok ? r.body : null
  }

  async me(): Promise<Me | null> {
    const r = await this.call("GET", "https://graph.microsoft.com/v1.0/me?%24select=id,displayName")
    if (!r?.id) return null
    return { id: String(r.id), displayName: String(r.displayName ?? "") }
  }

  async sendMessage(chatId: string, html: string): Promise<string | null> {
    const r = await this.call("POST", `${this.chatBase(chatId)}/messages`, { body: { contentType: "html", content: html } })
    return (r?.id as string | undefined) ?? null
  }

  // number = 성공 (빈 채팅방은 0), null = 호출 실패.
  // 실패를 0 으로 축약하면 cursor=0 이 되어 최근 50건 과거 이력이 모두 답신으로 처리된다.
  async latestMessageId(chatId: string): Promise<number | null> {
    const r = await this.call("GET", `${this.chatBase(chatId)}/messages?%24top=1&%24orderby=createdDateTime%20desc`)
    if (r == null) return null
    const id = r?.value?.[0]?.id as string | undefined
    return id ? Number(id) : 0
  }

  // null = 호출 실패 (커서를 전진시키지 않는다).
  async messagesSince(chatId: string, cursor: number): Promise<ChatMsg[] | null> {
    const r = await this.call("GET", `${this.chatBase(chatId)}/messages?%24top=50&%24orderby=createdDateTime%20desc`)
    if (r == null) return null
    const value = (r.value as ChatMsg[] | undefined) ?? []
    return value
      .filter((m) => m.messageType === "message" && Number(m.id) > cursor && !m.deletedDateTime && Boolean(m.body?.content))
      .sort((a, b) => Number(a.id) - Number(b.id))
  }

  async listChats(top: number): Promise<ChatSummary[] | null> {
    const url =
      `https://graph.microsoft.com/v1.0/me/chats?` +
      `%24top=${encodeURIComponent(String(top))}` +
      `&%24orderby=${encodeURIComponent("lastMessagePreview/createdDateTime desc")}` +
      `&%24expand=${encodeURIComponent("lastMessagePreview")}`
    const r = await this.call("GET", url)
    if (!r) return null
    const rows = (r.value as Array<Record<string, unknown>> | undefined) ?? []
    return rows.map((row) => ({
      id: String(row.id ?? ""),
      topic: typeof row.topic === "string" && row.topic ? row.topic : "",
      chatType: String(row.chatType ?? "unknown"),
      lastActivity:
        typeof row.lastUpdatedDateTime === "string"
          ? row.lastUpdatedDateTime
          : ((row.lastMessagePreview as { createdDateTime?: string } | undefined)?.createdDateTime ?? null),
    }))
  }
}
