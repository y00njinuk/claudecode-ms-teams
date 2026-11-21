// claudecode-ms-teams send: Claude Code 세션 없이 Markdown 한 건을 Teams 채팅방으로 보내고 결과를 종료 코드로 돌려준다.
// 예약 작업의 보고처럼 무인 스크립트에서 쓴다. 저장된 로그인 토큰만 쓰고 device flow 는 시작하지 않는다
// (코드가 무인 출력으로 새면 다른 사람이 대신 로그인할 수 있다).
import * as fs from "node:fs"
import * as path from "node:path"
import {
  Auth,
  azureConfigured,
  CHAT_ID_RE,
  configFilePath,
  Graph,
  loadConfig,
  renderOutbound,
  type Config,
  type Log,
} from "./core.js"

export const SEND_EXIT = { ok: 0, failed: 1, usage: 2, login: 3 } as const

export const SEND_USAGE = `claudecode-ms-teams send --chat <19:...> [--file <경로>] [--title <제목>]
  Markdown 한 건을 Teams 채팅방으로 보낸다. --file 을 생략하거나 - 로 주면 stdin 을 읽는다.
  메시지는 [<제목> · -] 헤더로 시작한다 (기본 제목: 현재 디렉토리 이름).
  종료 코드: 0 전송됨 · 1 발신 실패 · 2 사용법 오류 · 3 설정 · 로그인 필요`

// 세션 메시지와 구분할 세션 자리. renderHeader 형식이라 연동된 세션이 이 메시지를 요청으로 주입하지 않는다.
const NO_SESSION = "-"
const RELOGIN_HINT = "대화형 claude 세션에서 /microsoft-teams:connect 로 로그인하세요"

export type SendArgs = { chatId: string; file: string | null; title: string | null }

export function parseSendArgs(argv: string[]): SendArgs | { error: string } {
  const values: Record<string, string> = {}
  for (let i = 0; i < argv.length; i += 2) {
    const flag = argv[i]
    if (!["--chat", "--file", "--title"].includes(flag)) return { error: `알 수 없는 인자: ${flag}` }
    if (!argv[i + 1]) return { error: `${flag} 에 값이 없습니다` }
    values[flag] = argv[i + 1]
  }
  const chat = values["--chat"]
  if (!chat) return { error: "--chat 이 필요합니다" }
  if (CHAT_ID_RE.exec(chat)?.[1] !== chat) return { error: `chatId 형식이 아닙니다: ${chat}` }
  const file = values["--file"]
  return { chatId: chat, file: file && file !== "-" ? file : null, title: values["--title"] ?? null }
}

export type SendPorts = { auth: Pick<Auth, "hasRefreshToken">; graph: Pick<Graph, "sendMessage"> }
export type SendResult = { code: number; message: string }

export async function sendToChat(
  deps: SendPorts & { cfg: Config; configPath: string },
  req: { chatId: string; title: string; text: string },
): Promise<SendResult> {
  if (!req.text.trim()) return { code: SEND_EXIT.usage, message: "보낼 내용이 비어 있습니다" }
  if (!azureConfigured(deps.cfg)) {
    return { code: SEND_EXIT.login, message: `${deps.configPath} 의 azure.tenantId / clientId / scopes 가 비어 있습니다` }
  }
  if (!deps.auth.hasRefreshToken()) return { code: SEND_EXIT.login, message: `로그인 토큰이 없습니다. ${RELOGIN_HINT}` }
  const html = renderOutbound(req.title, NO_SESSION, req.text)
  const mid = await deps.graph.sendMessage(req.chatId, html)
  if (!mid) {
    // refresh 가 invalid_grant 를 받으면 Auth 가 토큰을 폐기로 표시한다.
    if (!deps.auth.hasRefreshToken()) return { code: SEND_EXIT.login, message: `로그인이 만료되었습니다. ${RELOGIN_HINT}` }
    return { code: SEND_EXIT.failed, message: "발신 실패 — 네트워크 · chatId · ChatMessage.Send 권한을 확인하세요 (원인은 위 오류 로그)" }
  }
  const truncated = html !== renderOutbound(req.title, NO_SESSION, req.text, Infinity)
  return { code: SEND_EXIT.ok, message: `전송됨 (message id ${mid})${truncated ? " — Teams 메시지 크기 제한으로 본문 뒷부분을 잘랐습니다" : ""}` }
}

// 플러그인 로그(makeLogger)는 시작할 때 파일을 비우므로 실행 중인 세션의 로그를 지우지 않도록 쓰지 않는다.
const stderrLog: Log = {
  info: () => {},
  debug: () => {},
  warn: (m, d) => process.stderr.write(`[claudecode-ms-teams] ${m}${d ? ` ${JSON.stringify(d)}` : ""}\n`),
  error: (m, d) => process.stderr.write(`[claudecode-ms-teams] ${m}${d ? ` ${JSON.stringify(d)}` : ""}\n`),
}

async function readText(file: string | null): Promise<string | { error: string }> {
  if (file) {
    try {
      return fs.readFileSync(file, "utf8")
    } catch (e) {
      return { error: `파일을 읽을 수 없습니다: ${(e as Error).message}` }
    }
  }
  if (process.stdin.isTTY) return { error: "--file 이나 stdin 으로 보낼 내용을 주세요" }
  // readFileSync(0) 은 non-blocking 파이프에서 EAGAIN 을 던질 수 있어 스트림으로 읽는다.
  return new Promise((resolve) => {
    let data = ""
    process.stdin.setEncoding("utf8")
    process.stdin.on("data", (c) => (data += c))
    process.stdin.on("end", () => resolve(data))
    process.stdin.on("error", (e) => resolve({ error: `stdin 을 읽을 수 없습니다: ${e.message}` }))
  })
}

export async function runSend(argv: string[]): Promise<number> {
  const args = parseSendArgs(argv)
  if ("error" in args) {
    console.error(`✗ ${args.error}\n\n${SEND_USAGE}`)
    return SEND_EXIT.usage
  }
  const text = await readText(args.file)
  if (typeof text !== "string") {
    console.error(`✗ ${text.error}`)
    return SEND_EXIT.usage
  }
  const cfg = loadConfig()
  const auth = new Auth(cfg, stderrLog)
  // Graph 의 429 재시도 대기(sleep)는 unref 타이머라 그동안 프로세스가 끝나 버린다 (결과 없이 종료 코드 0).
  // 발신이 끝날 때까지 이벤트 루프를 붙잡아 둔다.
  const keepAlive = setInterval(() => {}, 60_000)
  let result: SendResult
  try {
    result = await sendToChat(
      { cfg, configPath: configFilePath(), auth, graph: new Graph(auth, stderrLog) },
      { chatId: args.chatId, title: args.title ?? path.basename(process.cwd()), text },
    )
  } finally {
    clearInterval(keepAlive)
    auth.dispose()
  }
  if (result.code === SEND_EXIT.ok) console.log(`✓ ${result.message}`)
  else console.error(`✗ ${result.message}`)
  return result.code
}
