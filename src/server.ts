#!/usr/bin/env node
// Claude Code 채널 MCP 서버 진입점. Claude Code 가 세션마다 stdio 로 하나씩 띄운다.
// Lazy connect: connect / select_chat 도구가 호출되기 전에는 외부 통신을 하지 않는다.
import * as fs from "node:fs"
import { MicrosoftTeamsChannel, CHANNEL_INSTRUCTIONS } from "./channel.js"
import { Auth, configFilePath, Graph, loadConfig, makeLogger, PLUGIN_NAME } from "./core.js"
import { removeRegistry, startIpcServer, sweepStaleRegistry, writeRegistry, type RegistryEntry } from "./ipc.js"
import { McpServer } from "./mcp.js"

function packageVersion(): string {
  try {
    return JSON.parse(fs.readFileSync(new URL("../package.json", import.meta.url), "utf8")).version ?? "0.0.0"
  } catch {
    return "0.0.0"
  }
}

async function main(): Promise<void> {
  const cfg = loadConfig()
  const log = makeLogger(cfg, `[${process.pid}]`)
  const claudePid = process.ppid
  // 비동기 처리 하나의 실패가 세션 내내 MCP 서버를 죽이지 않도록 로그만 남긴다.
  process.on("unhandledRejection", (e) => log.error("unhandledRejection", { error: String(e) }))
  const background = (label: string, task: Promise<unknown> | undefined) =>
    void task?.catch((e) => log.error(label, { error: String(e) }))

  let channel: MicrosoftTeamsChannel | null = null
  const auth = new Auth(cfg, log, {
    onLoginComplete: () => background("loginComplete.crash", channel?.onLoginComplete()),
    onLoginFailed: (reason) => channel?.onLoginFailed(reason),
    onReauthRequired: () => channel?.onReauthRequired(),
  })
  const graph = new Graph(auth, log)

  let registry: RegistryEntry | null = null
  const setActive = (active: boolean) => {
    if (!registry) return
    registry = { ...registry, active }
    try {
      writeRegistry(registry)
    } catch (e) {
      log.error("registry.write.failed", { error: String(e) })
    }
  }

  // channel 과 mcp 가 서로를 참조하므로 notify 는 mcp 생성 이후에만 호출되는 closure 로 넘긴다.
  let mcp: McpServer | null = null
  const ch = new MicrosoftTeamsChannel({
    cfg,
    log,
    auth,
    graph,
    configPath: configFilePath(),
    reloadConfig: loadConfig,
    sessionId: process.env.CLAUDE_CODE_SESSION_ID,
    push: (content, meta) => mcp?.notify("notifications/claude/channel", { content, meta }),
    onActiveChange: setActive,
  })
  channel = ch

  const server = new McpServer({
    name: PLUGIN_NAME,
    version: packageVersion(),
    instructions: CHANNEL_INSTRUCTIONS,
    // 권한·질문 응답은 채널 permission relay 가 아니라 PermissionRequest hook 으로 처리한다
    // (relay 는 AskUserQuestion 답과 always 승인을 실을 수 없고 채널 모드에서만 동작한다).
    capabilities: { tools: {}, experimental: { "claude/channel": {} } },
    tools: ch.tools(),
    input: process.stdin,
    output: process.stdout,
    log,
  })
  mcp = server

  const shutdown = () => {
    auth.dispose()
    if (registry) removeRegistry(registry.claudePid, registry.serverPid)
    process.exit(0)
  }
  server.onClose(shutdown)
  process.on("SIGTERM", shutdown)
  process.on("SIGINT", shutdown)
  process.on("SIGHUP", shutdown)
  server.start()

  try {
    sweepStaleRegistry()
    const ipc = await startIpcServer(
      {
        stop: (body) => background("stop.crash", ch.handleStop(body)),
        // 결정이 날 때까지 hook 의 요청을 잡아둔다. 실패는 결정 없음으로 돌려 터미널 다이얼로그에 맡긴다.
        permission: (body, signal) =>
          ch
            .handlePermission(body, signal)
            .catch((e) => {
              log.error("permission.crash", { error: String(e) })
              return null
            })
            .then((decision) => ({ decision })),
        answered: (body) => ch.handleAnswered(body),
        "subagent-stop": (body) => ch.handleSubagentStop(body),
      },
      log,
    )
    registry = { serverPid: process.pid, claudePid, port: ipc.port, token: ipc.token, active: false }
    writeRegistry(registry)
    log.info("server.started", { claudePid, port: ipc.port })
  } catch (e) {
    // IPC 실패 시 Teams 발신만 비활성화되고 도구·채널 수신은 계속 동작한다.
    log.error("ipc.start.failed", { error: String(e) })
  }
}

main().catch((e) => {
  process.stderr.write(`[${PLUGIN_NAME}] 서버 시작 실패: ${String(e)}\n`)
  process.exit(1)
})
