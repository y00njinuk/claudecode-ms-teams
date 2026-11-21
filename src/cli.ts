#!/usr/bin/env node
// claudecode-ms-teams install | uninstall [--purge] | doctor | send
import * as path from "node:path"
import { fileURLToPath } from "node:url"
import { doctor, install, uninstall } from "./install.js"
import { runSend, SEND_USAGE } from "./send.js"

const PACKAGE_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..")

const HELP = `claudecode-ms-teams — Claude Code ↔ Microsoft Teams 채널 플러그인

사용법:
  claudecode-ms-teams install            Claude Code 에 플러그인 등록 (npm 글로벌 설치 시 자동 실행)
  claudecode-ms-teams uninstall [--purge] 플러그인 등록 해제 (--purge: 설정·토큰도 삭제)
  claudecode-ms-teams doctor             설치 상태 진단

${SEND_USAGE}
`

async function main(argv: string[]): Promise<number> {
  const cmd = argv[0] ?? "help"
  if (cmd === "install") {
    const r = install({ packageDir: PACKAGE_DIR })
    for (const s of r.steps) console.log(`✓ ${s}`)
    console.log(`✓ 설정 파일${r.configCreated ? " 생성" : ""}: ${r.configPath}`)
    console.log("")
    console.log("다음 단계:")
    console.log(`  1. ${r.configPath} 의 azure.tenantId / clientId / scopes 를 채웁니다.`)
    console.log(`  2. Claude Code 를 실행합니다. Teams 메시지를 세션 요청으로 받으려면 채널 모드로: ${r.launchCommand}`)
    console.log("  3. 세션에서 /microsoft-teams:connect 를 실행합니다.")
    return 0
  }
  if (cmd === "uninstall") {
    const steps = uninstall({ purge: argv.includes("--purge") })
    if (!steps.length) console.log("등록된 항목이 없습니다.")
    for (const s of steps) console.log(`✓ ${s}`)
    return 0
  }
  if (cmd === "doctor") {
    const checks = doctor({ packageDir: PACKAGE_DIR })
    for (const c of checks) console.log(`${c.ok ? "✓" : "✗"} ${c.label}: ${c.detail}`)
    return checks.every((c) => c.ok) ? 0 : 1
  }
  if (cmd === "send") return runSend(argv.slice(1))
  console.log(HELP)
  return cmd === "help" || cmd === "--help" || cmd === "-h" ? 0 : 1
}

// process.exit 는 파이프로 나가는 출력을 자를 수 있어 exitCode 만 정한다.
main(process.argv.slice(2)).then(
  (code) => {
    process.exitCode = code
  },
  (e) => {
    console.error(`✗ ${(e as Error).message}`)
    process.exitCode = 1
  },
)
