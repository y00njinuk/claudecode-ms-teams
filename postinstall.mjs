#!/usr/bin/env node
// 글로벌 npm 설치 시 Claude Code 에 플러그인을 자동 등록한다.
// 실행 조건: npm_config_global=true 또는 CLAUDECODE_MS_TEAMS_AUTO_INSTALL=1. CI 에서는 opt-in 없이 건너뛴다.
// 실패해도 npm install 자체는 성공시키고, 사용자는 `claudecode-ms-teams install` 로 재시도한다.
import { dirname } from "node:path"
import { fileURLToPath } from "node:url"

const HERE = dirname(fileURLToPath(import.meta.url))
const TAG = "[claudecode-ms-teams] postinstall:"

function shouldRun() {
  const optIn = process.env.CLAUDECODE_MS_TEAMS_AUTO_INSTALL === "1"
  const isCI = process.env.CI === "true" || process.env.CONTINUOUS_INTEGRATION === "true"
  if (isCI && !optIn) {
    console.log(`${TAG} CI 감지 — 건너뜀`)
    return false
  }
  if (process.env.npm_config_global !== "true" && !optIn) {
    console.log(`${TAG} 로컬 install — 건너뜀 (등록하려면 \`claudecode-ms-teams install\`)`)
    return false
  }
  return true
}

async function main() {
  if (!shouldRun()) return
  try {
    const { install } = await import("./dist/install.js")
    const r = install({ packageDir: HERE })
    for (const s of r.steps) console.log(`${TAG} ${s}`)
    console.log(`${TAG} 설정 파일: ${r.configPath}`)
    console.log(`${TAG} 실행: ${r.launchCommand}`)
    console.log(`${TAG} 진단: claudecode-ms-teams doctor`)
  } catch (err) {
    console.warn(`${TAG} ${err?.message ?? err}`)
    console.warn(`${TAG} 등록 미완료 — \`claudecode-ms-teams install\` 로 재시도`)
  }
}

main()
