// 배포본은 Bun 이 아니라 Node 로 실행된다. 런타임마다 동작이 다른 부분은 빌드 결과를 Node 로 돌려 확인한다.
import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { spawnSync } from "node:child_process"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"

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
})
