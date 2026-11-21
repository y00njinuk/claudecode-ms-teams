// node-runtime.test.ts 가 Node 로 실행한다: 응답을 기다리던 클라이언트가 끊기면 핸들러 signal 이 abort 되는지 출력한다.
import http from "node:http"
import { pathToFileURL } from "node:url"

const { startIpcServer, TOKEN_HEADER } = await import(pathToFileURL(`${process.argv[2]}/ipc.js`).href)
const log = { info() {}, warn() {}, error() {}, debug() {} }
const state = { abortedBeforeDisconnect: false, abortedAfterDisconnect: false }
let aborted = false
const ipc = await startIpcServer(
  {
    permission: (_body, signal) => {
      signal.addEventListener("abort", () => (aborted = true))
      return new Promise(() => {})
    },
  },
  log,
)
const req = http.request({
  host: "127.0.0.1",
  port: ipc.port,
  path: "/permission",
  method: "POST",
  headers: { "Content-Type": "application/json", [TOKEN_HEADER]: ipc.token },
})
req.on("error", () => {})
req.end("{}")
await new Promise((r) => setTimeout(r, 200))
state.abortedBeforeDisconnect = aborted
req.destroy()
await new Promise((r) => setTimeout(r, 300))
state.abortedAfterDisconnect = aborted
console.log(JSON.stringify(state))
process.exit(0)
