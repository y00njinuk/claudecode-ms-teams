// node --import 로 불러 globalThis.fetch 를 바꾼다 (네트워크 없이 Graph 응답을 흉내낸다).
// 첫 발신은 429 (Retry-After 1초), 재시도는 201 로 응답한다.
let calls = 0
globalThis.fetch = async () => {
  calls++
  if (calls === 1) return new Response("{}", { status: 429, headers: { "Retry-After": "1" } })
  return new Response(JSON.stringify({ id: "1700000000042" }), { status: 201, headers: { "Content-Type": "application/json" } })
}
