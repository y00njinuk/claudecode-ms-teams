# ARCHITECTURE.md — claudecode-ms-teams

## 시스템 개요

Claude Code 세션과 Microsoft Teams 채팅을 양방향으로 잇는 채널 플러그인.
opencode-ms-teams 의 Auth · Graph · 트리거 · 렌더링 코어를 그대로 쓰고, OpenCode 플러그인 훅 자리를
Claude Code 의 세 가지 확장점(채널 MCP 서버 · command hook · skill)으로 바꿨다.

```
                       ┌──────────────── claude (PID = 레지스트리 키) ────────────────┐
                       │                                                              │
 Stop · StopFailure    │                                                              │
 SubagentStop          │                                                              │
 PermissionRequest ──► hook.js ──HTTP 127.0.0.1 + token──► server.js (MCP stdio, 세션당 1개)
 PostToolUse(Ask..)    │   ▲ 결정(JSON)을 기다렸다가 stdout    │   ▲                     │
   (command hook)      │   └──────────────────────────────────┘   │ tools/call          │
                       │      notifications/claude/channel ◄─ server.js               │
                       └──────────────────────────────────────────┼───────────────────┘
                                                                  │
                                         Microsoft Graph (delegated, device code)
                                         POST /chats/{id}/messages · GET 폴링
```

런타임 의존성 제로 — Node 내장 `fetch` · `http` · `node:` 빌트인만 쓴다. MCP 도 최소 JSON-RPC 를 직접 구현한다 (`src/mcp.ts`).

## 컴포넌트

| 파일 | 역할 |
|---|---|
| `src/core.ts` | 설정 · 로그 · Auth(device code / refresh) · Graph · 트리거 정규화 · Teams HTML 렌더링 |
| `src/mcp.ts` | newline-delimited JSON-RPC MCP 서버 (initialize · ping · tools/list · tools/call · notification) |
| `src/channel.ts` | 세션 상태 머신: 도구(connect / select_chat / status / disconnect), Stop 발신, 폴링, 권한·질문 대기와 답장 처리 |
| `src/server.ts` | MCP 서버 진입점. channel · Auth · IPC 를 조립하고 종료 시 레지스트리를 정리 |
| `src/ipc.ts` | hook → 서버 localhost IPC (즉시 202 / 발신 · 결정까지 대기) 와 `run/<claudePid>.json` 레지스트리 |
| `src/hook.ts` | command hook 클라이언트. 항상 exit 0, stdout 은 PermissionRequest 결정에만 쓴다 |
| `src/install.ts` · `src/cli.ts` · `postinstall.mjs` | 로컬 마켓플레이스 등록 · 플러그인 설치 · 진단 |
| `src/send.ts` | 세션 없이 Teams 채팅방으로 한 건 보내는 `send` CLI (예약 작업 보고용) |
| `src/todo.ts` | 선택 기능 Microsoft To Do 도구 (`todo_*`) · 날짜 변환 · 목록 해석 · Graph 오류 안내 |
| `.claude-plugin/plugin.json` | 플러그인 매니페스트 (`mcpServers` 인라인) |
| `.claude-plugin/marketplace.json` | 패키지 루트 자체를 마켓플레이스로 노출 (`source: "./"`) |
| `hooks/hooks.json` | `Stop` · `StopFailure` → `hook.js stop`, `SubagentStop` → `hook.js subagent-stop`, `PermissionRequest(*)` → `hook.js permission`, `PostToolUse(AskUserQuestion)` → `hook.js answered` |
| `skills/connect/SKILL.md` | `/microsoft-teams:connect` |

## OpenCode 판과의 대응

| 기능 | opencode-ms-teams | claudecode-ms-teams |
|---|---|---|
| 답신 → 세션 | `pollLoop` → `promptAsync` | `notifications/claude/channel` push |
| idle 발신 | `session.idle` + `lastAssistantContext()` | `Stop` hook 의 `last_assistant_message` |
| 권한 알림·응답 | `permission.asked` + `postSessionIdPermissionsPermissionId` | `PermissionRequest` hook 이 Teams 답장을 기다렸다가 `decision` 출력 |
| 질문 알림·응답 | `question.asked` + `/question/{id}/reply` | 같은 `PermissionRequest` hook (`tool_name=AskUserQuestion`) 이 `updatedInput.answers` 로 응답 |
| TUI 선응답 정리 | `permission.replied` / `question.replied` 이벤트 | `PostToolUse(AskUserQuestion)` · `Stop` / `StopFailure` · `SubagentStop` · hook 연결 종료 |
| 연동 시작 | `/connect-teams` + marker 감지 | `/microsoft-teams:connect` skill → `connect` 도구 |
| device code 안내 | 응답 텍스트에 append + `stripAuthNotice` | 도구 결과 / 로그인 완료는 `kind="system"` 채널 메시지 |
| sub session gate | `parentID` + `sessionParent` 캐시 | `Stop` 은 main 에서만 발화 (subagent 는 `SubagentStop`), sub-agent 의 권한·질문 알림에는 `agent_type` 배너 |
| model/agent 재사용 | 필요 | 불필요 (같은 세션에 직접 주입) |

## 세션 식별 (hook ↔ 서버)

Claude Code 는 stdio MCP 서버와 command hook 을 모두 `claude` 프로세스의 **직계 자식**으로 띄운다 (2.1.319 실측).

- 서버: `process.ppid` = claude PID → `run/<claudePid>.json` 에 `{serverPid, port, token, active}` 기록 (0600, 원자적 rename)
- hook: `CLAUDE_PID` 환경변수 → `process.ppid` 순서로 레지스트리 조회
- `session_id` 는 키로 쓰지 않는다. `/clear` 에서 바뀌지만 MCP 서버 프로세스는 그대로이기 때문이다.
- `active=false` (채팅방 미선택) 면 hook 은 HTTP 호출 없이 끝난다. 매 턴 발화하는 Stop hook 의 비용을 줄이기 위함.
- 서버는 `answered` · `subagent-stop` 요청에 즉시 202 로 응답한다. `stop` 은 Teams 발신이 끝날 때까지 (최대 `STOP_WAIT_MS` = 8초),
  `permission` 은 결정이 날 때까지 응답을 잡아둔다 (아래 "발신 파이프라인", "권한 · 질문 응답").
- 기동 시 죽은 서버의 레지스트리를 sweep 하고, 종료(stdin EOF / SIGTERM / SIGINT / SIGHUP) 시 자기 파일만 지운다.
  Claude Code 는 종료할 때 stdin 을 먼저 닫지 않고 SIGINT → (100ms) SIGTERM → (400ms) SIGKILL 을 보낸다 (2.1.330 `-p` 실측).

## Lazy connect

플러그인을 설치하면 모든 세션에서 MCP 서버가 뜨지만, `connect` / `select_chat` / `todo_*` 도구가 호출되기 전에는
Azure · Graph 로 어떤 요청도 보내지 않는다. 인증 이벤트(재로그인 안내, 로그인 완료)도 connect 이후에만 채널로 보낸다.
`send` CLI 는 사용자가 직접 실행하는 별도 프로세스라 이 규칙의 대상이 아니다.

## 인증 보호

device code 는 그 코드를 먼저 입력한 사람의 계정으로 로그인시킨다. 코드가 Teams 채팅방으로 새면 채팅방 멤버가
대신 로그인해 세션을 조작(메시지 주입, `#allow`)할 수 있으므로 다음을 지킨다.

- device flow 는 `connect` 도구에서만 시작한다. refresh 가 `invalid_grant` 를 받으면 `revoked` 로 표시하고
  토큰 요청을 멈춘 뒤, "터미널에서 `/microsoft-teams:connect` 를 실행하라" 는 `kind="system"` 안내를 한 번만 보낸다.
  자동 재시작하면 무인 상태에서 코드가 응답에 실려 나가고, 15분마다 새 턴이 깨어난다.
- device flow 진행 중에는 Stop 발신 텍스트와 권한·질문 알림에서 user code 를 `[로그인 코드 숨김]` 으로 가린다.
- device flow 로그인이 끝나면 `/me` 를 다시 조회한다. 기존 계정과 id 가 다르면 연동을 해제하고 발신자 게이트를 비운다.
- `tokens.json` 은 여러 세션의 서버가 공유하므로 tmp + rename 으로 원자적으로 쓴다. refresh 응답에 `refresh_token` 이 없으면 기존 값을 유지한다.
- `connect` 의 `relogin: true` 는 저장된 로그인이 있어도 device flow 를 시작한다. 새 토큰이 저장될 때까지 기존 토큰을 그대로 쓰므로 Teams 는 계속 동작하고,
  시작 전에 `/me` 를 확인해 두어 로그인 후 계정이 바뀌었으면 위와 같이 연동을 해제한다 (확인에 실패하면 relogin 을 시작하지 않는다). 코드 가림도 같다.
  skill 은 터미널 사용자가 요청할 때만 relogin 하도록 안내한다 (Teams 메시지만으로는 하지 않는다).

## 장애 격리

- `Graph.call()` 은 JSON 이 아닌 본문(프록시·캡티브 포털 HTML, 본문 중간 끊김)을 `null` 로 흡수한다.
  원인이 필요한 To Do 도구는 같은 재시도 규칙(401 → refresh 후 1회, 429 → `Retry-After` 후 1회)을 쓰는 `Graph.request()` 로
  `{ ok, status, body, error }` (`auth` · `network` · `ssl` · `badBody` · `http`) 를 받는다. `call()` 은 그 결과를 `null` 로 축약할 뿐이다.
- `server.ts` 의 fire-and-forget 호출은 모두 `.catch` 로 감싸고, `unhandledRejection` 은 로그만 남긴다. 그러지 않으면 한 번의 실패로 세션 내내 MCP 서버가 사라진다.
- `pollOnce` 는 메시지 단위로 예외를 격리한다. 커서가 이미 배치 끝으로 전진했으므로 하나의 실패가 나머지 메시지를 잃게 해서는 안 된다.

## 발신 파이프라인 (Stop)

```
Stop hook → POST /stop {session_id, cwd, text}
  └─ channel.handleStop
       ├─ 채팅방 미선택 / 빈 텍스트 → skip
       ├─ device flow 진행 중이면 user code 가림
       ├─ 중복 가드 (직전 발신과 같은 텍스트가 10초 안에 다시 오면 skip)
       ├─ renderOutbound(<cwd basename>, shortSessionId(session), text)   → "[<제목> · <세션 8자>]" 헤더
       └─ Graph POST → sentIds.add(mid), lastActivity 갱신
```

Stop hook 은 발신이 끝날 때까지 기다린다. 서버가 `holdAtMost(handleStop, STOP_WAIT_MS)` 로 응답을 최대 8초 잡아두고, hook 의 IPC 대기(`ipcTimeoutMs("stop")`)는
그보다 1초 길며 hooks.json 의 Stop timeout(10초)보다 짧다 (테스트로 강제). Claude Code 2.1.330 `-p` 실측으로 Stop hook 이 실행되는 동안에는 MCP 서버가 살아 있고,
hook 이 끝나면 곧바로 서버에 SIGINT 를 보내 0.5초 안에 SIGKILL 한다. 그래서 202 로 먼저 응답하면 `claude -p` 처럼 턴 직후 종료하는 실행에서 마지막 응답이 발신 전에 사라진다.
8초가 지나면 응답만 돌려주고 발신은 취소하지 않는다 (429 `Retry-After` 대기 등). 채팅방이 연동된 세션은 턴 종료 처리가 발신 시간만큼 늦어진다.

Stop 은 턴마다 한 번 발화하므로 OpenCode 판의 무기한 fingerprint 가드는 쓰지 않는다 (같은 답을 반복하는 정상 대화가 막힌다).
커서는 발신 시점에 전진시키지 않는다. 자기 발신은 `sentIds` 로 거른다.

플러그인이 보내는 모든 메시지(Stop 응답 · 권한/질문 알림 · 확인 메시지)는 `renderHeader()` 의 `[<제목> · <세션 8자 또는 ->]` 헤더로 시작한다.
수신 쪽 `isPluginMessage()` 가 이 헤더로 플러그인 발신을 알아본다 (아래 "수신 파이프라인").

렌더러(`mdToHtml`)는 코드 블록·표를 placeholder 로 빼둔 뒤 나머지 텍스트만 escape 하고 되돌린다.
렌더 결과처럼 생긴 원문 HTML 을 분할 기준으로 쓰면 escape 없이 통과한다.

## 세션 없는 발신 (`send` CLI)

`claudecode-ms-teams send --chat <19:...> [--file <경로>] [--title <제목>]` 은 Claude Code 세션 없이 Markdown 한 건을 보낸다.
예약 작업이 `claude -p` 결과를 보고하는 용도다. `select_chat` · Stop hook 을 거치지 않으므로 세션 안에서 도구를 허용하거나 발신 완료를 로그로 확인할 필요가 없고,
보고 세션이 채팅방을 연동하지 않으니 권한 요청이 Teams 답장을 기다리며 멈추지도 않는다.

- 저장된 `tokens.json` 만 쓴다. 만료됐으면 서버와 같은 `Auth` 로 refresh 해 원자적으로 저장하지만 device flow 는 시작하지 않는다 (무인 출력으로 코드가 새면 다른 사람이 로그인할 수 있다).
- 메시지는 `[<제목> · -]` 헤더로 시작한다 (기본 제목은 현재 디렉토리 이름). 같은 채팅방에 연동된 세션이 이 메시지를 요청으로 주입하지 않는다 (`isPluginMessage`).
- 진단은 stderr 로만 쓴다. 플러그인 로그(`makeLogger`)는 시작할 때 파일을 비우므로 실행 중인 세션의 로그를 지우지 않도록 쓰지 않는다.
- 종료 코드: 0 전송됨 · 1 발신 실패 (네트워크 · Graph 오류) · 2 사용법 오류 (인자 · 빈 본문 · 파일 읽기) · 3 설정 · 로그인 필요 (azure 미설정 · 토큰 없음 · refresh `invalid_grant`).
- chatId 는 인자로만 받는다. 설정 파일에 기본 채팅방을 두지 않는다 (chatId 는 세션 메모리 전용이라는 원칙과 설정 단일 소스 원칙을 유지한다).

## 수신 파이프라인 (폴링)

`select_chat` 에서 `latestMessageId()` 로 커서를 seed 하고 폴링을 시작한다. seed 가 `null` (조회 실패) 이면 연동하지 않는다.
`delayFor(now - lastActivity)` 간격 (기본 5s → 15s → 30s) 으로 `messagesSince(cursor)` 를 호출한다.

키워드 트리거 없이 채팅방의 모든 메시지를 받는다. 각 메시지 처리 순서:

1. 커서를 배치의 최대 id 로 먼저 전진 (처리 결과와 무관)
2. `sentIds` 에 있으면 skip (자기 발신)
3. 발신자 게이트: 로그인 본인(`/me`) 또는 `access.allowedUserIds`
4. `isPluginMessage()` — 플러그인 헤더로 시작하면 skip
5. 답장 verb (`#allow` / `#allow-always` / `#allow-reset` / `#reject` / `#answer` / `#skip`) → `executeReply` (세션에 절대 주입하지 않음)
6. 이 메시지보다 먼저 알린 main 질문이 정확히 하나면 그 질문의 답으로 처리 (`#skp` 처럼 `#영문`으로 시작하면 제외)
7. 그 외 → 예전 호출 키워드(`#claude` · `가재야` 등)가 있으면 떼고 `notifications/claude/channel` (`meta: message_id, sender, chat`)

매칭(트리거 · 헤더 · 선택지 label)은 `normalizeForTrigger()` (NFKC) 결과로 하지만, 세션에 넘기는 본문과 `#answer` 값은 `cleanInbound()` 한
원문 문자를 그대로 쓴다 (`payloadVerbatim()`). NFKC 는 `ㅇㅇ` 을 조합형 자모로, `①` 을 `1` 로 바꾼다.

4 가 없으면 루프가 생긴다. delegated 권한이라 모든 세션의 발신이 본인 계정으로 올라오므로, 같은 채팅방에 연동된 다른 세션(다른 머신 ·
opencode-ms-teams 포함)의 응답이 본인 메시지로 보인다. 세션 A 의 응답이 B 에 주입되고 B 의 응답이 다시 A 에 주입된다. 자기 세션도 `sendMessage` 응답이
오기 전에 폴링이 그 메시지를 읽으면 `sentIds` 로 거를 수 없다. 헤더는 Teams 가 이모지를 `<emoji>` 태그로 바꿔 돌려줘도 남도록 텍스트 형식만 본다.
제목은 NFKC 후 대괄호 · 개행을 빼고 (전각 `］` 이 수신 쪽 NFKC 에서 `]` 가 되어 헤더가 깨지는 것 방지), `htmlToText` 는 숫자 엔티티(`&#183;`)도 푼다.
세션 자리는 영문자를 하나 이상 포함한 hex 8자로 제한한다 (`[회의록 · 20261120]` 같은 사용자 메시지와 겹치지 않게. 숫자만 나온 UUID 앞자리는 `-` 로 쓴다).
이 판별은 실제 Teams 가 `[` · `·` · `]` 를 텍스트로 돌려준다는 전제에 기댄다 (같은 헤더를 쓰는 opencode-ms-teams 에서 쓰던 형식이다).

6 은 main 질문이 떠 있는 동안 그 턴이 멈춰 있기 때문이다. 이때 온 채널 메시지는 턴이 끝날 때까지 큐에 쌓이므로 (아래) 새 요청으로 넣어도 처리되지 않는다.
질문 알림보다 먼저 보낸 메시지는 질문을 보기 전에 쓴 새 요청이므로 답으로 쓰지 않는다.
sub-agent 의 질문은 main 턴을 막지 않을 수 있어 `#answer` 로만 받는다.
남는 빈틈: 터미널에서 Esc 로 질문을 닫으면 PostToolUse · Stop 이 발화하지 않아 질문이 대기로 남고, 다음 Teams 메시지 하나(질문 묶음이면 질문 수만큼)가
그 질문의 답으로 쓰인다. 확인 메시지(`✅ 전달됨 answer …` / `📝 답변 기록됨 … (건너뛰기: #skip)`)로 드러나므로 사용자가 `#skip` 하거나 다시 보낸다.

Claude Code 는 바쁠 때 도착한 채널 메시지를 큐에 쌓았다가 다음 턴에 함께 전달하므로, OpenCode 판의 `waiting` 게이트가 필요 없다.

## 권한 · 질문 응답 (PermissionRequest hook)

Claude Code 는 권한 다이얼로그를 띄울 때 `PermissionRequest` hook 을 실행한다. 다음은 Claude Code 2.1.322 에서 실측한 동작이며,
공식 hooks 문서에는 입력·출력 형식만 있고 다이얼로그와의 경합은 적혀 있지 않다.

- hook 이 실행 중인 동안에도 터미널 다이얼로그는 떠 있고, **먼저 온 응답이 적용된다** (hook 이 결정을 출력하면 다이얼로그가 닫힌다).
- `AskUserQuestion` · `ExitPlanMode` 에도 발화한다. auto 모드에서도 질문 다이얼로그에는 발화한다.
- sub-agent (백그라운드 포함) 의 권한 요청에는 입력에 `agent_id` · `agent_type` 이 실려 오고, 그 agent 가 끝나면 같은 `agent_id` 로 `SubagentStop` 이 발화한다.
  백그라운드 agent 의 요청은 main 의 `Stop` 뒤에도 열려 있을 수 있다.
- 터미널이 먼저 응답해도 Claude Code 는 hook 프로세스를 죽이지 않는다. 뒤늦은 출력은 무시된다.
- `decision.updatedInput.answers` (질문 문구 → 답 문자열, 복수 선택은 `", "` 로 연결) 로 질문에 답할 수 있다. 선택지에 없는 문자열도 받는다.
- `decision.updatedPermissions` 에 입력의 `permission_suggestions` 를 그대로 돌려주면 터미널의 `Yes, and …` 선택과 같은 효과가 난다
  (규칙 추가 · 디렉토리 추가 · 모드 전환).
- `ExitPlanMode` 는 `allow` + 원래 입력을 그대로 담은 `updatedInput` 으로 승인된다 (`allow` 단독은 확인하지 않았다).

```
PermissionRequest hook → POST /permission {session_id, cwd, tool_name, tool_input, permission_suggestions, agent_id, agent_type}
  └─ channel.handlePermission            (응답을 잡아둔다)
       ├─ 채팅방 미선택 → {decision: null} 로 즉시 종료 (hook 은 active=false 면 요청조차 보내지 않는다)
       ├─ 자동 승인 도구 (질문 제외, 예외 경로 아님) → 알림 없이 즉시 allow
       ├─ 코드 발급: 질문 q + base36 5자 / 그 외 p + base36 5자, pending 등록 (alertId = null)
       ├─ Teams 알림 발신 → alertId = 알림 메시지 id (실패하면 pending 폐기 → 즉시 종료)
       └─ Teams 답장 → executeReply → settle(decision) → hook 이 stdout 으로 출력 → Claude Code 가 적용
```

| 답장 | 대상 | hook 결정 |
|---|---|---|
| `#allow [p…]` | 권한 | `{behavior: "allow"}` (`ExitPlanMode` 는 `updatedInput` 포함) |
| `#allow-always [p…]` | 권한 | `#allow` 와 같고, 제안이 있으면 `updatedPermissions: permission_suggestions` 추가. 그 도구를 세션 자동 승인에 등록 |
| `#reject [p…]` | 권한 | `{behavior: "deny", message}` |
| `[#answer [q…[-n]]] <답>` | 질문 | 모든 질문이 채워지면 `{behavior: "allow", updatedInput: {...tool_input, answers}}` |
| `#skip [q…]` | 질문 | `{behavior: "deny", message}` |
| `#allow-reset` | — | 결정 없음. 세션 자동 승인을 비운다 |

코드는 생략한다. 생략하면 **그 답장보다 먼저 알림이 나간** 같은 종류(권한 / 질문)의 대기 요청만 후보로 보고 (`alertId < 답장 id`,
Teams 메시지 id 는 시간순), 후보가 정확히 하나일 때만 적용한다. 답장을 쓰는 사이 새로 온 요청 — 예를 들어 늦게 도착하거나 재전송된 `#allow` 뒤에
생긴 `rm -rf` 요청 — 은 사용자가 보지 못했으므로 후보가 아니다. 후보가 여럿이면 적용하지 않고 `⚠️ 대상 지정 필요` 로 각 요청의 코드를 알린다
(가장 최근 · 가장 오래된 요청을 고르면 터미널에서 이미 답한 요청이 남아 있을 때 엉뚱한 요청을 승인할 수 있다).
예외로 코드 없는 `#allow-always` 는 후보가 모두 같은 도구면 어느 하나를 고르지 않고 그 도구만 자동 승인에 등록한다
(그 결과 예외 경로가 아닌 같은 도구 요청이 승인되고, 어느 요청의 `permission_suggestions` 도 저장하지 않는다).
`#answer` 의 첫 단어는 대기 중인 질문 코드일 때만 코드로 본다 (`python` · `pytest` 같은 답이 코드 형식과 겹친다).

답의 값은 선택지 label (대소문자 · 전각 표기 무시) → 선택지 번호 → 그 외는 입력 그대로 순으로 해석한다. `text:` 접두는 해석 없이 그대로 쓴다.
숫자만 온 답이 선택지 범위를 벗어나거나 답이 비면 `❌ 답변 미반영` 으로 이유를 알리고 질문을 계속 대기시킨다
(이 세션의 질문이므로 조용히 버리지 않는다. 모르는 코드 · 대기 요청 없음은 여전히 무시한다).
질문이 여러 개면 코드 없는 답은 아직 답하지 않은 첫 질문에 채우고 (sub-code 를 붙이면 그 질문), 부분 응답은 `📝 답변 기록됨` 으로 확인한다.

### 세션 자동 승인 (`#allow-always`)

`#allow-always` 는 그 요청의 도구 이름을 서버 메모리의 자동 승인 목록에 넣는다. 이후 같은 도구의 PermissionRequest 는 알림 없이 즉시 `allow`
(`ExitPlanMode` 는 `updatedInput` 포함) 를 돌려준다. 같은 도구로 이미 대기 중인 요청도 (예외 경로가 아니면) 그 자리에서 승인한다.

- 범위는 도구 단위다. Claude Code 의 `permission_suggestions` 는 명령 접두 단위라 매번 다른 명령을 쓰는 원격 세션에서는 계속 묻게 된다.
- 제외 (`autoAllowBlocker`): `AskUserQuestion` (답이 필요하다), workspace 밖 경로(`file_path` · `notebook_path` · `path`, Glob 은 `pattern` 도.
  `~` 로 시작하면 밖으로 본다), 설정 · git 경로(`.claude/` · `.git/` · `.mcp.json` — 고치면 hook · 권한 규칙 · MCP 설정으로 임의 명령 실행까지 이어진다),
  workspace(cwd) 를 모르는 경우. 예외로 다시 물을 때는 알림에 이유를 적는다.
- 도구 인자 전체를 검사하지는 않는다 (Bash 명령 내용, MCP 도구 인자). `Bash` 를 등록하면 모든 명령이 승인된다.
- 해제: `#allow-reset`, `disconnect`, 계정 변경(연동 해제), 다른 채팅방 선택, 서버 재시작. Stop 이나 `/clear` 로는 지우지 않는다.
- 자동 승인된 요청은 Teams 로 알리지 않는다 (도구 호출마다 메시지가 쌓인다). `status` 도구가 목록을 보여준다.
- 함께 저장한 Claude Code 규칙(`updatedPermissions`)은 Claude Code 설정에 남으므로 `#allow-reset` 으로 지워지지 않는다.

### pending 수명

Claude Code 는 터미널에서 먼저 응답했다는 사실을 알려주지 않으므로 다음 신호로 pending 을 폐기한다. 폐기하면 기다리던 hook 에
`{decision: null}` 을 돌려줘 프로세스를 끝내고, 그 코드로 온 답장은 무시한다.

| 신호 | 폐기 대상 |
|---|---|
| `PostToolUse(AskUserQuestion)` (`hook.js answered`) | 같은 질문 문구를 가진 질문 (PermissionRequest 입력에는 `tool_use_id` 가 없어 문구로 대조) |
| `Stop` · `StopFailure` (`hook.js stop`, 응답이 비어도 전달) | main 의 모든 pending (턴이 끝났으면 main 의 다이얼로그는 닫혔다). sub-agent (`agent_id` 있음) 의 pending 은 유지 |
| `SubagentStop` (`hook.js subagent-stop`) | 그 `agent_id` 의 pending |
| hook 연결 종료 (hook timeout · 세션 종료) | 그 hook 의 pending. `ipc.ts` 가 `AbortSignal` 로 알린다 |
| 24시간 TTL (`REPLY_TTL_MS` = hooks.json 의 PermissionRequest `timeout`, 테스트로 강제) | 만료된 pending |
| `disconnect` · 계정 변경 · 다른 채팅방 선택 | 전부 (이전 채팅방에 알린 요청을 새 채팅방의 코드 없는 답장이 승인하지 않게) |
| `#allow-always` | 같은 도구의 다른 대기 요청 (예외 경로 제외) — 폐기가 아니라 allow 로 settle |

남는 빈틈: 터미널에서 권한을 승인·거부한 뒤 **같은 턴(또는 그 sub-agent)이 끝나기 전에** Teams 답장이 오면 결정은 hook 으로 전달되지만 Claude Code 가 무시한다.
사용자가 턴을 중단(interrupt)하면 `Stop` 이 발화하지 않으므로 main 의 pending 은 다음 턴이 끝날 때까지 남는다.
그래서 확인 메시지는 "반영됨" 이 아니라 `✅ 전달됨` 으로 쓴다. 모든 도구에 PostToolUse hook 을 걸면 닫을 수 있지만 도구 호출마다
프로세스를 띄우는 비용이 커서 하지 않는다.

### 채널 permission relay 를 쓰지 않는 이유

`claude/channel/permission` capability 로도 권한 요청을 중계할 수 있지만 (channels-reference), verdict 가 `allow` / `deny` 뿐이라
질문의 답과 always 승인을 실을 수 없고 채널 모드에서만 동작한다. hook 과 함께 선언하면 같은 요청이 두 번 알림되므로 선언하지 않는다.
발신자 게이트(로그인 본인 + `access.allowedUserIds`)는 hook 결정에도 그대로 전제 조건이다.

`PreToolUse` 의 `permissionDecision: "allow"` + `updatedInput.answers` 도 문서화된 방법이지만 (hooks 문서), PreToolUse 는 다이얼로그가
뜨기 전에 끝나야 하므로 기다리는 동안 터미널에서 답할 수 없게 된다.

## Microsoft To Do (선택 기능)

같은 `Auth` · `Graph` · `tokens.json` 으로 To Do 를 다룬다. 로그인 한 번으로 Teams 와 함께 쓴다.

### 기능 게이트

MCP `tools/list` 는 서버 시작 시 정한 배열을 그대로 돌려주므로, 노출 여부도 **서버 시작 시 설정**(`azure.scopes`)으로 정한다 (`todoAccess()`).

| `azure.scopes` | 도구 | instructions |
|---|---|---|
| `Tasks.ReadWrite` | 7개 (`todo_list_lists` · `todo_list_tasks` · `todo_get_task` · `todo_create_task` · `todo_update_task` · `todo_delete_task` · `todo_checklist`) | To Do 지침 추가 (삭제는 확인 후) |
| `Tasks.Read` 만 | 조회 3개 | 조회만 가능 |
| 없음 | 없음 — 0.4.0 과 같은 `tools/list` · instructions | — |

scope 를 바꾸면 세션을 다시 시작해야 한다. `status` 는 시작 시 범위와 지금 설정이 다르면 재시작이 필요하다고 알린다.
예외로 `connect` 의 `relogin` 인자(scope 추가 후 재로그인 경로)는 To Do 설정과 무관하게 항상 있다. 그 외 Teams 도구 · instructions 는 0.4.0 과 같다.
`REQUIRED_SCOPES` (Teams 필수) 는 바꾸지 않는다. scope 비교는 대소문자와 `https://graph.microsoft.com/` 접두를 무시한다.

### 날짜 · 시간대

- 입력: `dueDate` = `YYYY-MM-DD`, `reminder` = `YYYY-MM-DDTHH:mm` (시각 생략 시 09:00), 빈 문자열 = 해제 (`reminderDateTime: null` + `isReminderOn: false`).
- 보낼 때는 `{ dateTime: <현지 시각>, timeZone: todo.timeZone }` 그대로 보낸다. 설정이 비어 있으면 시스템 IANA 시간대 (`Intl`).
  UTC 로 바꿔 보내지 않는 이유: To Do 는 기한(`dueDateTime`)을 날짜로만 쓰고 보낸 시간대 기준 날짜를 그대로 남긴다는 사례가 있다
  (Microsoft Q&A 1237380, 사용자 보고). UTC 로 바꾸면 UTC 보다 늦은 시간대에서 날짜가 하루 당겨진다.
- 응답은 UTC 로 온다. `todo.timeZone` 의 현지 시각으로 바꿔 보여준다 (기한은 날짜만). Windows 이름이면 변환 표가 없으므로 시스템 IANA 시간대로 바꾼다
  (UTC 그대로 날짜만 자르면 UTC 보다 이른 시간대에서 기한이 하루 이르게 보인다).
- Graph `dateTimeTimeZone` 문서는 Windows 시간대와 추가 시간대 목록(`Asia/Seoul` 포함)을 받는다고 적지만, 생성 · 수정 메서드가 모두 지원하지는 않을 수 있다고 한다.
  400 이 나면 `todo.timeZone` 을 Windows 이름으로 바꾸라고 안내한다.

### 토큰 scope 불일치

여러 세션이 `tokens.json` 을 공유하므로, scope 를 추가하기 전에 시작한 세션(또는 구버전)이 refresh 하면 Tasks 권한이 빠진 토큰이 저장될 수 있다.
To Do 도구는 호출할 때마다 저장된 토큰의 `scope` 를 본다 (`Auth.tokenScope()`). 필요한 권한이 없으면

1. 설정을 다시 읽어 `cfg` 를 in-place 갱신하고 (설정에 Tasks scope 가 없으면 세션 재시작 안내),
2. `Auth.refreshWithTodoScopes()` 로 전체 scope refresh 를 한 번 시도한 뒤,
3. 그래도 없으면 재로그인(`connect` relogin) 을 안내한다. 토큰 `scope` 를 모르면(필드 없음) 그대로 Graph 응답에 맡긴다.

### 동의 없는 scope 로 refresh

설정에 `Tasks.*` 를 넣었지만 Azure 앱에 동의가 없으면 refresh 가 `invalid_grant` + `suberror: consent_required` (`AADSTS65001`) 로 거절된다고 알려져 있다
(Microsoft Q&A 사례, 실측 아님 — `test/fixtures/aad-consent-required.json`). 허용되지 않은 scope 는 `70000` 으로 올 수도 있다.
기존 처리대로면 `revoked` 가 되어 Teams 까지 멈춘다.

- 요청 scope 에 `Tasks.*` 가 있으면 오류 코드와 무관하게 `invalid_grant` 를 `revoked` 로 두지 않고 `Tasks.*` 를 뺀 scope 로 한 번 더 refresh 한다.
  그것도 `invalid_grant` 면 그때 폐기로 본다 (refresh token 자체가 폐기된 경우). 이후 refresh 는 Tasks 없이 한다 (`todoConsentMissing`).
- To Do 도구 호출 시 `refreshWithTodoScopes()` 가 진행 중인 refresh 가 끝나기를 기다린 뒤 전체 scope 로 다시 시도하고,
  여전히 거절되면 "Tasks 권한 동의 필요", 갱신 자체가 실패하면(네트워크) 재시도를 안내한다.
- device flow 로그인이 성공하면 플래그를 지운다. `Tasks.*` 를 요청하지 않은 `invalid_grant` 는 기존처럼 `revoked` 로 처리한다.

### Graph 호출과 오류 안내

- 모든 path segment(목록 · task · checklist id) 는 `encodeURIComponent`. OData 질의는 `%24top` · `%24orderby` · `%24filter` 로 쓴다.
- 목록 · task 조회는 `@odata.nextLink` 를 최대 10페이지, 요청한 `top` 까지 따라간다. nextLink 가 `https://graph.microsoft.com/` 이 아니면 따르지 않는다 (토큰 유출 방지).
- 목록 인자는 id 또는 표시 이름. 생략하면 `wellknownListName == "defaultList"` (표시 이름은 로캘마다 다르다). 같은 이름이 여럿이면 id 로 지정하라고 안내한다.
- `todo_create_task` 의 checklist 는 task 를 만든 뒤 `POST …/checklistItems` 로 하나씩 만든다 (생성 본문에 함께 넣는 방식은 확인하지 못했다). 중간에 실패하면 만든 task id 와 추가된 개수를 알린다.
- 실패는 `{ text, isError: true }` 로 조치를 함께 돌려준다: 토큰 없음 → `/microsoft-teams:connect`, 403 → Tasks 권한 · 재로그인, 404 → id 확인,
  400 → Graph `error.message` + 입력 형식 (날짜가 있으면 시간대 안내), 429 · 5xx → 재시도, 네트워크 · 인증서 · 본문 오류 → 기존 SSL 힌트와 같은 문구.

### 실측하지 않은 항목

생성 · 수정 · 삭제 · checklist 호출, `Asia/Seoul`(IANA) `timeZone` 수용 여부, 동의 없는 scope 의 refresh 오류 형태, 기존 refresh token 으로 새 scope 를 받을 수 있는지,
`flaggedEmails` 목록 쓰기 거부 형태, `$top` 상한은 단위 테스트(fetch mock)로만 확인했다. 실환경에서 확인하면 이 절을 갱신한다.

## 설치 메커니즘

npm 글로벌 설치 경로를 로컬 디렉토리 마켓플레이스로 등록한다. 로컬 경로 마켓플레이스의 상대 경로 플러그인은
캐시로 복사되지 않고 제자리에서 로드되므로 (`plugin list --json` 의 `readFromFolder`) npm 업그레이드가 다음 세션부터 반영된다.

`install()` 은 멱등이다: 같은 경로로 등록돼 있으면 아무것도 바꾸지 않고, 경로가 바뀌었으면 (nvm 전환 등) remove → add 한다.

제자리 로드라 코드는 npm 업그레이드 즉시 바뀌지만 Claude Code 가 기록한 설치 버전 (`plugin list --json` 의 `version`, `folderVersion` 과 별개) 은
옛 값으로 남는다. `install()` 은 기록된 버전이 패키지 버전과 다르면 `plugin marketplace update` → `plugin update` 로 맞춘다 (실패해도 설치는 성공 처리).

## 설정

`~/.config/claudecode-ms-teams/claudecode-ms-teams.json` 단일 소스. `XDG_CONFIG_HOME` 만 존중한다.
chatId 는 설정에 저장하지 않는다 (세션 메모리 전용). `connect` 는 호출 시마다 설정을 다시 읽어 `cfg` 객체를 in-place 갱신한다
(Auth 가 같은 객체를 참조하므로).
