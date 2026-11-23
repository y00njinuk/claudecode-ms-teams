# CLAUDE.md — claudecode-ms-teams

Claude Code ↔ Microsoft Teams 양방향 채널 플러그인 (npm 배포). 설계는 `ARCHITECTURE.md`, 이 파일은 코드를 바꿀 때 지킬 규칙이다.

## 저장소 구조

```
claudecode-ms-teams/
├── .claude-plugin/
│   ├── plugin.json          ← 플러그인 매니페스트 (name=microsoft-teams, mcpServers 인라인)
│   └── marketplace.json     ← 패키지 루트를 마켓플레이스로 노출 (name=claudecode-ms-teams)
├── hooks/hooks.json         ← Stop · StopFailure / SubagentStop / PermissionRequest / PostToolUse(AskUserQuestion) → dist/hook.js
├── skills/connect/SKILL.md  ← /microsoft-teams:connect
├── src/
│   ├── core.ts              ← 설정 · 로그 · Auth · Graph · 트리거 · 렌더링
│   ├── mcp.ts               ← 최소 MCP stdio 서버
│   ├── channel.ts           ← 세션 상태 머신 (도구 · Stop 발신 · 폴링 · 권한/질문 응답)
│   ├── server.ts            ← MCP 서버 진입점
│   ├── ipc.ts               ← hook ↔ 서버 localhost IPC + 레지스트리
│   ├── hook.ts              ← command hook 클라이언트
│   ├── install.ts · cli.ts  ← 설치/제거/진단
│   ├── send.ts              ← 세션 없는 발신 CLI (`claudecode-ms-teams send`)
│   ├── todo.ts              ← 선택 기능 Microsoft To Do 도구 (`todo_*`)
├── postinstall.mjs          ← 글로벌 설치 시 자동 등록
├── test/                    ← bun test
└── dist/                    ← tsc 산출물 (.gitignore, npm tarball 포함)
```

## 핵심 불변식

1. **런타임 의존성 제로**: Node 내장 `fetch` · `http` · `node:` 빌트인만 쓴다. MCP SDK 도 쓰지 않는다 (플러그인이 npm 설치 경로에서 제자리 로드되므로 의존성 설치 단계가 없다).
2. **stdout 은 MCP 전용**: `server.ts` / `channel.ts` / `core.ts` / `todo.ts` 에서 `console.log` 금지. 진단은 `log` (파일) 로만. `hook.ts` 는 `permission` 이벤트에서 Teams 답장으로 정해진 결정 JSON 만 stdout 에 쓰고, 그 외에는 아무것도 쓰지 않으며 항상 exit 0 으로 끝낸다.
3. **Lazy connect**: `connect` / `select_chat` / `todo_*` 도구 호출 전에는 Azure · Graph 요청을 보내지 않는다. 인증 이벤트의 채널 push 도 `connectRequested` 이후에만. (`send` CLI 는 사용자가 직접 실행하는 별도 프로세스다.)
4. **세션 식별은 claude PID**: 레지스트리 키는 서버의 `process.ppid`, hook 은 `CLAUDE_PID` → `process.ppid`. `session_id` 는 `/clear` 에서 바뀌므로 키로 쓰지 않는다. MCP 서버 command 를 셸·npx 래퍼로 감싸면 ppid 가 바뀌어 hook 이 서버를 못 찾는다 — `node ${CLAUDE_PLUGIN_ROOT}/dist/server.js` 직접 실행을 유지할 것.
5. **발신자 게이트**: 세션 주입과 답장 명령(권한 승인 · 질문 응답)은 로그인 본인 또는 `access.allowedUserIds` 만 허용한다. PermissionRequest hook 결정은 이 게이트를 전제로 한다. 게이트는 트리거 판정보다 먼저 적용한다.
6. **답장 verb 는 세션에 주입하지 않는다**: `REPLY_VERB_KEYWORDS` (`#allow-always` · `#allow-reset` · `#allow` · `#reject` · `#answer` · `#skip`) 매칭 시 파싱 성공/실패와 무관하게 `executeReply` 로만 간다. 모르는 코드 · 형식 오류 · 대기 요청 없음은 Teams 후속 메시지 없이 로그만 남긴다.
   코드는 선택이다. 코드 없는 답장은 답장 id 보다 `alertId` 가 작은(먼저 알린) 같은 종류의 대기 요청이 정확히 하나일 때만 적용하고, 여럿이면 적용하지 않고 코드를 안내한다 (최근 · 오래된 요청을 추측해 고르지 말 것). 코드 없는 `#allow-always` 만 예외로, 후보가 모두 같은 도구면 그 도구를 등록한다 (어느 요청의 `permission_suggestions` 도 저장하지 않는다). 세션에 넘기는 본문은 NFKC 하지 않는다 (`payloadVerbatim`).
7. **커서**: 폴링 배치의 최대 id 로 게이트와 무관하게 전진한다. 발신 시점에는 전진시키지 않는다 (자기 발신은 `sentIds`). `latestMessageId` / `messagesSince` 의 `null` (실패) 을 0 · 빈 배열로 축약하지 말 것.
8. **TRIGGER_RE**: alternation 은 길이 내림차순 유지 (테스트로 강제). 정규화 순서는 `normalizeForTrigger()` 주석 기준. 수신은 키워드 없이 모든 메시지를 받으며, 예전 호출 키워드(`#claude` 등)는 떼고 전달만 한다.
9. **chatId 는 설정에 두지 않는다**: 세션(서버 프로세스) 메모리 전용.
10. **설정 단일 소스**: `~/.config/claudecode-ms-teams/claudecode-ms-teams.json` (`XDG_CONFIG_HOME` 만 존중). 설정용 환경변수 오버라이드 도입 금지. `os.homedir()` 직접 호출 금지 — `homeDir()` 경유.
11. **버전 동기화**: `package.json` 과 `.claude-plugin/plugin.json` 의 `version` 은 같아야 한다 (테스트로 강제).
12. **이름**: 플러그인·MCP 서버 이름 `microsoft-teams` (채널 `source` 속성이 된다), 마켓플레이스 `claudecode-ms-teams`. 사용자 노출 문자열은 "Microsoft Teams" 로 쓴다.
13. **Teams HTML**: 코드 블록은 class 없는 `<pre style><code>`. `class="language-*"` 와 `white-space:pre` div 금지.
14. **토큰**: `tokens.json` 은 0600, tmp + rename 으로 원자적 저장. 저장소 · tarball 에 넣지 않는다.
15. **device flow 는 `connect` 도구에서만 시작**: `send` CLI 도 시작하지 않는다 (토큰이 없거나 만료면 종료 코드 3). refresh 실패(`invalid_grant`) 시 자동 재시작 금지 — 코드가 Stop 발신으로 Teams 에 새면 채팅방 멤버가 대신 로그인한다. 진행 중에는 Teams 로 나가는 Stop 발신 · 권한·질문 알림 · 확인 메시지 모두 `redactDeviceCode` 로 user code 를 가리고, 로그인 후 `/me` id 가 바뀌면 연동을 해제한다.
16. **장애 격리**: `server.ts` 의 비동기 호출은 `.catch` 필수, `Graph.call()` 은 본문 파싱 실패를 `null` 로 흡수, `pollOnce` 는 메시지 단위 try/catch.
17. **권한 · 질문 응답은 PermissionRequest hook 으로만**: `claude/channel/permission` capability 를 선언하지 않는다 (같은 요청이 두 번 알림된다). PermissionRequest 대기를 PreToolUse 로 옮기지 말 것 — 기다리는 동안 터미널 다이얼로그가 뜨지 않는다.
18. **pending 은 반드시 settle**: 폐기할 때는 `dropPending` 으로 `settle(null)` 을 호출해 기다리는 hook 프로세스를 끝낸다 (`Map.delete` 만 하면 hook 이 timeout 까지 남는다). 알림 발신에 실패한 pending 은 바로 폐기한다. pending 을 정리하는 신호는 Stop · StopFailure (main), SubagentStop (그 `agent_id`), PostToolUse(AskUserQuestion), hook 연결 종료, TTL, disconnect · 계정 변경 · 다른 채팅방 선택 (전부) 이고, `#allow-always` 는 같은 도구의 대기 요청을 allow 로 settle 한다.
19. **hook timeout = 답장 코드 TTL**: `hooks.json` 의 PermissionRequest `timeout`(초) 과 `REPLY_TTL_MS` 는 같아야 한다 (테스트로 강제). hook 의 장시간 대기는 `node:http` 로 한다 (`fetch` 는 헤더 대기 5분 제한).
20. **Claude Code 동작 가정은 실측 버전을 적는다**: PermissionRequest 와 다이얼로그의 경합, `updatedInput.answers`, `updatedPermissions` 는 2.1.322, 종료 시 MCP 서버 시그널 순서는 2.1.330 실측이다. 바꿀 때는 `ARCHITECTURE.md` 의 실측 목록을 다시 확인한다.
21. **플러그인 발신은 헤더로 시작한다**: Teams 로 보내는 모든 메시지는 `renderHeader()` 의 `[<제목> · <shortSessionId>]` 로 시작해야 한다 (sub-agent 배너 등은 헤더 뒤, `send` CLI 는 세션 자리에 `-`). 수신은 `isPluginMessage()` 로 이를 걸러 세션 간 응답 루프를 막는다 — 새 발신 종류를 추가하면 `channel.test.ts` 의 루프 방지 테스트에 넣는다.
22. **자동 승인은 세션 메모리 전용**: `autoAllowTools` 는 설정 · 레지스트리에 저장하지 않는다. `AskUserQuestion`, workspace 밖 경로(`~` 포함), `.claude/` · `.git/` · `.mcp.json`, workspace 를 모를 때는 자동 승인하지 않는다 (`autoAllowBlocker`). `disconnect` · 다른 채팅방 선택에서 비운다.
23. **Stop hook 은 발신이 끝날 때까지 기다린다**: 서버는 `stop` 응답을 `holdAtMost(handleStop, STOP_WAIT_MS)` 로 잡아둔다. 즉시 202 로 되돌리지 말 것 — claude 는 Stop hook 이 끝나면 MCP 서버를 SIGINT 후 0.5초 안에 SIGKILL 해 `claude -p` 의 마지막 응답이 사라진다. `STOP_WAIT_MS` < `ipcTimeoutMs("stop")` < hooks.json Stop · StopFailure timeout (테스트로 강제).
24. **`send` CLI 는 플러그인 로그를 쓰지 않는다**: `makeLogger` 는 시작할 때 로그 파일을 비우므로 실행 중인 세션의 로그가 지워진다. 진단은 stderr, 결과는 종료 코드 (0 · 1 발신 실패 · 2 사용법 · 3 설정·로그인). chatId 는 인자로만 받는다 (불변식 9 · 10).
25. **To Do 는 선택 기능**: 노출(`tools/list` · instructions)은 서버 시작 시 `azure.scopes` 로 정한다 (`Tasks.ReadWrite` 7개 · `Tasks.Read` 조회 3개 · 없음 0개, 없으면 0.4.0 과 동일). `REQUIRED_SCOPES` 에 Tasks 를 넣지 말 것. `Tasks.*` 를 요청한 refresh 의 `invalid_grant` 는 바로 `revoked` 로 두지 않고 Tasks 를 뺀 scope 로 한 번 더 받는다 (Teams 보호, 그것도 실패하면 폐기). 기한은 날짜를 설정 시간대 그대로 보낸다 — UTC 로 바꿔 보내지 말 것 (To Do 가 보낸 시간대의 날짜를 그대로 쓴다). 응답의 `@odata.nextLink` 는 Graph 주소일 때만 따른다.

## 개발 워크플로

```bash
npm install --no-audit --no-fund
bun test              # 전부 pass (node-runtime.test.ts 는 tsc 빌드 결과를 Node 로 실행한다)
bunx tsc --noEmit     # 오류 0
npm run build
claude plugin validate .
npm publish --dry-run # tarball 구성 확인
claude --plugin-dir . # 도구·hook 실기 확인 (채널 수신은 설치본 + 채널 모드로 확인)
```

새 기능: 실패 테스트 먼저 (`test/`) → 구현 → 위 검증 통과 → `ARCHITECTURE.md` · `README.md` · 이 파일 갱신.

## Code Style

- 주석은 코드만으로 의도가 드러나지 않을 때만, 간결하게.
- 이름만으로 목적이 보이게 짓는다.

## Git Commit Guidelines

- 한국어로 작성. 제목과 본문은 빈 줄로 구분, 본문은 필요할 때만.
- 1 변경 = 1 커밋, 커밋당 최대 5개 파일.
- `ARCHITECTURE.md` 에 반영할 내용이 없는지 매번 검토한다.
- 제목: `<Type>: <subject>` (50자 이내, 마침표 없음, 명령조). type: feat, fix, chore, refactor, docs, style, test, perf, ci, build, revert
