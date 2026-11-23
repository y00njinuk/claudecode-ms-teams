# claudecode-ms-teams

Claude Code ↔ Microsoft Teams 양방향 채널 플러그인.

- Claude Code 턴이 끝나면 마지막 응답을 지정한 Microsoft Teams 채팅방으로 보낸다.
- 연동한 Teams 채팅방에 보내는 메시지는 **키워드 없이 모두** 실행 중인 Claude Code 세션에 전달된다.
- 도구 실행 권한 요청, 질문(AskUserQuestion · interview), plan 승인을 Teams 로 알리고, **코드 입력 없이 Teams 답장만으로 응답**할 수 있다
  (`#allow` · `#allow-always` · `#reject` · 답 그대로 · `#skip`). `#allow-always` 한 번이면 그 도구는 세션이 끝날 때까지 다시 묻지 않는다.
  터미널 다이얼로그도 그대로 떠 있어 먼저 한 응답이 적용된다.
- `claudecode-ms-teams send` 로 세션 없이 스크립트(예약 작업 보고 등)에서 Teams 로 보내고 결과를 종료 코드로 받는다.
- (선택) 같은 로그인으로 **Microsoft To Do** 의 목록 · task · checklist 를 조회 · 생성 · 수정한다 (`todo_*` 도구, `Tasks.ReadWrite` scope 를 넣었을 때만).

Microsoft Graph **delegated 권한 + 폴링** 방식이라 Azure Bot, 공개 HTTPS 엔드포인트(터널), Teams 앱 업로드가 필요 없다.
메시지는 봇이 아니라 로그인한 본인 계정으로 올라간다.
[opencode-ms-teams](https://github.com/y00njinuk/opencode-ms-teams) 의 Claude Code 판이다
(차이는 [opencode-ms-teams 와 다른 점](#opencode-ms-teams-와-다른-점) 참고).

## 요구사항

| 항목 | 조건 |
|---|---|
| Claude Code | 2.1.322 에서 검증. Teams → 세션 요청 주입은 [Channels](https://code.claude.com/docs/en/channels) 지원 버전 + claude.ai 로그인 또는 Console API key 필요 (Bedrock · Vertex · Foundry 미지원) |
| 조직 설정 | Team/Enterprise 플랜에서 채널을 쓰려면 관리자가 managed settings 의 `channelsEnabled` 를 켜야 한다 |
| Node.js | 20 이상 |
| Azure AD 앱 | delegated scope: `Chat.Create Chat.Read Chat.ReadBasic Chat.ReadWrite ChatMessage.Read ChatMessage.Send offline_access User.Read` |
| Azure AD 앱 (선택: Microsoft To Do) | delegated `Tasks.ReadWrite` (조회만 쓰면 `Tasks.Read`). 조직에 따라 관리자 동의가 필요하다 |

Azure AD 앱 등록/권한 신청은 조직의 Azure 관리자 안내를 따른다. opencode-ms-teams 에서 쓰던 앱(tenantId/clientId/scopes)을 그대로 써도 된다.

## 설치

```bash
npm install -g claudecode-ms-teams
```

글로벌 설치 시 postinstall 이 자동으로 다음을 수행한다.

1. 설정 파일 생성: `~/.config/claudecode-ms-teams/claudecode-ms-teams.json` (이미 있으면 그대로 둠)
2. npm 설치 경로를 Claude Code 로컬 마켓플레이스 `claudecode-ms-teams` 로 등록
3. 플러그인 `microsoft-teams@claudecode-ms-teams` 설치 (user scope)

자동 등록이 실패했거나 건너뛰었다면 `claudecode-ms-teams install` 로 다시 실행한다. 상태 확인은 `claudecode-ms-teams doctor`.

### 업그레이드

```bash
npm install -g claudecode-ms-teams@latest
```

플러그인은 npm 설치 경로에서 그대로 로드되므로 다음 세션부터 새 버전이 적용된다.
postinstall 이 Claude Code 마켓플레이스와 플러그인 버전 기록(`claude plugin list` 의 `version`)도 함께 갱신한다.
이미 실행 중인 세션은 다시 시작해야 새 hook 이 적용된다.

## 사용

### 1. 설정

`~/.config/claudecode-ms-teams/claudecode-ms-teams.json` 의 `azure.tenantId` / `azure.clientId` / `azure.scopes` 를 채운다.

### 2. 실행

```bash
claude --dangerously-load-development-channels plugin:microsoft-teams@claudecode-ms-teams
```

일반 `claude` 로 실행해도 플러그인은 로드된다. Teams 에서 세션으로 **새 요청을 보내는 기능만** 채널 모드가 필요하다.

| 기능 | 일반 실행 | 채널 모드 |
|---|---|---|
| 턴 종료 시 응답 Teams 발신 | ✅ | ✅ |
| 권한 요청 · 질문 · plan 승인 알림과 Teams 답장 응답 | ✅ | ✅ |
| Teams 채팅방 메시지 → 세션에 요청 주입 | ❌ | ✅ |

> Channels 는 research preview 기능이다. 이 플러그인은 Anthropic allowlist 에 없으므로 채널 모드로 실행할 때마다
> `--dangerously-load-development-channels` 플래그가 필요하다 (조직 관리자가 `allowedChannelPlugins` 에 등록하면 `--channels` 로 실행 가능).

턴이 끝나면 Stop hook 이 Teams 발신이 끝날 때까지 (최대 8초) 기다린다. 그래서 `claude -p` 처럼 응답 직후 종료해도 마지막 응답이 전송된다.
스크립트에서 결과만 보내려면 세션을 연동하지 말고 [`send`](#6-스크립트에서-보내기-send) 를 쓴다.

### 3. 연동

세션에서 `/microsoft-teams:connect` 를 실행한다.

1. 처음이면 device code 안내(URL + 코드)가 나온다. 브라우저에서 회사 계정으로 로그인한다.
2. 로그인이 끝나면 최근 채팅방 목록이 표시된다 (채널 모드가 아니면 "로그인했어" 라고 말하면 다시 조회한다).
3. 번호나 `19:...` chatId 를 고르면 연동이 끝난다. 그 채팅방의 메시지는 모두 이 세션에 전달되므로 **이 세션 전용 채팅방**을 고른다.

채팅방 선택은 세션 메모리에만 보관한다. Claude Code 를 다시 시작하면 다시 고른다 (로그인 토큰은 유지).

### 4. Teams 에서 요청 보내기

연동한 채팅방에 보낸 메시지는 **키워드 없이 그대로** 세션에 전달된다. 연동 시점 이전의 메시지는 전달하지 않는다.

| Teams 에서 보낸 메시지 | 세션에 전달되는 내용 |
|---|---|
| `카산드라 compaction 전략 알려줘` | `카산드라 compaction 전략 알려줘` |
| `#claude 테스트 돌려줘` · `가재야, 테스트 돌려줘` | `테스트 돌려줘` (예전 호출 키워드는 떼고 전달) |
| `#allow` · `#reject` 등 답장 명령 | **전달 안 됨** — 권한·질문 응답으로만 처리 |
| 이 플러그인이 보낸 응답·알림 (다른 세션이 보낸 것 포함) | **전달 안 됨** |
| 다른 사용자의 메시지 | **전달 안 됨** (`access.allowedUserIds` 에 있으면 전달) |

- 예전 호출 키워드(`#claude` · `#agent` · `#ai` · `#ask` · `#question` · `가재야` · `막둥아`)는 더 이상 필요 없다. 붙여 보내면 떼고 전달한다.
- 세션이 질문(AskUserQuestion)을 기다리는 동안 (질문 알림 뒤에) 보낸 메시지는 새 요청이 아니라 **그 질문의 답**으로 쓴다 (터미널 질문 창에 입력하는 것과 같다).
- 보낸 문자는 그대로 전달한다 (`ㅇㅇ` · `①` · `x²` 등을 바꾸지 않는다).
- 플러그인이 보내는 메시지는 모두 `[<제목> · <세션 8자>]` 헤더로 시작한다. 이 헤더로 시작하는 메시지는 세션에 넣지 않는다
  (같은 계정으로 연동된 다른 세션의 응답을 주고받는 루프 방지). 직접 보내는 메시지를 이 형식으로 시작하지 않는다.

### 5. Teams 답장으로 권한 · 질문에 응답하기

세션이 권한 승인이나 질문 응답을 기다리면 Teams 로 알림이 온다. **코드 없이 명령만 보내면** 대기 중인 요청에 적용된다.
터미널로 돌아가지 않고 Teams 답장만으로 응답할 수 있다.

| 명령 | 설명 |
|---|---|
| `#allow` | 권한을 이번 한 번만 승인 (plan 승인 요청이면 계획 승인) |
| `#allow-always` | 승인하고, **이 세션이 끝날 때까지 같은 도구의 요청은 묻지 않고 자동 승인**. Claude Code 가 제안한 "앞으로 묻지 않기" 규칙이 있으면 함께 저장한다 (알림에 표시) |
| `#allow-reset` | `#allow-always` 로 켠 자동 승인을 모두 끈다 |
| `#reject` | 권한 거부 |
| `<선택지>` 또는 `<번호>` | 질문에 응답. `#answer` 없이 답만 보내면 된다 (복수 선택 질문은 쉼표로 여러 개) |
| `text: <자유 텍스트>` | 선택지와 대조하지 않고 그대로 응답 (숫자 자체가 답이거나 쉼표가 들어간 답) |
| `#skip` | 질문 전체를 건너뜀 |

#### 권한 요청

```
[⚠️ Microsoft Teams: 권한 요청 · 3ab7f0e2] · pa7k9m
Tool: Bash
설명: 테스트 실행
  npm test
Workspace: /home/user/my-project

Teams 답장으로도 응답 가능:
• #allow
• #allow-always — 승인하고 이 세션에서 Bash 요청은 계속 자동 승인 (Claude Code 규칙도 저장: Bash(npm test))
• #reject
터미널에서 먼저 응답하면 그 응답이 적용됩니다.
```

`#allow` 로 답하면 세션이 바로 진행되고 `[✅ 전달됨 · 3ab7f0e2] allow · Bash: npm test` 확인 메시지가 온다.
workspace 밖 파일을 건드리는 요청은 `⚠️ Workspace 외부 경로` 로 강조한다.

#### 자동 승인 (`#allow-always`)

- 범위는 **도구 단위**다. `Bash` 요청에 `#allow-always` 하면 이후의 모든 `Bash` 명령이 승인된다 (명령 내용과 무관). `Edit` · MCP 도구 등 다른 도구는 계속 묻는다.
- 같은 도구로 이미 대기 중인 요청도 함께 승인한다 (아래 예외 제외). 턴이 바뀌거나 sub-agent 가 요청해도 계속 적용된다.
- plan 승인 요청에 쓰면 이 세션의 이후 계획도 자동 승인한다.
- 예외 — 자동 승인 중이어도 다시 묻는다 (알림에 이유 표시):
  - 질문(AskUserQuestion) — 답이 필요하다
  - workspace 밖 경로 (`~/…` 포함, Glob 은 패턴도 검사)
  - `.claude/` · `.git/` · `.mcp.json` — 고치면 hook · 권한 규칙으로 임의 명령 실행까지 이어질 수 있다
  - 세션의 workspace 를 아직 모를 때
- 대기 요청이 여럿일 때 코드 없이 `#allow-always` 를 보내면, 모두 같은 도구일 때만 그 도구를 등록해 승인한다 (이때 Claude Code 규칙은 저장하지 않는다). 도구가 섞여 있으면 코드를 붙인다.
- 자동 승인된 요청은 Teams 로 따로 알리지 않는다. 현재 자동 승인 중인 도구는 `status` 도구로 확인한다.
- 해제: `#allow-reset`, 연동 해제(`disconnect`), 다른 채팅방 선택, Claude Code 재시작. Claude Code 에 함께 저장된 규칙은 `#allow-reset` 으로 지워지지 않으므로 터미널의 `/permissions` 에서 정리한다.

#### 질문 (AskUserQuestion)

```
[❓ Microsoft Teams: 세션 질문 · 3ab7f0e2] · q3x9m2
배포 승인
production 에 배포할까요?
1. Yes — 지금 배포
2. No — 취소

Teams 답장으로도 응답 가능:
• <선택지 또는 번호> — 그대로 답장
• text: <자유 텍스트> — 선택지와 대조하지 않고 그대로 전달
• #skip
```

| 상황 | Teams 답장 | 세션이 받는 답 |
|---|---|---|
| 단일 선택 | `yes` 또는 `1` | `Yes` |
| 복수 선택 | `staging, dev` | `staging, dev` |
| 선택지에 없는 답 | `내일 오전에 다시 물어봐` | `내일 오전에 다시 물어봐` |
| 자유 텍스트 강제 | `text: Feat: 알림 큐, 비활성화` | `Feat: 알림 큐, 비활성화` |
| 질문 여러 개 | `prod` → `📝 답변 기록됨 … 다음: <다음 질문>` → `main` | 질문 순서대로 하나씩 받아 모두 채워지면 한 번에 전달 |
| 건너뛰기 | `#skip` | 질문 거부 |

- 선택지 label 은 대소문자를 구분하지 않는다. 선택지에 없는 답은 입력한 그대로(자유 텍스트) 전달한다.
- 번호가 선택지 범위를 벗어나면 `❌ 답변 미반영` 으로 알리고 질문은 계속 기다린다. 숫자 자체가 답이면 `text: 7` 처럼 보낸다.
- sub-agent 가 낸 질문은 새 요청과 구분하기 위해 `#answer <답>` 으로 답한다 (알림에 표시된다).
- 질문 알림이 오기 전에 보낸 메시지와 `#skp` 처럼 `#` 으로 시작하는 오타는 답으로 쓰지 않는다.
- 터미널에서 Esc 로 질문을 닫으면 Claude Code 가 알려주지 않아 질문이 대기로 남고, 다음 Teams 메시지가 그 답으로 쓰인다.
  확인 메시지(`✅ 전달됨 answer …`)로 드러나므로 다시 보내면 된다 (질문 묶음이면 `#skip`).

#### 공통 규칙

- **대기 중인 요청이 하나면 코드가 필요 없다.** 코드 없는 답장은 그 답장을 보내기 전에 알림이 온 요청에만 적용된다
  (답장을 쓰는 사이 새로 온 요청은 승인하지 않는다). 같은 종류(권한 / 질문)의 요청이 여럿 대기 중이면 엉뚱한 요청을 승인하지 않도록 적용하지 않고,
  `⚠️ 대상 지정 필요` 로 각 요청의 코드를 알려준다. 그때만 `#allow pa7k9m` · `#answer q3x9m2 yes` · `#answer q3x9m2-2 dev` 처럼 코드를 붙인다
  (코드는 알림 헤더에도 표시된다. `p`/`q` + 영숫자 5자, 대소문자 무관).
- 대기 요청은 24시간 뒤 만료된다.
- **터미널에서 먼저 응답하면 그 응답이 적용된다.** 터미널에서 답한 질문은 바로 대기에서 빠지고, 권한 요청은 턴(sub-agent 의 요청은 그 sub-agent)이 끝날 때 빠진다.
- 대기 중인 요청이 없거나, 모르는 코드 · 형식이 틀린 답장(`#allow 지금` 등)은 **후속 메시지 없이 무시**한다.
- 확인 메시지는 `✅ 전달됨` 이다. Claude Code 가 적용 결과를 알려주지 않으므로, 같은 턴 안에 터미널에서 이미 응답한 권한 요청에 뒤늦게 답하면 전달은 되지만 적용되지 않는다.
- sub-agent 가 낸 요청에는 `🌿 sub-agent <이름>` 배너가 붙는다.

### 6. 스크립트에서 보내기 (`send`)

세션 없이 Markdown 한 건을 채팅방으로 보낸다. 예약 작업이 `claude -p` 결과를 아침에 보고하는 경우처럼 무인 스크립트에서 쓴다.

```bash
cat results/pending/*.md | claude -p "$(cat prompts/report.md)" > report.md
claudecode-ms-teams send --chat "19:...@thread.v2" --title claude-nightly --file report.md   # --file 생략 시 stdin
```

| 종료 코드 | 의미 | 조치 |
|---|---|---|
| 0 | 전송됨 (stdout 에 message id) | — |
| 1 | 발신 실패 (네트워크 · Graph 오류, 원인은 stderr) | chatId · 네트워크 확인 후 재시도 |
| 2 | 사용법 오류 (인자 · 빈 본문 · 파일 읽기) | 인자 확인 |
| 3 | 설정 · 로그인 필요 (azure 미설정 · 토큰 없음 · 만료) | 대화형 `claude` 에서 `/microsoft-teams:connect` 로 로그인 |

- 대화형 세션에서 한 번 로그인해 둔 `tokens.json` 을 쓴다. `send` 는 로그인(device code)을 시작하지 않는다.
- chatId 는 `/microsoft-teams:connect` 목록에 나온 `19:...` 값이다. 메시지는 `[<제목> · -]` 헤더로 시작해 같은 채팅방에 연동된 세션에는 전달되지 않는다.
- 본문이 Teams 메시지 크기(약 25KB HTML)를 넘으면 뒷부분을 자르고 알린다.

### 7. Microsoft To Do (선택)

같은 Microsoft 로그인으로 To Do 를 쓴다. 별도 앱 · 로그인은 필요 없다. 쓰지 않으면 아무것도 바뀌지 않는다 (도구도 노출하지 않는다).

1. Azure AD 앱에 delegated `Tasks.ReadWrite` 권한을 추가한다 (조직 정책에 따라 관리자 동의).
2. 설정의 `azure.scopes` 끝에 `Tasks.ReadWrite` 를 추가한다 (조회만 하려면 `Tasks.Read`).
3. **Claude Code 세션을 다시 시작한다.** 도구 목록은 세션 시작 시 정해진다.
4. 이미 로그인해 있었다면 터미널에서 "Teams 다시 로그인해줘" 처럼 요청해 `/microsoft-teams:connect` 를 `relogin` 으로 실행한다.
   새 로그인이 끝날 때까지 기존 로그인으로 Teams 는 계속 동작한다. 관리자 동의가 이미 있으면 재로그인 없이 바로 쓰일 수도 있다.

| `azure.scopes` | 도구 |
|---|---|
| `Tasks.ReadWrite` | `todo_list_lists` · `todo_list_tasks` · `todo_get_task` · `todo_create_task` · `todo_update_task` · `todo_delete_task` · `todo_checklist` |
| `Tasks.Read` 만 | 조회 3개 (`todo_list_lists` · `todo_list_tasks` · `todo_get_task`) |
| 없음 | 없음 (Teams 기능만) |

세션에서 자연어로 요청하면 된다.

```
To Do 목록 보여줘
"업무" 목록에 내일까지 "주간 보고서 작성" 추가하고 체크리스트로 초안 · 검토 넣어줘
보고서 task 금요일 오후 3시에 알림 걸어줘
보고서 완료 처리해줘
```

- 목록은 id 또는 이름으로 지정한다. 생략하면 기본 목록(To Do 의 "작업" / "Tasks")을 쓴다. 이름이 같은 목록이 여럿이면 id 로 지정하라고 안내한다.
- 기한은 날짜(`YYYY-MM-DD`), 알림은 `YYYY-MM-DDTHH:mm` (시각 생략 시 09:00) 으로 해석한다. 빈 값으로 수정하면 해제한다.
  시간대는 설정의 `todo.timeZone`, 비어 있으면 시스템 시간대(예: `Asia/Seoul`)다. 조회 결과의 시각도 이 시간대로 보여준다.
- 삭제(`todo_delete_task`)는 복구할 수 없다. Claude 는 삭제 전에 확인을 받도록 안내받는다.
- `status` 도구가 To Do 사용 상태와 저장된 토큰의 Tasks 권한을 보여준다.

```
To Do: 사용 (Tasks.ReadWrite) · 토큰 권한 Tasks.ReadWrite
```

| 안내 | 원인 · 조치 |
|---|---|
| `저장된 로그인 토큰에 Tasks.ReadWrite 권한이 없습니다` | scope 추가 전에 로그인했거나 다른 세션이 예전 설정으로 토큰을 갱신했다. 터미널에서 `/microsoft-teams:connect` 로 다시 로그인(relogin) |
| `Microsoft To Do 권한(Tasks) 동의가 필요합니다` | Azure 앱에 Tasks 권한 · 관리자 동의가 없다. 이때도 Teams 기능은 계속 동작한다 (To Do scope 를 빼고 토큰을 갱신) |
| `HTTP 403` | Tasks 권한 없음. `azure.scopes` 와 Azure 앱 권한 확인 후 재로그인 |
| `HTTP 404` | 목록 · task · checklist id 가 틀렸다 |
| `HTTP 400 … todo.timeZone` | 시간대를 Graph 가 받지 않으면 `todo.timeZone` 을 Windows 시간대 이름(예: `Korea Standard Time`)으로 바꾼다 |

### 보안

- 세션을 조작할 수 있는 사람은 **로그인한 본인**뿐이다. 동료에게 허용하려면 `access.allowedUserIds` 에 Entra 사용자 ID 를 넣는다.
  허용된 사람의 메시지는 모두 세션에 전달되고, Teams 답장으로 도구 실행을 승인(`#allow-always` 포함)하고 질문에 답할 수 있다는 점에 유의한다.
- `#allow-always` 는 도구 단위로 넓게 승인한다 (`Bash` 면 모든 명령). 원격에서 켜 둔 자동 승인은 필요 없어지면 `#allow-reset` 으로 끈다.
- 로그인(device code)은 터미널에서 `/microsoft-teams:connect` 로만 시작한다. 토큰이 만료·폐기되면 자동으로 다시 로그인하지 않고 안내만 하며, 로그인 진행 중 Teams 로 나가는 응답에서는 코드를 가린다. 로그인 후 계정이 바뀌었으면 연동을 해제한다.
- 토큰은 `tokens.json` (권한 0600) 에만 저장한다.
- hook ↔ MCP 서버 통신은 `127.0.0.1` 과 세션별 임의 토큰으로 제한한다.
- **한 채팅방에는 세션 하나만 연동한다.** 채팅방의 메시지와 코드 없는 답장 명령(`#allow` 등)은 연동된 모든 세션에 적용된다. 다른 사람과 대화하는 채팅방도 고르지 않는다 (본인 메시지가 모두 세션에 들어간다).

## 설정 파일

```jsonc
{
  "azure": { "tenantId": "", "clientId": "", "scopes": "" },
  "polling": { "intervalMs": 5000, "backoff1Ms": 600000, "backoff2Ms": 3600000 },
  "access": { "allowedUserIds": [] },
  "log": { "file": "~/.config/claudecode-ms-teams/claudecode-ms-teams.log", "enabled": false },
  "tokens": { "file": "~/.config/claudecode-ms-teams/tokens.json" },
  "todo": { "timeZone": "" }
}
```

| 섹션 | 설명 |
|---|---|
| `azure` | Azure AD 앱 설정. `scopes` 에 `offline_access` 필수 (refresh token) |
| `polling` | 기본 폴링 주기. 마지막 활동 후 `backoff1Ms` 가 지나면 3배, `backoff2Ms` 가 지나면 6배로 늘어난다 (기본 5초 → 15초 → 30초) |
| `access.allowedUserIds` | 본인 외에 세션 요청과 답장 명령을 보낼 수 있는 Entra 사용자 ID |
| `log` | `enabled=false` (기본) 면 로그 파일을 만들지 않는다. `true` 면 세션 시작 시 비우고 기록한다 |
| `tokens` | 토큰 파일 경로 |
| `todo.timeZone` | Microsoft To Do 기한 · 알림을 해석하는 시간대 (IANA 예: `Asia/Seoul`, 또는 Windows 이름). 비어 있으면 시스템 시간대 |

`XDG_CONFIG_HOME` 이 설정돼 있으면 `~/.config` 대신 그 경로를 쓴다.
사내망 프록시 인증서 오류가 나면 `NODE_EXTRA_CA_CERTS=<사내 CA 번들>` 을 설정한 뒤 Claude Code 를 다시 시작한다.

## CLI

```
claudecode-ms-teams install             Claude Code 에 플러그인 등록 (업그레이드 후에는 버전 기록 갱신)
claudecode-ms-teams uninstall [--purge] 등록 해제 (--purge: 설정·토큰도 삭제)
claudecode-ms-teams doctor              설치 상태 진단 (Microsoft To Do 설정 · 토큰 권한 포함)
claudecode-ms-teams send --chat <19:...> [--file <경로>] [--title <제목>]
                                        세션 없이 채팅방으로 한 건 발신 (종료 코드 0 · 1 · 2 · 3)
```

## 제거

npm 은 글로벌 패키지 제거 시 스크립트를 실행하지 않으므로 등록 해제를 먼저 한다.

```bash
claudecode-ms-teams uninstall   # 설정·토큰까지 지우려면 --purge
npm uninstall -g claudecode-ms-teams
```

## opencode-ms-teams 와 다른 점

| 항목 | opencode-ms-teams | claudecode-ms-teams | 이유 |
|---|---|---|---|
| 세션 요청 | 키워드(`#opencode` 등)로 시작하는 메시지만 | 연동한 채팅방의 모든 메시지 (키워드 불필요) | 키워드 입력 부담 제거. 두 플러그인이 같은 채팅방을 쓰면 `#opencode` 메시지도 Claude 에 전달되므로 채팅방을 나눈다 |
| 발신자 | 채팅방 멤버 누구나 | 로그인 본인 + `access.allowedUserIds` | 답장으로 도구 실행을 승인할 수 있으므로 발신자를 인증한다 |
| 확인 메시지 | `✅ 반영됨` / `❌ 적용 실패` | `✅ 전달됨` | Claude Code 는 hook 결정의 적용 결과를 돌려주지 않는다 |
| 답장 코드 | 항상 필요 | 대기 요청이 여럿일 때만 필요 | 모바일에서 입력 부담을 줄이기 위함 |
| `#allow-always` | 같은 패턴을 항상 승인 | 세션 동안 같은 도구를 자동 승인 (+ Claude Code 제안 규칙 저장). `#allow-reset` 으로 해제 | 매번 승인하는 부담을 줄이기 위함 |
| 질문 응답 | `#answer <code> <label>` | 답만 보내면 된다 (label · **선택지 번호** · `text:` · 선택지에 없는 답은 자유 텍스트) | 모바일에서 답하기 쉽게 |
| 연동 시작 | `/connect-teams` | `/microsoft-teams:connect` | 플러그인 skill 네임스페이스 |

## 개발

```bash
npm install --no-audit --no-fund
bun test              # 테스트
bunx tsc --noEmit     # 타입 체크
npm run build         # dist 빌드
claude --plugin-dir . # 이 저장소를 플러그인으로 바로 로드
```

설계는 저장소의 `ARCHITECTURE.md` 참고.

## License

MIT
