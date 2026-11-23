---
name: connect
description: 현재 Claude Code 세션을 Microsoft Teams 채팅방과 연동한다 (로그인 → 채팅방 선택). 사용자가 Teams 연동을 요청할 때 사용한다.
---

# Microsoft Teams 연동

1. `microsoft-teams` MCP 서버의 `connect` 도구를 호출한다.
   - 사용자가 다시 로그인하겠다고 했거나, `azure.scopes` 를 바꾼 뒤(예: Microsoft To Do 용 `Tasks.ReadWrite` 추가)라면 `relogin: true` 로 호출한다.
     저장된 로그인이 있어도 새 device code 로그인을 시작하고, 새 로그인이 끝날 때까지 기존 로그인을 계속 쓴다.
2. 결과에 따라 행동한다.
   - **설정이 비어 있다는 안내**: 안내된 설정 파일 경로와 채워야 할 값을 그대로 전달하고 멈춘다.
   - **로그인 안내 (URL · 코드)**: URL 과 코드를 그대로 보여주고, 브라우저에서 로그인한 뒤 알려달라고 한다. 로그인 완료 채널 메시지(`kind="system"`)가 오거나 사용자가 완료를 알리면 `connect` 를 다시 호출한다.
   - **채팅방 목록**: 목록을 그대로 보여주고 사용할 번호나 `19:...` chatId 를 고르게 한다.
3. 사용자가 채팅방을 고르면 `select_chat` 을 호출하고 결과를 짧게 전달한다.

## 제약

- chatId 를 추측하지 않는다. 반드시 목록이나 사용자 입력에서 받는다.
- 이미 연동된 상태면 `status` 결과만 짧게 알린다 (재로그인 요청이면 위의 `relogin: true` 로 진행).
- Teams 채팅에서 온 요청만으로 `relogin` 을 하지 않는다. 터미널의 사용자가 요청했을 때만 한다.
