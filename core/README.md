# duty-scheduler-core

간호사 듀티 스케줄 **알고리즘 코어**. 구글시트 버전(`../Code.gs`)에서 시트·메뉴 I/O를
뺀 순수 함수만 추출한 것으로, **Node 어디서든(서버·도커·Lambda) 단독 실행**된다.

구글시트 버전은 **그대로 유지**되며 이 폴더와 독립적이다. 이건 "나중에 서버로 갈 때
알고리즘을 다시 안 짜도 되게" 미리 떼어둔 코어다.

## 구성

| 파일 | 역할 |
|---|---|
| `_algorithm.js` | **자동 생성.** Code.gs에서 추출한 순수 함수(61개). 직접 수정 금지. |
| `extract-from-codegs.js` | 추출기. Code.gs 수정 후 다시 돌리면 `_algorithm.js` 동기화. |
| `scheduler.js` | **진입점.** `buildConfig` / `generateSchedule` (객체 입력 → 스케줄 출력). |
| `run-example.js` | 단독 실행 데모(기본 11명 표 출력). |
| `parity-check.js` | core ↔ Code.gs 결과 100% 동일 검증. |

## 사용

```bash
node run-example.js            # 기본 11명(2026/6) 표
node run-example.js 2026 7     # 연/월 지정
node run-example.js 2026 6 42  # 시드 고정(재현)
node parity-check.js           # Code.gs와 동일성 검증
node extract-from-codegs.js    # Code.gs 변경 후 코어 재생성
```

## 코드에서 호출

```js
const { generateSchedule } = require('./scheduler');

const result = generateSchedule({
  year: 2026, month: 6,
  need: { D: 2, E: 2, N: 2 },          // 근무별 필요인원
  offMin: 10, offMax: 11,              // 월 오프 범위
  nurses: [
    { name: '편혜경', role: '차지', reqOff: '21,22', dutyCount: 'D:14/N:5/E:0', nightMax: 2 },
    { name: '서문휘정', role: '액팅', reqOff: '12,13', dutyCount: 'N:6' },
    // ...
  ],
  // seed: 42,                          // (선택) 결과 재현
});

// result = { year, month, numDays, attempts, clean, score, nurses:[{name,role,shifts:[...]}] }
```

### 입력 필드

- **공통(선택, 기본값)**: `need{D,E,N}=2`, `maxConsec=4`, `nightLen=3`,
  `offBeforeNight=1`, `offAfterNight=1`, `maxNightBlocks=2`, `nightMode=2`,
  `offMin=10`, `offMax=11`, `attempts=300`, `maxAttempts=6000`, `timeBudgetSec=45`, `seed`
- **nurse**: `name`, `role`('차지'|'액팅'), `reqOff`('3,10,21' 또는 `{3:true,...}`),
  `dutyCount`('D:14/N:5/E:0' 또는 `{D,E,N}`), `prevNightDays`, `prefShift`('D'|'E'|'N'),
  `prefStrength`(1~3), `nightMax`(빈칸=nightLen)

### 출력

- `clean`: 빈칸·하드위반 없음 여부
- `score`: 세부 점수(`unfilled`, `overStaff`, `consecOffViol`, `patViol`, `offDev`,
  `dutyMiss`, `donMiss` 등 — 낮을수록 좋음)
- `nurses[].shifts`: 길이 `numDays` 의 `'D'|'E'|'N'|'S'|'O'` 배열

## 서버로 확장할 때

이 코어를 그대로 두고 껍데기만 씌우면 된다:

```
[웹 UI / CSV 업로드] → [Express API] → generateSchedule(input) → [표/엑셀 출력]
                                ↑ Dockerfile 로 패키징 → AWS(App Runner/Fargate/Lambda)
```

알고리즘 재작성은 **0**. 입력 파싱·인증·저장(DB)·출력 포맷만 추가하면 된다.
