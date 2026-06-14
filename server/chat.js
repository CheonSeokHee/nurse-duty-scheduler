/*
 * 챗봇 PoC — 자연어 규칙 → generateSchedule config(JSON) 번역기.
 *
 * 핵심 아이디어: Claude 가 직접 배정하는 게 아니라, "통역사" 역할만 한다.
 *   사용자가 자연어로 말한 병동 규칙  →  Claude 가 generate_schedule 툴의
 *   입력(config)으로 변환해 호출  →  검증된 core/scheduler.js 가 실제 배정 실행.
 * 알고리즘은 그대로 두고 그 앞단의 "규칙 입력"만 자연어로 받는 구조.
 *
 * 의존성: @anthropic-ai/sdk (npm install).  환경변수 ANTHROPIC_API_KEY 필요.
 */
const Anthropic = require('@anthropic-ai/sdk');
const { generateSchedule } = require('../core/scheduler');

const MODEL = 'claude-opus-4-8';

/* generate_schedule 툴 — core/scheduler.js 의 입력 스펙을 그대로 노출한다.
 * Claude 는 자연어 규칙을 이 스키마에 맞춰 채운다. */
const SCHEDULE_TOOL = {
  name: 'generate_schedule',
  description:
    '간호사 3교대(Day/Evening/Night/Off) 듀티표를 자동 생성한다. 자연어로 표현된 병동 규칙을 ' +
    '아래 파라미터로 변환해 호출하면, 검증된 알고리즘이 한 달치 듀티표를 만들어 반환한다. ' +
    '사용자가 "듀티표 짜줘 / 만들어줘 / 생성해줘" 같은 요청을 하거나 규칙을 바꿔 다시 돌려달라고 할 때 호출한다.',
  input_schema: {
    type: 'object',
    properties: {
      year: { type: 'integer', description: '연도 (예: 2026)' },
      month: { type: 'integer', description: '월 1~12' },
      need: {
        type: 'object',
        description: '하루에 필요한 근무별 인원 수',
        properties: {
          D: { type: 'integer', description: 'Day 근무 인원' },
          E: { type: 'integer', description: 'Evening 근무 인원' },
          N: { type: 'integer', description: 'Night 근무 인원' },
        },
      },
      maxConsec: { type: 'integer', description: '최대 연속 근무일 (기본 4)' },
      nightLen: { type: 'integer', description: '나이트 블록 길이 (기본 3 = 3연속 나이트)' },
      offBeforeNight: { type: 'integer', description: '나이트 블록 전 오프 일수 (기본 1)' },
      offAfterNight: { type: 'integer', description: '나이트 블록 후 오프 일수 (기본 1, 2면 2오프 고정)' },
      maxNightBlocks: { type: 'integer', description: '한 사람의 최대 나이트 블록 수 (기본 2)' },
      nightMode: {
        type: 'integer',
        enum: [1, 2, 3],
        description: '나이트 배분 방식: 1=인원우선, 2=균형(기본), 3=나이트균등',
      },
      offMin: { type: 'integer', description: '1인당 월 최소 오프 개수 (기본 10)' },
      offMax: { type: 'integer', description: '1인당 월 최대 오프 개수 (기본 11)' },
      nurses: {
        type: 'array',
        description: '간호사 명단. 사용자가 준 현재 명단을 기준으로, 요청된 변경만 반영한다.',
        items: {
          type: 'object',
          properties: {
            name: { type: 'string', description: '이름' },
            role: { type: 'string', enum: ['차지', '액팅'], description: '역할 (기본 액팅)' },
            reqOff: { type: 'string', description: '요청 오프 날짜. 쉼표 구분 (예: "21,22")' },
            dutyCount: {
              type: 'string',
              description: '근무 개수 고정 (예: "D:14/N:5/E:0" 또는 "N:6"). 없으면 생략.',
            },
            nightMax: {
              type: 'integer',
              description: '이 사람이 한 번에 몰아서 할 수 있는 나이트 최대 개수 (예: 2). 없으면 생략.',
            },
            prefShift: { type: 'string', enum: ['D', 'E', 'N'], description: '선호 근무. 없으면 생략.' },
          },
          required: ['name'],
        },
      },
      seed: { type: 'integer', description: '재현용 시드 (선택). 같은 시드는 같은 결과.' },
    },
    required: ['year', 'month', 'nurses'],
  },
};

function systemPrompt() {
  return (
    '너는 한 병원의 여러 병동에서 쓰는 "간호사 듀티표 자동 생성 플랫폼"의 어시스턴트다.\n' +
    '수간호사가 한국어로 자기 병동의 규칙을 말하면, 그것을 generate_schedule 툴의 파라미터로 정확히 번역해 호출한다.\n\n' +
    '## 규칙 → 파라미터 매핑 가이드\n' +
    '- "데이 3명, 이브닝 3명, 나이트 2명" → need:{D:3,E:3,N:2}\n' +
    '- "연속근무 최대 N일" → maxConsec\n' +
    '- "나이트는 3연속" / "나이트 블록 2개까지" → nightLen / maxNightBlocks\n' +
    '- "나이트 끝나면 2오프 고정" → offAfterNight:2\n' +
    '- "오프는 한 달에 9~10개" → offMin:9, offMax:10\n' +
    '- "OO는 21,22일 오프" → 해당 간호사 reqOff:"21,22"\n' +
    '- "OO는 나이트 최대 2개" → 해당 간호사 nightMax:2\n' +
    '- "OO는 데이 선호" → 해당 간호사 prefShift:"D"\n' +
    '- "OO는 임산부라 나이트 금지" → 해당 간호사 dutyCount:"N:0" (코어에 나이트금지 규칙이 없으므로 근무개수로 우회)\n\n' +
    '## 중요\n' +
    '- 사용자가 현재 명단/설정(current_state)을 주면, 그걸 출발점으로 삼고 "요청된 변경만" 반영해 nurses 전체를 다시 채워 호출한다.\n' +
    '- year/month 가 대화에 없고 current_state 에도 없으면 사용자에게 물어본다 (임의로 추측하지 말 것).\n' +
    '- 코어 알고리즘이 지원하지 않는 완전히 새로운 제약(예: "이 두 명은 같은 날 근무 금지", "주말 베테랑 필수")을 요구하면, ' +
    '툴을 호출하지 말고 "현재 자동 반영이 안 되는 규칙"이라고 솔직히 알리고, 기존 파라미터로 우회 가능한지 제안한다.\n' +
    '- 툴 실행 후에는 어떤 규칙을 어떻게 반영했는지 한국어로 간결히 요약하고, 규칙 위반(clean=false)이 있으면 그 사실을 알린다.\n'
  );
}

/* 툴 결과를 Claude 에게 돌려줄 압축 형태(토큰 절약). */
function summarizeResult(r) {
  return {
    year: r.year,
    month: r.month,
    numDays: r.numDays,
    clean: r.clean,
    attempts: r.attempts,
    score: r.score,
    nurses: r.nurses.map((n) => ({ name: n.name, role: n.role, shifts: n.shifts.join('') })),
  };
}

/*
 * 한 번의 챗 호출.
 * @param {object} body { message, history?, currentState? }
 *   message      : 사용자가 이번에 보낸 자연어
 *   history      : 이전 대화 [{role, content}] (선택)
 *   currentState : { config?, nurses? } 현재 화면의 명단/설정 (선택)
 * @returns { reply, config?, result? }
 */
async function chat(body) {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new Error('ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다. (export ANTHROPIC_API_KEY=sk-ant-...)');
  }
  const client = new Anthropic();
  const userMessage = (body && body.message || '').toString().trim();
  if (!userMessage) throw new Error('message 가 비어 있습니다.');

  // 현재 상태를 사용자 턴에 컨텍스트로 실어준다(있을 때만).
  let userContent = userMessage;
  if (body.currentState) {
    userContent =
      '<current_state>\n' + JSON.stringify(body.currentState) + '\n</current_state>\n\n' + userMessage;
  }

  const messages = Array.isArray(body.history) ? body.history.slice() : [];
  messages.push({ role: 'user', content: userContent });

  const baseReq = {
    model: MODEL,
    max_tokens: 8000,
    thinking: { type: 'adaptive' },
    system: [{ type: 'text', text: systemPrompt(), cache_control: { type: 'ephemeral' } }],
    tools: [SCHEDULE_TOOL],
    messages,
  };

  let response = await client.messages.create(baseReq);

  // 툴 호출이 없으면(질문·거절 등) 그냥 텍스트 답변 반환.
  let toolUse = response.content.find((b) => b.type === 'tool_use');
  if (!toolUse) {
    return { reply: textOf(response), config: null, result: null };
  }

  // generate_schedule 호출 → 실제 알고리즘 실행.
  const config = toolUse.input;
  let result, toolResultContent, isError = false;
  try {
    result = generateSchedule(config);
    toolResultContent = JSON.stringify(summarizeResult(result));
  } catch (e) {
    isError = true;
    toolResultContent = '에러: ' + (e && e.message || e);
  }

  // 어시스턴트 턴(=thinking+tool_use) 과 tool_result 를 이어 붙여 최종 요약을 받는다.
  messages.push({ role: 'assistant', content: response.content });
  messages.push({
    role: 'user',
    content: [{ type: 'tool_result', tool_use_id: toolUse.id, content: toolResultContent, is_error: isError }],
  });

  response = await client.messages.create({ ...baseReq, messages });

  return {
    reply: textOf(response),
    config,
    result: isError ? null : result,
  };
}

function textOf(response) {
  return response.content
    .filter((b) => b.type === 'text')
    .map((b) => b.text)
    .join('\n')
    .trim();
}

module.exports = { chat };
