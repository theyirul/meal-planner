// POST /parse
// 부모가 올린 식단표에서 브라우저가 뽑은 텍스트를 받아, Claude로 날짜·끼니·메뉴·알레르기를 구조화한다.
// 파일 자체는 받지 않는다(텍스트만). 저장도 하지 않는다 — 결과는 부모 기기 localStorage에만 남는다.
//
// 응답: 성공이면 Anthropic SSE 스트림을 그대로 넘긴다(브라우저가 text_delta를 이어 붙인다).
//        실패면 {"error":"사람이 읽을 문장"} JSON + 4xx/5xx.
//
// 환경 변수
//   ANTHROPIC_API_KEY  (secret)
//   ALLOWED_ORIGINS    쉼표 구분. 예: https://theyirul.github.io,http://localhost:8000
//   MONTHLY_LIMIT      한 달 전체 호출 상한 (기본 100)
//   DAILY_PER_IP       IP 하나당 하루 상한 (기본 6)
// 바인딩
//   USAGE              KV. 상한 카운터

// 속도 우선(2026-10-08): Opus로 한 달치 3~5분 → 부모가 기다리다 나간다. 정확도는 동대문 정답지로 재검증
const MODEL = "claude-opus-5-5";
const MAX_INPUT_CHARS = 120_000;

const AN = {1:"난류(계란)",2:"우유",3:"메밀",4:"땅콩",5:"대두",6:"밀",7:"고등어",8:"게",9:"새우",10:"돼지고기",11:"복숭아",12:"토마토",13:"아황산류",14:"호두",15:"닭고기",16:"쇠고기",17:"오징어",18:"조개류",19:"잣"};

const SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["title", "year", "month", "items", "notes"],
  properties: {
    title: { type: "string", description: "원본 머리말에 적힌 식단표 이름(기관·반·식단 종류). 예: '동대문구센터 일반식(1-2세)'. 날짜·발행일은 빼고 30자 이내. 없으면 빈 문자열" },
    year: { type: "integer" },
    month: { type: "integer" },
    items: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["date", "sec", "name", "allergy", "check", "why"],
        properties: {
          date: { type: "string", description: "YYYY-MM-DD" },
          sec: { type: "string", enum: ["오전간식", "점심", "오후간식"] },
          name: { type: "string" },
          allergy: { type: "array", items: { type: "integer", enum: Object.keys(AN).map(Number) } },
          check: { type: "boolean", description: "부모가 원본과 꼭 대조해야 하는 항목이면 true" },
          why: { type: "string", description: "check가 true인 이유 한 줄. false면 빈 문자열" },
        },
      },
    },
    notes: { type: "array", items: { type: "string" }, description: "전체에 대한 주의사항(연령 혼재, 날짜 없는 특별식 제외 등)" },
  },
};

const SYSTEM = `너는 어린이집 식단표를 구조화하는 도구다. 알레르기가 있는 아이의 부모가 이 결과로 아이가 먹을 수 있는 음식을 고른다. 틀린 알레르기 번호는 아이 건강에 직결되므로, 정확성이 가장 중요하다.

입력은 PDF·한글(HWP)·엑셀에서 뽑은 텍스트다.
- PDF 텍스트는 줄마다 [x=숫자] 좌표가 붙어 있다. 식단표는 보통 요일이 가로 열이고 끼니가 세로 행이므로, x 좌표로 같은 열(같은 날짜)에 속한 메뉴를 묶어라.
- 한글(HWP) 표는 '<표 n>' 안에 '행r | 열a~b(행c~d 병합): 내용' 형식이다. 칸 하나 안의 줄바꿈은 ⏎로 표시했다. 날짜 머리칸(예: '열17~19: 1(목)')과 열 범위가 같은 칸이 그 날짜의 메뉴다. 병합칸은 표시된 행·열 범위 전체를 차지한다. 칸 안에서 ⏎ 다음 줄이 번호만이면(예: '소고기자장밥 ⏎ ①⑤⑥⑩') 바로 앞 메뉴의 알레르기 번호다. 괄호 숫자 표기((5,6))도 알레르기 번호다.

알레르기 번호는 식품의약품안전처 표준이다: ${Object.entries(AN).map(([n, v]) => `${n}.${v}`).join(" ")}.
원본은 번호를 원문자(①②), 괄호 숫자((5,6)), 점 구분(5.6.) 등으로 쓴다. 원문자는 조리법 단계 번호로도 쓰이니, 메뉴명 바로 옆에 붙은 것만 알레르기로 읽어라.

규칙:
- 원본에 적힌 알레르기 번호만 넣는다. 재료로 추정해서 번호를 더하지 않는다. 단, 메뉴명에 재료가 명백한데(예: 달걀찜, 우유) 번호가 없으면 check=true로 표시하고 why에 적는다.
- 확신이 없는 항목(셀이 밀린 듯함, 번호가 어느 메뉴 것인지 애매함, 글자가 깨짐, 날짜 배정이 애매함)은 check=true.
- 끼니는 셋 중 하나로 맞춘다: 오전간식(아침·오전 간식), 점심(중식), 오후간식. 저녁 등 그 밖의 끼니는 빼고 notes에 적는다.
- 메뉴명은 원본 표기 그대로. 알레르기 표기는 이름에서 뺀다.
- 한 칸에 '/'나 '또는'으로 여러 메뉴가 있으면 원본 구분대로 나눈다. '또는'으로 고르는 메뉴는 한 항목으로 둔다.
- 날짜가 없는 특별식(생일식단 등)은 items에 넣지 말고 notes에 적는다. 휴일·빈 칸은 만들지 않는다.
- 여러 연령(1-2세/3-5세)이 섞여 있으면: 연령이 지정됐으면 그 연령만 뽑고 notes에 적는다. 지정되지 않았으면 모든 항목을 그대로 뽑되 전부 check=true, why="여러 연령이 섞인 식단표예요. 우리 아이 반 것인지 확인하세요"로 하고 notes 맨 앞에도 적는다.
- 텍스트에 [떨어진번호]가 붙은 칸은 원문자 알레르기 번호가 줄바꿈되어 메뉴명에서 떨어져 나온 조각이다. 같은 줄 앞 메뉴 것이라고 단정하지 마라 — 표에서 아래·위 줄 메뉴의 번호가 이어진 것일 수 있다. 가장 그럴듯한 메뉴에 붙이되, 그 조각과 같은 줄·바로 위 줄·바로 아래 줄의 메뉴는 모두 check=true로 하고 why에 "떨어진 번호(⑯⑱ 등)가 어느 메뉴 것인지 원본 확인"을 적는다. 조각을 버리지 않는다.
- 조리법은 옮기지 않는다(앱이 따로 붙인다). 원본에 조리법이 섞여 있어도 메뉴명과 알레르기만 뽑는다.`;

export default {
  async fetch(request, env, ctx) {
    const origin = request.headers.get("Origin") || "";
    const allowed = (env.ALLOWED_ORIGINS || "").split(",").map((s) => s.trim()).filter(Boolean);
    const cors = {
      "Access-Control-Allow-Origin": allowed.includes(origin) ? origin : allowed[0] || "",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
      Vary: "Origin",
    };
    if (request.method === "OPTIONS") return new Response(null, { headers: cors });
    const url = new URL(request.url);
    if (url.pathname !== "/parse" || request.method !== "POST") return new Response("not found", { status: 404, headers: cors });
    if (!allowed.includes(origin)) return fail("허용되지 않은 주소에서 온 요청이에요.", 403, cors);
    if (!env.ANTHROPIC_API_KEY) return fail("서버에 AI 키가 아직 설정되지 않았어요.", 503, cors);

    let body;
    try { body = await request.json(); } catch { return fail("잘못된 요청이에요.", 400, cors); }
    const text = String(body.text || "");
    // 반은 묻지 않는다 — 부모가 올리는 파일이 곧 우리 아이 반 식단이다. 예전 앱이 보내는 값만 존중한다
    const age = body.age === "toddler" ? "3-5세(유아반)" : body.age === "infant" ? "1-2세(영아반)" : "";
    const hint = String(body.hint || "").slice(0, 100);
    if (text.trim().length < 50) return fail("파일에서 글자를 거의 못 읽었어요. 스캔 이미지 PDF일 수 있어요.", 400, cors);
    if (text.length > MAX_INPUT_CHARS) return fail("파일이 너무 커요. 한 달치 식단표만 올려 주세요.", 413, cors);

    // 상한: 한 달 전체 + IP별 하루. KV는 즉시 일관되지 않아 몇 건 넘칠 수 있다 — 비용 폭주만 막으면 된다.
    if (env.USAGE) {
      const now = new Date(Date.now() + 9 * 3600e3).toISOString(); // KST
      const mKey = "m:" + now.slice(0, 7);
      const ip = request.headers.get("CF-Connecting-IP") || "unknown";
      const dKey = "d:" + now.slice(0, 10) + ":" + ip;
      const [m, d] = await Promise.all([env.USAGE.get(mKey), env.USAGE.get(dKey)]);
      if (+m >= +(env.MONTHLY_LIMIT || 100)) return fail("이번 달 분석 한도를 다 썼어요. 다음 달 1일에 다시 열려요.", 429, cors);
      if (+d >= +(env.DAILY_PER_IP || 6)) return fail("오늘은 더 올릴 수 없어요. 내일 다시 시도해 주세요.", 429, cors);
      ctx.waitUntil(Promise.all([
        env.USAGE.put(mKey, String(+m + 1), { expirationTtl: 40 * 86400 }),
        env.USAGE.put(dKey, String(+d + 1), { expirationTtl: 2 * 86400 }),
      ]));
    }

    // AI 응답(SSE)을 풀지 않고 그대로 흘려보낸다. 조각마다 풀어 다시 포장하면 무료 요금제 CPU 한도를 넘는다
    // (2026-10-08 실사용 시험에서 "Worker exceeded CPU time limit"). 푸는 일은 브라우저가 한다.
    let upstream;
    try {
      upstream = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "x-api-key": env.ANTHROPIC_API_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-beta": "server-side-fallback-2026-07-01",
          "content-type": "application/json",
        },
        body: JSON.stringify({
          model: MODEL,
          max_tokens: 64000,
          stream: true,
          fallbacks: "default",
          output_config: { effort: "low", format: { type: "json_schema", schema: SCHEMA } },
          system: SYSTEM,
          messages: [{
            role: "user",
            content: `${age ? "연령: " + age + "\n" : ""}${hint ? "어린이집/지역: " + hint + "\n" : ""}\n아래는 식단표 원본 텍스트다.\n\n${text}`,
          }],
        }),
      });
    } catch (e) {
      console.log("upstream fetch error", e && e.message);
      return fail("AI 서버에 연결하지 못했어요. 잠시 뒤 다시 시도해 주세요.", 502, cors);
    }
    if (!upstream.ok) {
      const detail = await upstream.text();
      console.log("upstream status", upstream.status, detail.slice(0, 300));
      const busy = upstream.status === 429 || upstream.status === 529;
      return fail(busy ? "AI가 지금 붐벼요. 잠시 뒤 다시 시도해 주세요." : "분석 중 오류가 났어요. 다시 시도해 주세요.", busy ? 503 : 502, cors);
    }
    return new Response(upstream.body, { headers: { ...cors, "Content-Type": "text/event-stream; charset=utf-8", "Cache-Control": "no-store" } });
  },
};

function fail(msg, status, cors) {
  return new Response(JSON.stringify({ error: msg }) + "\n", { status, headers: { ...cors, "Content-Type": "application/x-ndjson; charset=utf-8" } });
}
