// POST /parse
// 부모가 올린 식단표에서 브라우저가 뽑은 텍스트를 받아, Claude로 날짜·끼니·메뉴·알레르기를 구조화한다.
// 사진·스캔 PDF는 글자가 없으니 브라우저가 줄인 그림(images)을 보내고 Claude 비전으로 읽는다(2026-10-13).
// 저장은 하지 않는다 — 결과는 부모 기기 localStorage에만 남는다.
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
// 그림: 한 달치 식단표는 1~2쪽. 브라우저가 긴 변 2000px JPEG로 줄여 보내므로 한 장 1MB면 넉넉하다
const MAX_IMAGES = 4;
const MAX_IMAGE_B64 = 1_500_000;

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
- 메뉴 옆 괄호에 적힌 대체메뉴(예: '동태살양념찜 ⏎ (닭채소조림(5,6,15))', '참치당근죽 (두부당근죽)', '자른포도(토마토)')는 별도 항목으로 나눈다. 이름은 '닭채소조림 (대체)'처럼 뒤에 ' (대체)'를 붙이고, 대체메뉴 자기 알레르기 번호를 넣는다. 원래 메뉴에는 원래 메뉴 번호만 남긴다. 대체메뉴가 있다는 이유만으로 check=true로 하지 않는다.
- 날짜가 없는 특별식(생일식단 등)은 items에 넣지 말고 notes에 적는다. 휴일·빈 칸은 만들지 않는다.
- 여러 연령(1-2세/3-5세)이 섞여 있으면: 연령이 지정됐으면 그 연령만 뽑고 notes에 적는다. 지정되지 않았으면 모든 항목을 그대로 뽑되 전부 check=true, why="여러 연령이 섞인 식단표예요. 우리 아이 반 것인지 확인하세요"로 하고 notes 맨 앞에도 적는다.
- 텍스트에 [떨어진번호]가 붙은 칸은 원문자 알레르기 번호가 줄바꿈되어 메뉴명에서 떨어져 나온 조각이다. 같은 줄 앞 메뉴 것이라고 단정하지 마라 — 표에서 아래·위 줄 메뉴의 번호가 이어진 것일 수 있다. 가장 그럴듯한 메뉴에 붙이되, 그 조각과 같은 줄·바로 위 줄·바로 아래 줄의 메뉴는 모두 check=true로 하고 why에 "떨어진 번호(⑯⑱ 등)가 어느 메뉴 것인지 원본 확인"을 적는다. 조각을 버리지 않는다.
- 조리법은 옮기지 않는다(앱이 따로 붙인다). 원본에 조리법이 섞여 있어도 메뉴명과 알레르기만 뽑는다.
- 한 칸 안의 'A vs B'(자율선택급식 등)는 고르는 메뉴이므로 '또는'처럼 한 항목으로 두고, 양쪽 번호를 합쳐 넣는다.
- 영양 정보(kcal·단백질 등 숫자 줄), 원산지표, QR 안내, 반찬 없는 행사 문구는 메뉴가 아니다.`;

// 그림으로 받았을 때만 덧붙이는 규칙. 사진은 텍스트보다 오독 위험이 크다
const SYSTEM_IMAGE = `

입력이 식단표 그림(휴대폰 사진·스캔)이다.
- 표의 칸 경계를 먼저 잡고, 날짜 머리칸 아래 같은 칸에 든 메뉴만 그 날짜로 묶어라. 사진이 기울었거나 휘어 있어도 칸 단위로 읽는다.
- 알레르기 번호는 숫자 하나하나를 확인한다. 흐리거나, 잘렸거나, 그림자·반사로 가려졌거나, 비슷한 숫자(1과 7, 3과 8, 5와 6, 13과 18 등)로 헷갈릴 여지가 조금이라도 있으면 읽은 대로 넣되 check=true, why에 "번호가 흐려요. 원본에서 (N) 확인" 식으로 적는다. 메뉴명 글자가 불확실해도 check=true.
- 사진 밖으로 잘려 안 보이는 날짜·칸이 있으면 지어내지 말고 notes에 "N일 칸이 사진에서 잘렸어요"처럼 적는다.
- 식단표에 알레르기 범례(①난류 ②우유 …)가 있으면 위 표준 번호와 하나씩 비교하라. 표기만 다르고 같은 뜻(아황산염/아황산류, 계란/난류)은 차이가 아니다. 어느 번호든 뜻이 다르면 notes 맨 앞에 "이 식단표는 ⑱=전복, ⑲=홍합으로 적었어요(표준은 ⑱조개류, ⑲잣). 앱은 표준 번호로 보여 줘요"처럼 적는다. 번호는 원본 그대로 둔다.
  · 원본 뜻이 표준 뜻 안에 들어가면(⑱전복은 조개류의 하나) notes에만 적고 항목은 check하지 않는다. 표준으로 보여 줘도 빠뜨리는 쪽으로 틀리지 않는다.
  · 들어가지 않으면(⑲홍합인데 표준 ⑲은 잣) 그 번호가 붙은 항목을 모두 check=true, why에 "원본 ⑲는 홍합이에요(앱에선 잣으로 보여요)"처럼 적는다.`;

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
    const images = Array.isArray(body.images) ? body.images : [];
    if (images.length > MAX_IMAGES) return fail(`사진은 ${MAX_IMAGES}장까지 올릴 수 있어요.`, 413, cors);
    for (const im of images) {
      if (!im || !/^image\/(jpeg|png|webp)$/.test(im.media_type) || typeof im.data !== "string") return fail("사진 형식을 읽지 못했어요.", 400, cors);
      if (im.data.length > MAX_IMAGE_B64) return fail("사진이 너무 커요. 다시 시도해 주세요.", 413, cors);
    }
    if (!images.length && text.trim().length < 50) return fail("파일에서 글자를 거의 못 읽었어요. 스캔 이미지 PDF일 수 있어요.", 400, cors);
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
          system: images.length ? SYSTEM + SYSTEM_IMAGE : SYSTEM,
          messages: [{
            role: "user",
            content: images.length
              ? [
                  ...images.map((im) => ({ type: "image", source: { type: "base64", media_type: im.media_type, data: im.data } })),
                  { type: "text", text: `${age ? "연령: " + age + "\n" : ""}${hint ? "어린이집/지역: " + hint + "\n" : ""}위 그림은 식단표 원본이다(${images.length}장).${text.trim() ? "\n\n함께 올린 다른 파일의 텍스트:\n\n" + text : ""}` },
                ]
              : `${age ? "연령: " + age + "\n" : ""}${hint ? "어린이집/지역: " + hint + "\n" : ""}\n아래는 식단표 원본 텍스트다.\n\n${text}`,
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
