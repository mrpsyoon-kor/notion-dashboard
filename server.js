// 노션 "DSIT 고객" / "현장 업무일지 (N)" 데이터를 그대로 읽어서 보여주는 대시보드 서버
// - 별도 데이터베이스 없음. 노션이 원본(source of truth).
// - 이 서버는 브라우저가 노션 API에 직접 접근할 수 없어서(인증/CORS) 중간에서 대신 조회해주는 역할만 함.
// - /public 폴더의 화면(index.html)이 아래 API를 불러서 화면에 그림.
// - 로그인 없이는 화면/API 둘 다 접근 불가 (세션 쿠키 기반 인증)

const express = require("express");
const session = require("express-session");
const path = require("path");
const ExcelJS = require("exceljs");
require("dotenv").config();

const app = express();
app.use(express.urlencoded({ extended: true })); // 로그인 폼(form) 데이터 파싱용
app.use(express.json());

// 로그인 정보는 반드시 환경변수로 설정하세요 (Render > Environment).
// 아래 값은 환경변수를 깜빡했을 때를 대비한 기본값일 뿐이며, 코드가 공개 저장소에 있다면
// 반드시 Render Variables에 LOGIN_USERNAME / LOGIN_PASSWORD를 따로 설정해서 이 기본값을 덮어써야 합니다.
const LOGIN_USERNAME = process.env.LOGIN_USERNAME || "dsit2024";
const LOGIN_PASSWORD = process.env.LOGIN_PASSWORD || "11111";
const SESSION_SECRET = process.env.SESSION_SECRET || "dsit-dashboard-please-change-this-secret";

app.use(
  session({
    secret: SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      maxAge: 8 * 60 * 60 * 1000, // 8시간 유지, 이후엔 재로그인 필요
    },
  })
);

function requireAuth(req, res, next) {
  if (req.session && req.session.loggedIn) return next();
  return res.redirect("/login");
}

// 로그인 화면
app.get("/login", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "login.html"));
});

// 로그인 처리
app.post("/login", (req, res) => {
  const { username, password } = req.body;
  if (username === LOGIN_USERNAME && password === LOGIN_PASSWORD) {
    req.session.loggedIn = true;
    return res.redirect("/");
  }
  return res.redirect("/login?error=1");
});

// 로그아웃
app.post("/logout", (req, res) => {
  req.session.destroy(() => res.redirect("/login"));
});

// 대시보드 화면 (로그인 필요)
app.get("/", requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// 견적서 생성 화면 (로그인 필요)
app.get("/quote", requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, "public", "quote.html"));
});

// 견적서 생성 - 업로드된 원본 엑셀 템플릿(견적서_Template.xlsx)의 빈 칸(품목/수량/단가/비고)만 채워서
// 그대로 돌려줌. 번호/총액/합계는 템플릿에 이미 있는 수식이 그대로 계산해줌(직접 계산해서 넣지 않음).
app.post("/api/quote/generate", requireAuth, async (req, res) => {
  try {
    const { customerName, items } = req.body;

    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: "품목을 1개 이상 입력해주세요." });
    }
    if (items.length > 10) {
      return res.status(400).json({ error: "이 템플릿은 품목을 최대 10개까지만 담을 수 있습니다." });
    }

    const templatePath = path.join(__dirname, "templates", "견적서_Template.xlsx");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    // 엑셀에서 파일을 열 때 (저장된 캐시값 대신) 모든 수식을 강제로 다시 계산하도록 지정.
    // 이게 없으면 TODAY() 같은 날짜 수식이 예전에 저장했던 값 그대로 보일 수 있음.
    workbook.calcProperties = workbook.calcProperties || {};
    workbook.calcProperties.fullCalcOnLoad = true;
    const sheet = workbook.worksheets[0];

    // 템플릿의 품목 입력 영역은 14행부터 시작 (No./총액은 이미 수식으로 자동 계산됨)
    for (let i = 0; i < 10; i++) {
      const row = 14 + i;
      const it = items[i];
      sheet.getCell(`C${row}`).value = it ? (it.item || "") : "";
      sheet.getCell(`D${row}`).value = it && it.qty !== "" ? Number(it.qty) : null;
      sheet.getCell(`E${row}`).value = it && it.unitPrice !== "" ? Number(it.unitPrice) : null;
      sheet.getCell(`G${row}`).value = it ? (it.note || "") : "";
    }

    const buffer = await workbook.xlsx.writeBuffer();

    const todayStr = new Date().toISOString().slice(0, 10);
    const safeCustomer = (customerName || "고객").replace(/[\\/:*?"<>|]/g, "").trim() || "고객";
    const filename = `견적서_${safeCustomer}_${todayStr}.xlsx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(filename)}"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error("견적서 생성 실패:", err);
    res.status(500).json({ error: "견적서 생성 중 오류가 발생했습니다: " + err.message });
  }
});

// API도 로그인 필요
app.use("/api", requireAuth);

// 그 외 정적 파일(css/js 등). index.html은 위에서 별도 처리하므로 자동 서빙은 꺼둠
app.use(express.static("public", { index: false }));

const NOTION_TOKEN = process.env.NOTION_TOKEN;
const NOTION_VERSION = "2025-09-03";

// 노션 데이터소스 ID (워크스페이스에 이미 있는 실제 DB에서 확인한 값)
const CUSTOMER_DATA_SOURCE_ID = "ba5f6122-d621-4c1a-9dc8-55df939cea33"; // DSIT 고객
const WORKLOG_DATA_SOURCE_ID = "2df557af-30aa-8169-a2bb-000b0d4f2c9c"; // 현장 업무일지 (N)
const BILLING_DATA_SOURCE_ID = "1b6557af-30aa-80e8-bfc6-000b36e4e7b5"; // 청구 내역DB

// 노션 응답은 최대 100개씩 페이지네이션 되므로, 전체를 다 가져올 때까지 반복 조회
async function queryAllPages(dataSourceId) {
  let results = [];
  let cursor = undefined;

  while (true) {
    const response = await fetch(`https://api.notion.com/v1/data_sources/${dataSourceId}/query`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${NOTION_TOKEN}`,
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(cursor ? { start_cursor: cursor, page_size: 100 } : { page_size: 100 }),
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`노션 API 오류 (${response.status}): ${errText}`);
    }

    const data = await response.json();
    results = results.concat(data.results);

    if (!data.has_more) break;
    cursor = data.next_cursor;
  }

  return results;
}

// 노션 속성(property) 객체에서 실제 값만 뽑아내는 헬퍼들
function getText(prop) {
  if (!prop) return "";
  if (prop.type === "title") return prop.title.map((t) => t.plain_text).join("");
  if (prop.type === "rich_text") return prop.rich_text.map((t) => t.plain_text).join("");
  return "";
}
function getSelect(prop) {
  return prop?.select?.name || "";
}
function getStatus(prop) {
  return prop?.status?.name || "";
}
function getMultiSelect(prop) {
  return (prop?.multi_select || []).map((o) => o.name);
}
function getNumber(prop) {
  return typeof prop?.number === "number" ? prop.number : null;
}
function getDate(prop) {
  return prop?.date?.start || null;
}
function getCheckbox(prop) {
  return !!prop?.checkbox;
}
function getPhone(prop) {
  return prop?.phone_number || "";
}
function getRelationIds(prop) {
  return (prop?.relation || []).map((r) => r.id);
}

// 5분 캐시 - 매번 노션에 요청하면 느리고 API 한도에 걸릴 수 있어서 잠깐 저장해둠
let cache = { customers: null, worklogs: null, unpaidOverdue: null, fetchedAt: 0 };
const CACHE_MS = 5 * 60 * 1000;

async function loadData(forceRefresh = false) {
  const isFresh = Date.now() - cache.fetchedAt < CACHE_MS;
  if (!forceRefresh && isFresh && cache.customers && cache.worklogs) {
    return cache;
  }

  const [customerPages, worklogPages, billingPages] = await Promise.all([
    queryAllPages(CUSTOMER_DATA_SOURCE_ID),
    queryAllPages(WORKLOG_DATA_SOURCE_ID),
    queryAllPages(BILLING_DATA_SOURCE_ID),
  ]);

  const allCustomers = customerPages.map((page) => {
    const p = page.properties;
    return {
      id: page.id,
      url: page.url,
      name: getText(p["업체명"]),
      category: getStatus(p["고객 구분"]), // DSIT / 판매 / 계약 종료
      model: getText(p["기종"]),
      serial: getText(p["S/N"]),
      phone: getPhone(p["전화번호"]),
      email: p["이메일"]?.email || "",
      address: getText(p["주소"]),
      owner: getText(p["대표자"]),
      contact: getText(p["담당자"]),
      installDate: getDate(p["설치 일자"]),
      contractEnd: getDate(p["계약 종료일"]),
      baseFee: getNumber(p["기본요금"]),
      deposit: getNumber(p["보증금"]),
      unlimited: getCheckbox(p["무제한"]),
      monthlyBw: getNumber(p["월 사용매수(흑백)"]),
      monthlyColor: getNumber(p["월 사용매수(컬러)"]),
      billingDay: getSelect(p["청구일"]),
      billingMethod: getMultiSelect(p["청구방식"]),
      vatStatus: getSelect(p["VAT 여부"]),
      note: getText(p["비고란"]),
    };
  });

  // 대시보드에는 현재 사용 중인 DSIT 고객만 노출, 청구일(숫자) 오름차순 정렬
  const customers = allCustomers
    .filter((c) => c.category === "DSIT")
    .sort((a, b) => {
      const da = parseInt(a.billingDay, 10);
      const db = parseInt(b.billingDay, 10);
      const va = isNaN(da) ? 999 : da;
      const vb = isNaN(db) ? 999 : db;
      return va - vb;
    });

  const worklogs = worklogPages.map((page) => {
    const p = page.properties;
    return {
      id: page.id,
      url: page.url,
      customerName: getText(p["고객명"]),
      date: getDate(p["날짜"]),
      pic: getSelect(p["담당자PIC"]),
      category: getSelect(p["업무 구분"]),
      tasks: getMultiSelect(p["작업내역"]),
      memo: getText(p["메모"]),
      supplyStock: getText(p["소모품 재고(K/C/M/Y/W)"]),
      dsitAccounting: getCheckbox(p["DSIT 회계"]),
    };
  });

  // 고객 ID -> 고객정보 매핑 (청구내역의 "DSIT 고객" 관계연결로 고객을 찾기 위함).
  // 미입금 현황은 지금 DSIT로 분류된 고객뿐 아니라, 과거 청구건이라면 분류가 바뀐 고객도 나올 수 있어 전체(allCustomers) 기준으로 매핑.
  const customerById = new Map(allCustomers.map((c) => [c.id, c]));

  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  const now = Date.now();

  const billingRecords = billingPages.map((page) => {
    const p = page.properties;
    const relatedIds = getRelationIds(p["DSIT 고객"]);
    const customer = relatedIds.length > 0 ? customerById.get(relatedIds[0]) : null;
    return {
      id: page.id,
      url: page.url,
      createdAt: page.created_time, // 실제 "청구일" 필드가 거의 비어있어, 청구건이 생성된 날짜를 청구일 대용으로 사용
      status: getStatus(p["청구 상태"]),
      paid: getCheckbox(p["입금 완료"]),
      customerId: relatedIds[0] || null,
      customerName: customer ? customer.name : null,
      customerPhone: customer ? customer.phone : null,
      customerContact: customer ? customer.contact : null,
      baseFee: customer ? customer.baseFee : null,
    };
  });

  // 입금 미완료 + 청구(생성)일로부터 30일 이상 경과한 건만 추림
  const unpaidOverdue = billingRecords
    .filter((b) => !b.paid && b.createdAt)
    .map((b) => ({ ...b, daysOverdue: Math.floor((now - new Date(b.createdAt).getTime()) / ONE_DAY_MS) }))
    .filter((b) => b.daysOverdue >= 30)
    .sort((a, b) => b.daysOverdue - a.daysOverdue);

  cache = { customers, worklogs, unpaidOverdue, fetchedAt: Date.now() };
  return cache;
}

app.get("/api/dashboard", async (req, res) => {
  try {
    if (!NOTION_TOKEN) {
      return res.status(500).json({ error: "서버에 NOTION_TOKEN이 설정되어 있지 않습니다." });
    }
    const forceRefresh = req.query.refresh === "1";
    const data = await loadData(forceRefresh);
    res.json({ customers: data.customers, worklogs: data.worklogs, unpaidOverdue: data.unpaidOverdue, fetchedAt: data.fetchedAt });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`대시보드 서버 실행 중: http://localhost:${PORT}`));
