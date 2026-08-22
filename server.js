// 노션 "DSIT 고객" / "현장 업무일지 (N)" 데이터를 그대로 읽어서 보여주는 대시보드 서버
// - 별도 데이터베이스 없음. 노션이 원본(source of truth).
// - 이 서버는 브라우저가 노션 API에 직접 접근할 수 없어서(인증/CORS) 중간에서 대신 조회해주는 역할만 함.
// - /public 폴더의 화면(index.html)이 아래 API를 불러서 화면에 그림.

const express = require("express");
require("dotenv").config();

const app = express();
app.use(express.static("public"));

const NOTION_TOKEN = process.env.NOTION_TOKEN;
const NOTION_VERSION = "2025-09-03";

// 노션 데이터소스 ID (워크스페이스에 이미 있는 실제 DB에서 확인한 값)
const CUSTOMER_DATA_SOURCE_ID = "ba5f6122-d621-4c1a-9dc8-55df939cea33"; // DSIT 고객
const WORKLOG_DATA_SOURCE_ID = "2df557af-30aa-8169-a2bb-000b0d4f2c9c"; // 현장 업무일지 (N)

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

// 5분 캐시 - 매번 노션에 요청하면 느리고 API 한도에 걸릴 수 있어서 잠깐 저장해둠
let cache = { customers: null, worklogs: null, fetchedAt: 0 };
const CACHE_MS = 5 * 60 * 1000;

async function loadData(forceRefresh = false) {
  const isFresh = Date.now() - cache.fetchedAt < CACHE_MS;
  if (!forceRefresh && isFresh && cache.customers && cache.worklogs) {
    return cache;
  }

  const [customerPages, worklogPages] = await Promise.all([
    queryAllPages(CUSTOMER_DATA_SOURCE_ID),
    queryAllPages(WORKLOG_DATA_SOURCE_ID),
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

  cache = { customers, worklogs, fetchedAt: Date.now() };
  return cache;
}

app.get("/api/dashboard", async (req, res) => {
  try {
    if (!NOTION_TOKEN) {
      return res.status(500).json({ error: "서버에 NOTION_TOKEN이 설정되어 있지 않습니다." });
    }
    const forceRefresh = req.query.refresh === "1";
    const data = await loadData(forceRefresh);
    res.json({ customers: data.customers, worklogs: data.worklogs, fetchedAt: data.fetchedAt });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`대시보드 서버 실행 중: http://localhost:${PORT}`));
