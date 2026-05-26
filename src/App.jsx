import React, { useState, useMemo, useCallback, useRef } from 'react';
import {
  LucideUpload, LucideFileSpreadsheet, LucideDownload, LucideSparkles,
  LucideTrendingUp, LucideTrendingDown, LucideStar, LucideCalendar, LucideShirt, LucideTag,
  LucidePackage, LucideAlertCircle, LucideX, LucideImage, LucideBox, LucideArchive, LucideAward,
  LucideMessageSquare, LucideKey, LucideLoader2, LucideCalendarCheck, LucideSearch
} from 'lucide-react';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const THEMES = [
  { id: 'recommend', label: '추천 (통합 8개)', desc: '베스트·신상·급상승 분산 픽', icon: LucideAward, modes: ['xls'] },
  { id: 'bestseller', label: '베스트셀러', desc: '총 판매량 Top 8', icon: LucideStar, modes: ['xls', 'csv'] },
  { id: 'newProduct', label: '신상품 베스트', desc: '최근 N개월 등록 + 판매량', icon: LucideSparkles, modes: ['xls', 'csv'] },
  { id: 'brand', label: '브랜드별 베스트', desc: '상품명 prefix 코드로 분류', icon: LucideTag, modes: ['xls', 'csv'] },
  { id: 'category', label: '카테고리별 베스트', desc: '카테고리 선택 → Top 8', icon: LucideShirt, modes: ['xls', 'csv'] },
  { id: 'steady', label: '스테디셀러', desc: '오래됐지만 꾸준한 상품', icon: LucidePackage, modes: ['xls', 'csv'] },
  { id: 'frequency', label: '판매 빈도', desc: '판매일수 많은 안정 상품', icon: LucideCalendarCheck, modes: ['xls'] },
  { id: 'rising', label: '급상승(라이징)', desc: '초반 10% vs 후반 10% 증가율', icon: LucideTrendingUp, modes: ['xls'] },
  { id: 'overstock', label: '재고 과다', desc: '재고 많고 안 팔리는 상품 (재고 소진용)', icon: LucideArchive, modes: ['xls'] },
  { id: 'declining', label: '판매 감소', desc: '후반에 판매가 떨어진 상품 (재활성용)', icon: LucideTrendingDown, modes: ['xls'] },
  { id: 'package', label: '패키지 베스트', desc: '상품명에 PACK 포함된 상품 Top 8', icon: LucideBox, modes: ['csv'] },
  { id: 'custom', label: '직접 입력', desc: '자연어 조건으로 직접 추출', icon: LucideMessageSquare, modes: ['xls', 'csv'] },
];

const isPackage = (productName) => /PACK/i.test(productName || '');
const resolveCategory = (productName, rawCategory) =>
  isPackage(productName) ? '패키지' : (rawCategory || '미분류');

const PRODUCT_CODE_RE = /[A-Z]{4}\d{4}/;
const extractProductCode = (productName) => {
  const m = String(productName || '').match(PRODUCT_CODE_RE);
  return m ? m[0] : null;
};

const enrichPackageSeasons = (groups) => {
  const codeToSeason = new Map();
  for (const g of groups) {
    if (g.season && !isPackage(g.productName)) {
      const code = extractProductCode(g.productName);
      if (code && !codeToSeason.has(code)) codeToSeason.set(code, g.season);
    }
  }
  for (const g of groups) {
    if (!g.season && isPackage(g.productName)) {
      const code = extractProductCode(g.productName);
      if (code && codeToSeason.has(code)) g.season = codeToSeason.get(code);
    }
  }
  return groups;
};

const BRAND_CODES = ['FP', 'JM', 'WV', 'PS', 'EZ', 'TWN', 'PL', 'DY'];

const extractBrand = (productName) => {
  if (!productName) return '';
  const tokens = productName.split(/[\s\[\]()]+/).filter(Boolean);
  for (const tok of tokens) {
    if (BRAND_CODES.includes(tok)) return tok;
  }
  return '';
};

const parsePrice = (v) => {
  if (!v) return 0;
  const n = parseInt(String(v).replace(/[^0-9]/g, ''), 10);
  return isNaN(n) ? 0 : n;
};

const parseInt0 = (v) => {
  const n = parseInt(String(v ?? '').trim(), 10);
  return isNaN(n) ? 0 : n;
};

const SEASON_VALUES = ['S/S', 'F/W', '사계절'];
const extractSeason = (purchaseName) => {
  if (!purchaseName) return '';
  if (/사계절|ALL\s*SEASON|ALLSEASON/i.test(purchaseName)) return '사계절';
  const m = purchaseName.match(/(S\/S|F\/W)/i);
  return m ? m[1].toUpperCase() : '';
};

const cleanCellText = (v) => {
  if (!v) return '';
  return String(v).replace(/^=?"+|"+$/g, '').trim();
};

const parseHtmlXls = (htmlText) => {
  const doc = new DOMParser().parseFromString(htmlText, 'text/html');
  const rows = Array.from(doc.querySelectorAll('table tr'));
  if (rows.length < 2) throw new Error('테이블에 데이터가 없습니다.');

  const headerCells = Array.from(rows[0].querySelectorAll('td, th')).map(c => c.textContent.trim());

  const idx = (name) => headerCells.findIndex(h => h === name);
  const inDateCols = headerCells
    .map((h, i) => h.endsWith(' 입고') ? i : -1)
    .filter(i => i >= 0 && headerCells[i] !== '입고합계수량');
  const saleDateCols = headerCells
    .map((h, i) => h.endsWith(' 판매') ? i : -1)
    .filter(i => i >= 0 && headerCells[i] !== '판매합계수량');
  const dateLabels = saleDateCols.map(i => headerCells[i].replace(' 판매', ''));

  const colMap = {
    image: idx('대표이미지'),
    productName: idx('상품명'),
    barcode: idx('바코드번호'),
    productCode: idx('상품코드'),
    purchaseName: idx('사입상품명'),
    category: idx('상품분류명'),
    memo1: idx('메모1'),
    memo2: idx('메모2'),
    supplier: idx('공급처명'),
    optionName: idx('옵션명'),
    registDate: idx('등록일자'),
    cost: idx('원가'),
    price: idx('판매단가'),
    amount: idx('금액'),
    inStockTotal: idx('입고합계수량'),
    salesTotal: idx('판매합계수량'),
    currentStock: idx('현재재고'),
    unshipped: idx('미발송수'),
    canceled: idx('취소수량'),
  };

  const skus = [];
  for (let r = 1; r < rows.length; r++) {
    const cells = Array.from(rows[r].querySelectorAll('td'));
    if (cells.length < headerCells.length - 2) continue;
    const get = (i) => i >= 0 && cells[i] ? cells[i].textContent.trim() : '';

    let imageUrl = null;
    if (colMap.image >= 0 && cells[colMap.image]) {
      const img = cells[colMap.image].querySelector('img');
      if (img && img.getAttribute('src')) imageUrl = img.getAttribute('src');
    }

    const productName = get(colMap.productName);
    if (!productName) continue;

    const sales = saleDateCols.map(i => parseInt0(get(i)));
    const inStock = inDateCols.map(i => parseInt0(get(i)));
    const purchaseName = cleanCellText(get(colMap.purchaseName));

    skus.push({
      imageUrl,
      productName,
      barcode: get(colMap.barcode),
      productCode: cleanCellText(get(colMap.productCode)),
      purchaseName,
      season: extractSeason(purchaseName),
      category: resolveCategory(productName, get(colMap.category)),
      memo1: get(colMap.memo1),
      memo2: cleanCellText(get(colMap.memo2)),
      supplier: get(colMap.supplier),
      optionName: get(colMap.optionName),
      registDate: get(colMap.registDate),
      cost: parsePrice(get(colMap.cost)),
      price: parsePrice(get(colMap.price)),
      amount: parsePrice(get(colMap.amount)),
      inStock,
      inStockTotal: parseInt0(get(colMap.inStockTotal)),
      sales,
      salesTotal: parseInt0(get(colMap.salesTotal)),
      currentStock: parseInt0(get(colMap.currentStock)),
      unshipped: parseInt0(get(colMap.unshipped)),
      canceled: parseInt0(get(colMap.canceled)),
    });
  }

  return { skus, dateLabels };
};

const parseCsvLine = (line) => {
  const result = [];
  let cur = '';
  let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (c === ',' && !inQ) {
      result.push(cur);
      cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur);
  return result;
};

const parseCsvRows = (text) => {
  const rows = [];
  let row = [], cur = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') {
        if (text[i + 1] === '"') { cur += '"'; i++; }
        else inQ = false;
      } else cur += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === ',') { row.push(cur); cur = ''; }
      else if (c === '\r') { /* skip */ }
      else if (c === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
      else cur += c;
    }
  }
  if (cur.length || row.length) { row.push(cur); rows.push(row); }
  return rows;
};

const isAdListCsv = (text) => {
  const head = text.slice(0, 600);
  return head.includes('광고 이름') && head.includes('상품리스트');
};

const parseAdList = (text) => {
  const rows = parseCsvRows(text);
  if (rows.length < 3) throw new Error('광고리스트 데이터가 없습니다.');
  const header = rows[0].map(h => (h || '').trim());
  const idx = (name) => header.findIndex(h => h === name);
  const cNo = idx('No') >= 0 ? idx('No') : 0;
  const cName = idx('광고 이름') >= 0 ? idx('광고 이름') : 1;
  const cStatus = idx('상태');
  const cProduct = idx('상품리스트');
  const cCount = idx('횟수');
  if (cProduct < 0) throw new Error('상품리스트 컬럼을 찾을 수 없습니다.');

  const campaigns = [];
  let cur = null;
  for (let r = 2; r < rows.length; r++) {
    const row = rows[r];
    const get = (i) => i >= 0 && row[i] != null ? String(row[i]).trim() : '';
    const no = get(cNo);
    const name = get(cName);
    const prod = get(cProduct);
    if (no && name) {
      const m = name.match(/\((\d{3,4})\)\s*$/);
      cur = {
        no, name,
        manager: (name.split('_')[0] || '').trim(),
        status: get(cStatus),
        postCode: m ? m[1].padStart(4, '0') : '',
        products: [],
      };
      campaigns.push(cur);
    }
    if (cur && prod) {
      const codes = (prod.match(/[A-Z]{3,5}\d{3,5}/g) || []);
      cur.products.push({
        raw: prod,
        codes,
        count: parseInt0(get(cCount)) || 1,
      });
    }
  }
  if (campaigns.length === 0) throw new Error('캠페인을 찾을 수 없습니다.');
  return campaigns;
};

const decodeXml = (s) => String(s || '')
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'")
  .replace(/&amp;/g, '&');

const xlsxBuildRichMap = async (readText) => {
  const relsText = await readText('xl/richData/_rels/richValueRel.xml.rels');
  if (!relsText) return null;
  const relMap = {};
  for (const m of relsText.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)) relMap[m[1]] = m[2];
  const rvrText = await readText('xl/richData/richValueRel.xml');
  if (!rvrText) return null;
  const rvrels = [...rvrText.matchAll(/r:id="([^"]+)"/g)].map(m => m[1]);
  const rdvText = await readText('xl/richData/rdrichvalue.xml');
  if (!rdvText) return null;
  const rvRelIdx = [];
  for (const m of rdvText.matchAll(/<rv\b[^>]*>([\s\S]*?)<\/rv>/g)) {
    const vs = [...m[1].matchAll(/<v>([^<]*)<\/v>/g)];
    rvRelIdx.push(vs.length ? parseInt(vs[0][1], 10) : -1);
  }
  const metaText = await readText('xl/metadata.xml');
  if (!metaText) return null;
  const fmMatch = metaText.match(/<futureMetadata name="XLRICHVALUE"[\s\S]*?<\/futureMetadata>/);
  const fmRvb = [];
  if (fmMatch) {
    for (const m of fmMatch[0].matchAll(/<bk>([\s\S]*?)<\/bk>/g)) {
      const rm = m[1].match(/rvb i="(\d+)"/);
      fmRvb.push(rm ? parseInt(rm[1], 10) : -1);
    }
  }
  const vmMatch = metaText.match(/<valueMetadata[\s\S]*?<\/valueMetadata>/);
  const vmFmIdx = [];
  if (vmMatch) {
    for (const m of vmMatch[0].matchAll(/<bk>([\s\S]*?)<\/bk>/g)) {
      const rm = m[1].match(/<rc [^>]*v="(\d+)"/);
      vmFmIdx.push(rm ? parseInt(rm[1], 10) : -1);
    }
  }
  return (vmVal) => {
    const fmi = vmFmIdx[vmVal - 1];
    if (fmi == null) return null;
    const rvi = fmRvb[fmi];
    if (rvi == null) return null;
    const reli = rvRelIdx[rvi];
    if (reli == null) return null;
    const rid = rvrels[reli];
    const target = relMap[rid];
    if (!target) return null;
    return 'xl/' + target.replace('../', '');
  };
};

const isAdListXlsx = async (zip) => {
  const f = zip.file('xl/workbook.xml');
  if (!f) return false;
  const wbXml = await f.async('string');
  return /name="[^"]*광고리스트[^"]*"/.test(wbXml);
};

const parseAdListXlsx = async (zip) => {
  const readText = async (p) => {
    const f = zip.file(p);
    return f ? await f.async('string') : null;
  };

  const wbXml = await readText('xl/workbook.xml');
  const wbRels = await readText('xl/_rels/workbook.xml.rels');
  if (!wbXml || !wbRels) throw new Error('워크북을 읽을 수 없습니다.');
  const ridMap = {};
  for (const m of wbRels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)) ridMap[m[1]] = m[2];
  const sheets = [...wbXml.matchAll(/<sheet [^>]*name="([^"]*)"[^>]*r:id="([^"]*)"/g)];
  let target = null;
  for (const [, name, rid] of sheets) {
    if (name === 'SNS광고리스트') target = ridMap[rid];
  }
  if (!target) {
    for (const [, name, rid] of sheets) {
      if (name.includes('광고리스트') && !name.includes('전주')) { target = ridMap[rid]; break; }
    }
  }
  if (!target) throw new Error('광고리스트 시트를 찾을 수 없습니다.');
  const sx = await readText('xl/' + target.replace('../', ''));
  if (!sx) throw new Error('시트를 읽을 수 없습니다.');

  const ssText = await readText('xl/sharedStrings.xml');
  const sst = [];
  if (ssText) {
    for (const m of ssText.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      let t = '';
      for (const tm of m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)) t += tm[1];
      sst.push(decodeXml(t));
    }
  }

  const vmToMedia = await xlsxBuildRichMap(readText);
  const mediaCache = {};
  const mediaUrl = async (mp) => {
    if (!mp) return null;
    if (mediaCache[mp]) return mediaCache[mp];
    const f = zip.file(mp);
    if (!f) return null;
    const blob = await f.async('blob');
    const url = URL.createObjectURL(blob);
    mediaCache[mp] = url;
    return url;
  };

  const campaigns = [];
  let cur = null;
  for (const rm of sx.matchAll(/<row [^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const rowXml = rm[2];
    const cells = {};
    const cellVm = {};
    for (const cm of rowXml.matchAll(/<c [^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) {
      const ctag = cm[0];
      const refM = ctag.match(/r="([A-Z]+)\d+"/);
      if (!refM) continue;
      const col = refM[1];
      const vmM = ctag.match(/vm="(\d+)"/);
      if (vmM) cellVm[col] = parseInt(vmM[1], 10);
      const isStr = /t="s"/.test(ctag);
      const vM = ctag.match(/<v>([^<]*)<\/v>/);
      if (vM) {
        cells[col] = isStr ? (sst[parseInt(vM[1], 10)] || '') : vM[1];
      } else {
        const isM = ctag.match(/<is>[\s\S]*?<t[^>]*>([^<]*)<\/t>/);
        cells[col] = isM ? decodeXml(isM[1]) : '';
      }
    }
    const no = (cells['A'] || '').trim();
    const name = (cells['B'] || '').trim();
    const prod = (cells['G'] || '').trim();
    if (no && name) {
      const pm = name.match(/\((\d{3,4})\)\s*$/);
      cur = {
        no, name,
        manager: (name.split('_')[0] || '').trim(),
        status: (cells['C'] || '').trim(),
        postCode: pm ? pm[1].padStart(4, '0') : '',
        products: [],
      };
      campaigns.push(cur);
    }
    if (cur && prod) {
      const codes = prod.match(/[A-Z]{3,5}\d{3,5}/g) || [];
      // H열 = 그 상품 행의 광고 이미지 (행 단위로 G상품·H이미지가 짝)
      cur.products.push({ raw: prod, codes, thumbVm: cellVm['H'] || null, thumbUrl: null });
    }
  }

  if (vmToMedia) {
    for (const c of campaigns) {
      for (const p of c.products) {
        if (p.thumbVm) {
          const mp = vmToMedia(p.thumbVm);
          p.thumbUrl = await mediaUrl(mp);
        }
      }
    }
  }
  if (campaigns.length === 0) throw new Error('캠페인을 찾을 수 없습니다.');
  return campaigns;
};

const parseProductCsv = (text) => {
  const lines = text.split(/\r?\n/).filter(l => l.length > 0);
  if (lines.length < 2) throw new Error('데이터 행이 없습니다.');
  const headers = parseCsvLine(lines[0]);
  const idx = (name) => headers.findIndex(h => h.trim() === name);

  const colMap = {
    productName: idx('상품명'),
    category: idx('상품분류명'),
    purchaseName: idx('사입상품명'),
    memo1: idx('상품메모1'),
    registDate: idx('상품등록일자'),
    price: idx('평균판매가'),
    quantity: idx('수량'),
    avgDaily: idx('일평균수량'),
    revenue: idx('금액'),
    currentStock: idx('현재재고'),
    cost: idx('상품원가'),
    productCode: idx('상품코드'),
  };
  if (colMap.productName < 0) throw new Error('상품명 컬럼을 찾지 못했습니다.');

  const groups = [];
  for (let r = 1; r < lines.length; r++) {
    const cells = parseCsvLine(lines[r]);
    const get = (i) => i >= 0 && cells[i] != null ? cells[i].trim() : '';
    const productName = get(colMap.productName);
    if (!productName) continue;
    const purchaseName = get(colMap.purchaseName);
    const totalSales = parseInt0(get(colMap.quantity));
    const avgDaily = parseFloat(get(colMap.avgDaily)) || 0;
    groups.push({
      productName,
      category: resolveCategory(productName, get(colMap.category)),
      season: extractSeason(purchaseName),
      memo1: get(colMap.memo1),
      brand: extractBrand(productName),
      isPackage: isPackage(productName),
      registDate: get(colMap.registDate),
      price: parsePrice(get(colMap.price)),
      imageUrl: null,
      skus: [],
      totalSales,
      totalRevenue: parsePrice(get(colMap.revenue)),
      totalCanceled: 0,
      totalCurrentStock: parseInt0(get(colMap.currentStock)),
      earlySales: 0,
      lateSales: avgDaily,
      growthRate: avgDaily,
      cancelRate: 0,
      bestSku: null,
      bestSku2: null,
      avgDaily,
    });
  }
  return enrichPackageSeasons(groups);
};

const adNum = (v) => {
  if (v == null) return 0;
  if (typeof v === 'object') v = v.result ?? v.text ?? v.richText?.map(t => t.text).join('') ?? '';
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
};

const adText = (v) => {
  if (v == null) return '';
  if (typeof v === 'object') v = v.result ?? v.text ?? v.richText?.map(t => t.text).join('') ?? '';
  return String(v).trim();
};

const AGE_GROUPS = [
  { id: '신규', label: '신규', range: '≤14일', test: d => d != null && d <= 14 },
  { id: '중기', label: '중기', range: '15~60일', test: d => d != null && d > 14 && d <= 60 },
  { id: '노후', label: '노후', range: '>60일', test: d => d != null && d > 60 },
];

const ageGroupOf = (days) => {
  for (const g of AGE_GROUPS) if (g.test(days)) return g.id;
  return '미상';
};

const parsePostDate = (postCode, reportEnd) => {
  if (!postCode || postCode.length !== 4) return null;
  const mm = parseInt(postCode.slice(0, 2), 10);
  const dd = parseInt(postCode.slice(2), 10);
  if (!(mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31)) return null;
  const end = reportEnd ? new Date(reportEnd) : new Date();
  if (isNaN(end.getTime())) return null;
  let year = end.getFullYear();
  let d = new Date(year, mm - 1, dd);
  if (d.getTime() > end.getTime()) d = new Date(year - 1, mm - 1, dd);
  return d;
};

const splitProductId = (raw) => {
  const s = adText(raw);
  const idx = s.indexOf(',');
  if (idx < 0) return { id: s, name: s };
  return { id: s.slice(0, idx).trim(), name: s.slice(idx + 1).trim() };
};

const parseAdRawByProduct = (ws, headers) => {
  const findCol = (pred) => headers.findIndex(h => h && pred(h));
  const col = {
    name: findCol(h => h.includes('광고 이름')),
    product: findCol(h => h.includes('제품 ID')),
    roas: findCol(h => h.toUpperCase().includes('ROAS')),
    spend: findCol(h => h.includes('지출')),
    purchases: findCol(h => h.trim() === '구매'),
    clicks: findCol(h => h.includes('링크 클릭')),
    impressions: findCol(h => h.trim() === '노출'),
    body: findCol(h => h.includes('본문')),
    start: findCol(h => h.includes('보고 시작')),
    end: findCol(h => h.includes('보고 종료')),
  };
  if (col.name < 0 || col.product < 0) {
    throw new Error('광고 이름·제품 ID 컬럼을 찾을 수 없습니다.');
  }

  const lastRow = ws.rowCount || ws.actualRowCount || 1;
  const campMap = new Map();
  let minStart = '', maxEnd = '';

  for (let r = 2; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const cell = (i) => i >= 0 ? row.getCell(i + 1).value : null;
    const name = adText(cell(col.name));
    if (!name) continue;
    const start = adText(cell(col.start));
    const end = adText(cell(col.end));
    if (start && (!minStart || start < minStart)) minStart = start;
    if (end && (!maxEnd || end > maxEnd)) maxEnd = end;

    if (!campMap.has(name)) {
      campMap.set(name, {
        name, products: new Map(),
        spend: 0, impressions: 0, clicks: 0, revenue: 0, purchases: 0, body: '',
      });
    }
    const camp = campMap.get(name);
    const spend = adNum(cell(col.spend));
    const impressions = adNum(cell(col.impressions));
    const clicks = adNum(cell(col.clicks));
    const roas = adNum(cell(col.roas));
    const purchases = adNum(cell(col.purchases));
    camp.spend += spend;
    camp.impressions += impressions;
    camp.clicks += clicks;
    camp.revenue += roas * spend;
    camp.purchases += purchases;
    if (!camp.body && col.body >= 0) camp.body = adText(cell(col.body));

    const { id, name: pname } = splitProductId(cell(col.product));
    const key = id || pname;
    if (!camp.products.has(key)) {
      camp.products.set(key, {
        productId: id, productName: pname,
        spend: 0, impressions: 0, clicks: 0, revenue: 0, purchases: 0,
      });
    }
    const p = camp.products.get(key);
    p.spend += spend;
    p.impressions += impressions;
    p.clicks += clicks;
    p.revenue += roas * spend;
    p.purchases += purchases;
  }

  const endD = maxEnd ? new Date(maxEnd) : new Date();
  const campaigns = [];
  for (const camp of campMap.values()) {
    const products = [...camp.products.values()].map(p => ({
      ...p,
      ctr: p.impressions > 0 ? p.clicks / p.impressions : 0,
      cpc: p.clicks > 0 ? p.spend / p.clicks : 0,
      roas: p.spend > 0 ? p.revenue / p.spend : 0,
    })).sort((a, b) => b.spend - a.spend);

    const dateM = camp.name.match(/\((\d{3,4})\)\s*$/);
    const postCode = dateM ? dateM[1].padStart(4, '0') : '';
    const postDate = parsePostDate(postCode, maxEnd);
    const ageDays = postDate && !isNaN(endD.getTime())
      ? Math.max(0, Math.round((endD.getTime() - postDate.getTime()) / 86400000))
      : null;

    campaigns.push({
      name: camp.name,
      manager: (camp.name.split('_')[0] || '-').trim(),
      postCode, ageDays, ageGroup: ageGroupOf(ageDays),
      status: '',
      budget: 0,
      spend: camp.spend,
      revenue: camp.revenue,
      roas: camp.spend > 0 ? camp.revenue / camp.spend : 0,
      purchases: camp.purchases,
      carts: 0,
      impressions: camp.impressions,
      clicks: camp.clicks,
      ctr: camp.impressions > 0 ? camp.clicks / camp.impressions : 0,
      cpc: camp.clicks > 0 ? camp.spend / camp.clicks : 0,
      cpm: camp.impressions > 0 ? (camp.spend / camp.impressions) * 1000 : 0,
      reportStart: minStart,
      reportEnd: maxEnd,
      body: camp.body,
      products,
      productCount: products.length,
    });
  }
  if (campaigns.length === 0) throw new Error('캠페인 데이터가 없습니다.');
  return campaigns;
};

const parseAdPerformance = async (buf) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('시트를 찾을 수 없습니다.');

  const headers = [];
  const headerRow = ws.getRow(1);
  const colCount = ws.columnCount || ws.actualColumnCount || 13;
  for (let c = 1; c <= colCount; c++) headers.push(adText(headerRow.getCell(c).value));

  if (headers.some(h => h && h.includes('제품 ID'))) {
    return parseAdRawByProduct(ws, headers);
  }

  const findCol = (pred) => headers.findIndex(h => h && pred(h));
  const col = {
    name: findCol(h => h.includes('광고 이름') || h.replace(/\s/g, '') === '광고이름'),
    status: findCol(h => h.includes('게재')),
    budget: findCol(h => h.includes('예산') && !h.includes('유형')),
    spend: findCol(h => h.includes('지출')),
    revenue: findCol(h => h.includes('전환값')),
    roas: findCol(h => h.toUpperCase().includes('ROAS')),
    purchases: findCol(h => h === '구매'),
    carts: findCol(h => h.includes('장바구니')),
    cpc: findCol(h => h.toUpperCase().includes('CPC')),
    cpm: findCol(h => h.toUpperCase().includes('CPM')),
    start: findCol(h => h.includes('보고 시작')),
    end: findCol(h => h.includes('보고 종료')),
  };
  if (col.name < 0) {
    throw new Error('광고 이름 컬럼을 찾을 수 없습니다. 메타 광고 관리자에서 받은 파일이 맞는지 확인해주세요.');
  }

  const campaigns = [];
  const lastRow = ws.rowCount || ws.actualRowCount || 1;
  for (let r = 2; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const cell = (i) => i >= 0 ? row.getCell(i + 1).value : null;
    const name = adText(cell(col.name));
    if (!name) continue;

    const spend = adNum(cell(col.spend));
    const revenue = adNum(cell(col.revenue));
    const cpc = adNum(cell(col.cpc));
    const cpm = adNum(cell(col.cpm));
    const purchases = adNum(cell(col.purchases));
    const carts = adNum(cell(col.carts));
    const roasRaw = adNum(cell(col.roas));
    const roas = roasRaw || (spend > 0 ? revenue / spend : 0);
    const impressions = cpm > 0 ? (spend / cpm) * 1000 : 0;
    const clicks = cpc > 0 ? spend / cpc : 0;
    const dateM = name.match(/\((\d{3,4})\)\s*$/);
    const postCode = dateM ? dateM[1].padStart(4, '0') : '';
    const reportStart = adText(cell(col.start));
    const reportEnd = adText(cell(col.end));
    const postDate = parsePostDate(postCode, reportEnd);
    const endD = reportEnd ? new Date(reportEnd) : new Date();
    const ageDays = postDate && !isNaN(endD.getTime())
      ? Math.max(0, Math.round((endD.getTime() - postDate.getTime()) / 86400000))
      : null;

    campaigns.push({
      name,
      manager: (name.split('_')[0] || '-').trim(),
      postCode,
      ageDays,
      ageGroup: ageGroupOf(ageDays),
      status: adText(cell(col.status)),
      budget: adNum(cell(col.budget)),
      spend, revenue, roas, purchases, carts, cpc, cpm,
      impressions, clicks,
      ctr: impressions > 0 ? clicks / impressions : 0,
      buyRate: carts > 0 ? purchases / carts : 0,
      reportStart, reportEnd,
    });
  }
  if (campaigns.length === 0) throw new Error('캠페인 데이터가 없습니다.');
  return campaigns;
};

const groupByProduct = (skus) => {
  const map = new Map();
  for (const s of skus) {
    if (!map.has(s.productName)) {
      map.set(s.productName, {
        productName: s.productName,
        category: s.category,
        season: s.season,
        memo1: s.memo1,
        brand: extractBrand(s.productName),
        isPackage: isPackage(s.productName),
        registDate: s.registDate,
        price: s.price,
        imageUrl: null,
        skus: [],
        totalSales: 0,
        totalRevenue: 0,
        totalCanceled: 0,
        totalCurrentStock: 0,
        earlySales: 0,
        lateSales: 0,
      });
    }
    const g = map.get(s.productName);
    if (!g.imageUrl && s.imageUrl) g.imageUrl = s.imageUrl;
    if (!g.season && s.season) g.season = s.season;
    if (!g.price && s.price) g.price = s.price;
    g.skus.push(s);
    g.totalSales += s.salesTotal;
    g.totalRevenue += s.amount;
    g.totalCanceled += s.canceled;
    g.totalCurrentStock += s.currentStock;
    const len = s.sales.length;
    const w = Math.max(1, Math.round(len * 0.1));
    for (let i = 0; i < len; i++) {
      if (i < w) g.earlySales += s.sales[i];
      if (i >= len - w) g.lateSales += s.sales[i];
    }
  }
  for (const g of map.values()) {
    const sortedSkus = [...g.skus].sort((a, b) => b.salesTotal - a.salesTotal);
    g.bestSku = sortedSkus[0] || null;
    g.bestSku2 = sortedSkus[1] && sortedSkus[1].salesTotal > 0 ? sortedSkus[1] : null;
    g.windowDays = Math.max(1, Math.round((g.skus[0]?.sales.length || 1) * 0.1));
    g.growthRate = g.earlySales > 0 ? g.lateSales / g.earlySales : (g.lateSales > 0 ? 99 : 0);
    g.cancelRate = g.totalSales > 0 ? g.totalCanceled / g.totalSales : 0;
    const numDays = g.skus[0]?.sales.length || 1;
    g.avgDaily = g.totalSales / numDays;
    let salesDays = 0;
    for (let i = 0; i < numDays; i++) {
      let s = 0;
      for (const sku of g.skus) s += sku.sales[i] || 0;
      if (s > 0) salesDays++;
    }
    g.salesDays = salesDays;
  }
  return enrichPackageSeasons(Array.from(map.values()));
};

const monthsBetween = (dateStr, now = new Date()) => {
  if (!dateStr) return Infinity;
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return Infinity;
  return (now.getFullYear() - d.getFullYear()) * 12 + (now.getMonth() - d.getMonth());
};

const computeScore = (theme, opts, group) => {
  switch (theme) {
    case 'bestseller':
    case 'category':
    case 'brand':
    case 'package':
      return group.totalSales;
    case 'rising':
      if (group.totalSales < (opts.minSales ?? 2)) return -1;
      return group.growthRate * Math.log(1 + group.totalSales);
    case 'declining': {
      const early = group.earlySales;
      const late = group.lateSales;
      if (early < (opts.minEarly ?? 3)) return -1;
      if (late >= early) return -1;
      const dropRate = (early - late) / early;
      return dropRate * Math.log(1 + early);
    }
    case 'overstock': {
      const stock = group.totalCurrentStock;
      const sales = group.totalSales;
      if (stock < (opts.minStock ?? 30)) return -1;
      if (sales > (opts.maxSales ?? 10)) return -1;
      return stock / (sales + 1);
    }
    case 'newProduct': {
      const months = monthsBetween(group.registDate);
      const cutoff = opts.newMonths ?? 6;
      if (months > cutoff) return -1;
      return group.totalSales * (1 + (cutoff - months) / cutoff);
    }
    case 'steady': {
      const months = monthsBetween(group.registDate);
      if (months < (opts.minMonths ?? 12)) return -1;
      const avgDaily = group.avgDaily ?? 0;
      if (avgDaily < (opts.minAvgDaily ?? 1)) return -1;
      return avgDaily * Math.log(1 + months / 12);
    }
    case 'frequency': {
      if (group.totalSales < (opts.minSales ?? 2)) return -1;
      if ((group.salesDays ?? 0) < (opts.minSalesDays ?? 3)) return -1;
      return group.salesDays * Math.log(1 + group.totalSales);
    }
    default:
      return group.totalSales;
  }
};

const filterByTheme = (theme, opts, groups) => {
  let list = groups.filter(g => g.totalSales > 0);
  if (theme === 'category' && opts.categories?.length > 0) {
    list = list.filter(g => opts.categories.includes(g.category));
  }
  if (theme === 'brand' && opts.brand) {
    list = list.filter(g => g.brand === opts.brand);
  }
  if (theme === 'package') {
    list = list.filter(g => g.category === '패키지');
  }
  if (opts.seasonFilters?.length > 0) {
    list = list.filter(g => opts.seasonFilters.includes(g.season));
  }
  const skipStock = theme === 'package' || opts._mode === 'csv';
  if (opts.minCurrentStock > 0 && !skipStock) {
    list = list.filter(g => g.totalCurrentStock >= opts.minCurrentStock);
  }
  if (opts.searchQuery?.trim()) {
    const q = opts.searchQuery.trim().toLowerCase();
    list = list.filter(g => g.productName.toLowerCase().includes(q));
  }
  return list;
};

const applyDiversity = (sorted, maxPerCategory, limit) => {
  const counts = new Map();
  const result = [];
  for (const g of sorted) {
    const c = counts.get(g.category) || 0;
    if (c < maxPerCategory) {
      result.push(g);
      counts.set(g.category, c + 1);
    }
    if (result.length >= limit) break;
  }
  if (result.length < limit) {
    for (const g of sorted) {
      if (!result.includes(g)) {
        result.push(g);
        if (result.length >= limit) break;
      }
    }
  }
  return result;
};

const RECOMMEND_PLAN = [
  { theme: 'bestseller', n: 4 },
  { theme: 'newProduct', n: 2 },
  { theme: 'rising', n: 2 },
];

const pickRecommendation = (groups, opts) => {
  let seasonFiltered = opts.seasonFilters?.length > 0
    ? groups.filter(g => opts.seasonFilters.includes(g.season))
    : groups;
  if (opts.minCurrentStock > 0 && opts._mode !== 'csv') {
    seasonFiltered = seasonFiltered.filter(g => g.totalCurrentStock >= opts.minCurrentStock);
  }
  if (opts.searchQuery?.trim()) {
    const q = opts.searchQuery.trim().toLowerCase();
    seasonFiltered = seasonFiltered.filter(g => g.productName.toLowerCase().includes(q));
  }

  const picks = [];
  const used = new Set();

  for (const { theme: t, n } of RECOMMEND_PLAN) {
    const scored = seasonFiltered
      .map(g => ({ ...g, score: computeScore(t, opts, g), pickReason: t }))
      .filter(g => g.score > 0 && !used.has(g.productName));
    scored.sort((a, b) => b.score - a.score);
    let added = 0;
    for (const g of scored) {
      if (added >= n) break;
      picks.push(g);
      used.add(g.productName);
      added++;
    }
  }

  if (picks.length < 8) {
    const more = seasonFiltered
      .map(g => ({ ...g, score: g.totalSales, pickReason: 'bestseller' }))
      .filter(g => g.score > 0 && !used.has(g.productName));
    more.sort((a, b) => b.score - a.score);
    for (const g of more) {
      if (picks.length >= 8) break;
      picks.push(g);
      used.add(g.productName);
    }
  }

  if (opts.useDiversity) {
    const counts = new Map();
    const result = [];
    for (const p of picks) {
      const c = counts.get(p.category) || 0;
      if (c < (opts.maxPerCategory ?? 3)) {
        result.push(p);
        counts.set(p.category, c + 1);
      }
      if (result.length >= 8) break;
    }
    for (const p of picks) {
      if (result.length >= 8) break;
      if (!result.includes(p)) result.push(p);
    }
    return result;
  }

  return picks.slice(0, 8);
};

const buildCustomPrompt = (userQuery, mode) => `당신은 광고 후보 상품 선정을 돕는 데이터 어시스턴트입니다.

상품(group) 데이터의 필드:
- productName (string): 상품명
- category (string): 카테고리 (예: 반바지, 티셔츠, 패키지 등)
- season (string): 시즌 ("S/S" | "F/W" | "사계절" | "")
- brand (string): 브랜드 코드 (FP/JM/WV/PS/EZ/TWN/PL/DY)
- registDate (string): 등록일자, "YYYY-MM-DD"
- isPackage (boolean): 패키지 상품 여부
- totalSales (number): ${mode === 'csv' ? '누적' : '기간'} 판매수량
- totalRevenue (number): 매출액(원)
- totalCurrentStock (number): 현재 재고
- avgDaily (number): 일평균 판매량
- earlySales (number, xls만): 데이터 기간 초반 10% 구간 판매
- lateSales (number, xls만): 데이터 기간 후반 10% 구간 판매

사용자 자연어 조건:
"""
${userQuery}
"""

위 조건을 만족하는 상품을 뽑아내기 위한 JSON 명세를 출력하세요.
출력 스키마(반드시 JSON만, 코드블록·주석 없이):
{
  "filters": [ { "field": "필드명", "op": "==/!=/>/>=/</<=/contains/in/notIn", "value": <값 또는 배열> }, ... ],
  "sortBy": "필드명 (생략 가능)",
  "order": "asc" | "desc",
  "limit": <숫자, 기본 8>,
  "summary": "한 줄로 어떤 기준인지 요약 (한국어)"
}

규칙:
- filters는 모두 AND 결합
- 자연어에 "이상"/"≥"는 ">=", "초과"는 ">", "이하"는 "<=", "미만"은 "<"
- 카테고리/브랜드/시즌처럼 여러 값 매칭 시 op는 "in"
- 숫자 비교는 op ">=" "<=" 등 사용
- "패키지 제외" → {"field":"isPackage","op":"==","value":false}
- 사용자가 명시 안 했으면 sortBy="totalSales", order="desc"
- limit 기본 8`;

const callGemini = async (apiKey, prompt) => {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0.2 },
    }),
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Gemini ${res.status}: ${t.slice(0, 200)}`);
  }
  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('Gemini 응답이 비어있습니다.');
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error(`Gemini 응답 파싱 실패: ${text.slice(0, 200)}`);
  }
};

const matchFilter = (group, f) => {
  const v = group[f.field];
  switch (f.op) {
    case '==': return v == f.value;
    case '!=': return v != f.value;
    case '>': return Number(v) > Number(f.value);
    case '>=': return Number(v) >= Number(f.value);
    case '<': return Number(v) < Number(f.value);
    case '<=': return Number(v) <= Number(f.value);
    case 'contains': return String(v ?? '').toLowerCase().includes(String(f.value).toLowerCase());
    case 'in': return Array.isArray(f.value) && f.value.includes(v);
    case 'notIn': return Array.isArray(f.value) && !f.value.includes(v);
    default: return true;
  }
};

const applyCustomSpec = (groups, spec) => {
  let list = groups;
  if (Array.isArray(spec.filters)) {
    list = list.filter(g => spec.filters.every(f => matchFilter(g, f)));
  }
  if (spec.sortBy) {
    const order = spec.order === 'asc' ? 1 : -1;
    list = [...list].sort((a, b) => {
      const av = a[spec.sortBy], bv = b[spec.sortBy];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (av < bv) return -1 * order;
      if (av > bv) return 1 * order;
      return 0;
    });
  }
  const limit = Number(spec.limit) || 8;
  return list.slice(0, limit);
};

const pickItems = (groups, theme, opts) => {
  if (theme === 'recommend') return pickRecommendation(groups, opts);
  if (theme === 'custom') return opts._customResults || [];

  const filtered = filterByTheme(theme, opts, groups);
  const scored = filtered
    .map(g => ({ ...g, score: computeScore(theme, opts, g) }))
    .filter(g => g.score > 0);
  scored.sort((a, b) => b.score - a.score);

  const limit = 8;
  const isSingleCategory = (theme === 'category' && opts.categories?.length === 1) || theme === 'package';
  const useDiversity = opts.useDiversity && !isSingleCategory;

  if (useDiversity) {
    return applyDiversity(scored, opts.maxPerCategory ?? 3, limit);
  }
  return scored.slice(0, limit);
};

const fetchImageBuffer = async (url) => {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 3000);
    const res = await fetch(url, { mode: 'cors', signal: controller.signal });
    clearTimeout(timer);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await blob.arrayBuffer();
  } catch (e) {
    return null;
  }
};

const detectImageExt = (url) => {
  const m = String(url).toLowerCase().match(/\.(png|jpg|jpeg|gif)(?:\?|$)/);
  if (!m) return 'png';
  return m[1] === 'jpg' ? 'jpeg' : m[1];
};

const reasonLabel = (id) => THEMES.find(t => t.id === id)?.label || id || '-';

const buildBriefing = (theme, opts, mode, picks) => {
  const n = picks.length;
  const seasonText = opts.seasonFilters?.length > 0
    ? ` 시즌은 ${opts.seasonFilters.join('·')}만 대상으로 잡았습니다.`
    : '';
  const periodText = mode === 'csv' ? '전체 누적' : '데이터 기간';

  switch (theme) {
    case 'bestseller':
      return `${periodText} 동안 가장 많이 팔린 상위 ${n}개 상품입니다.${seasonText} 검증된 수요가 있어 광고 전환율이 가장 안정적이고, ROI를 최우선으로 생각하는 메인 캠페인에 투입하기 좋습니다. 새 메시지보다는 상품 자체의 강점을 부각하는 소재가 효과적입니다.`;
    case 'rising':
      return `데이터 기간 후반 10% 구간 판매량이 초반 10% 구간 대비 빠르게 증가한 라이징 ${n}개입니다.${seasonText} 트렌드 모멘텀이 살아있는 시점이라 노출을 늘리면 추가 성장 여력이 큽니다. 신선한 영상·UGC 같은 모멘텀 친화적인 광고 소재와 함께 운영하세요.`;
    case 'declining':
      return `데이터 기간 후반 10% 구간에 판매가 초반 대비 떨어진 ${n}개 상품입니다.${seasonText} 원래 수요는 검증됐지만 모멘텀이 식은 상태로, 광고 노출과 할인·리뉴얼 메시지로 다시 살릴 만한 후보입니다. 광고를 안 돌리면 그대로 사라질 가능성이 높습니다.`;
    case 'newProduct':
      return `최근 ${opts.newMonths}개월 이내 등록되어 초기 판매가 좋은 신상품 ${n}개입니다.${seasonText} 시장 안착 단계라 인지도 확보가 우선이며, 브랜드/콘셉트를 또렷이 전달하는 소재가 효과적입니다. 초기 광고 투입 효율이 평균 대비 높습니다.`;
    case 'category': {
      const cats = (opts.categories || []).join('·') || '선택 카테고리';
      return `${cats} 안에서 ${periodText} 판매량 상위 ${n}개입니다.${seasonText} 같은 카테고리라 콘셉트가 일관되어 광고 묶음(카탈로그·릴) 운용에 적합합니다. 한 카테고리에 집중하는 캠페인 셋업으로 효율이 높아집니다.`;
    }
    case 'brand':
      return `${opts.brand} 라인 중 ${periodText} 판매량 상위 ${n}개입니다.${seasonText} 브랜드 톤이 통일되어 시리즈·컬렉션 광고나 브랜드 인지도 캠페인에 잘 어울립니다. 한 라인을 단일 광고 슬롯으로 묶어 운영하기 좋아요.`;
    case 'package':
      return `상품명에 PACK이 포함된 패키지 상품 중 판매 상위 ${n}개입니다.${seasonText} 객단가가 높아 매출 효율이 좋고, "묶음 할인"·"세트 구성" 같은 가성비 메시지와 잘 맞습니다. 신규 고객 객단가 끌어올리는 용도로도 효과적입니다.`;
    case 'steady':
      return `등록 ${opts.minMonths}개월 이상이면서 일평균 ${opts.minAvgDaily}개 이상 꾸준히 팔리는 스테디셀러 ${n}개입니다.${seasonText} 검증된 효자 상품이라 광고 안정성이 가장 높고, 시즌 비수기의 매출 백본으로 적합합니다. 대규모 예산을 안전하게 태울 수 있습니다.`;
    case 'overstock':
      return `재고 ${opts.minStock}개 이상이면서 기간 판매가 ${opts.maxSales}개 이하인 적체 상품 ${n}개입니다.${seasonText} 광고로 재고 소진을 노리거나 할인 프로모션 대상으로 활용하기 좋습니다. 시즌 마감 전에 처리해야 손실이 줄어듭니다.`;
    case 'recommend': {
      const counts = picks.reduce((m, p) => ({ ...m, [p.pickReason]: (m[p.pickReason] || 0) + 1 }), {});
      const parts = Object.entries(counts).map(([k, v]) => `${reasonLabel(k)} ${v}`).join(' · ');
      return `${parts}로 구성된 균형 포트폴리오 ${n}개입니다.${seasonText} 단일 의도가 아니라 메인 매출(베스트셀러)·트렌드(급상승)·신선도(신상품)·재활성(판매 감소)·재고 소진(재고 과다)을 동시에 챙기는 종합 캠페인용 픽입니다. 광고 슬롯을 분산해 운영하기에 가장 효율적입니다.`;
    }
    case 'custom':
      return `사용자 정의 조건으로 추출한 ${n}개 상품입니다. 적용된 기준은 우측 사이드바의 "추출 완료" 박스에서 확인할 수 있습니다.`;
    default:
      return `${THEMES.find(t => t.id === theme)?.label || ''} 주제로 선정된 ${n}개 상품입니다.${seasonText}`;
  }
};

const sanitizeSheetName = (name) => {
  let n = String(name).replace(/[\\\/\*\?\[\]:]/g, '_').trim();
  if (n.length > 31) n = n.slice(0, 31);
  return n || 'Sheet';
};

const buildItemsSheet = async (wb, sheetName, items, theme, embedImages, onImage) => {
  const ws = wb.addWorksheet(sanitizeSheetName(sheetName));
  const isRecommend = theme === 'recommend';
  ws.columns = [
    { header: '순위', key: 'rank', width: 6 },
    { header: '이미지', key: 'image', width: 16 },
    { header: '상품명', key: 'productName', width: 42 },
    { header: '카테고리', key: 'category', width: 12 },
    { header: '시즌', key: 'season', width: 8 },
    { header: '총 판매량', key: 'totalSales', width: 12 },
    { header: '베스트 SKU 1위', key: 'bestSku1', width: 26 },
    { header: '1위 판매량', key: 'bestSku1Sales', width: 12 },
    { header: '베스트 SKU 2위', key: 'bestSku2', width: 26 },
    { header: '2위 판매량', key: 'bestSku2Sales', width: 12 },
    { header: '현재 재고', key: 'currentStock', width: 12 },
    { header: '매출액', key: 'revenue', width: 14 },
    ...(isRecommend ? [{ header: '선정 사유', key: 'score', width: 18 }] : []),
  ];
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
  ws.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 28;

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const rowData = {
      rank: i + 1,
      image: '',
      productName: it.productName,
      category: it.category,
      season: it.season || '-',
      totalSales: it.totalSales,
      bestSku1: it.bestSku?.optionName || '-',
      bestSku1Sales: it.bestSku?.salesTotal || 0,
      bestSku2: it.bestSku2?.optionName || '-',
      bestSku2Sales: it.bestSku2?.salesTotal || 0,
      currentStock: it.totalCurrentStock,
      revenue: it.totalRevenue,
    };
    if (isRecommend) rowData.score = reasonLabel(it.pickReason);
    const row = ws.addRow(rowData);
    row.height = 80;
    row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    row.getCell('productName').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    row.getCell('revenue').numFmt = '#,##0';
    row.getCell('currentStock').numFmt = '#,##0';
  }

  if (embedImages) {
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      onImage?.();
      if (!it.imageUrl || it.imageUrl === '이미지없음') {
        ws.getRow(i + 2).getCell('image').value = '이미지 없음';
        continue;
      }
      const buf = await fetchImageBuffer(it.imageUrl);
      if (!buf) {
        ws.getRow(i + 2).getCell('image').value = { text: '이미지 보기', hyperlink: it.imageUrl };
        ws.getRow(i + 2).getCell('image').font = { color: { argb: 'FF2563EB' }, underline: true };
        continue;
      }
      const imgId = wb.addImage({ buffer: buf, extension: detectImageExt(it.imageUrl) });
      ws.addImage(imgId, { tl: { col: 1.1, row: i + 1.1 }, ext: { width: 96, height: 96 } });
    }
  } else {
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (it.imageUrl && it.imageUrl !== '이미지없음') {
        ws.getRow(i + 2).getCell('image').value = { text: '이미지 보기', hyperlink: it.imageUrl };
        ws.getRow(i + 2).getCell('image').font = { color: { argb: 'FF2563EB' }, underline: true };
      } else {
        ws.getRow(i + 2).getCell('image').value = '없음';
      }
    }
  }
  return ws;
};

const buildMetaSheet = (wb, info) => {
  const meta = wb.addWorksheet('실행정보');
  meta.columns = [{ key: 'k', width: 22 }, { key: 'v', width: 90 }];
  for (const [k, v] of info) {
    const row = meta.addRow({ k, v });
    if (typeof v === 'string' && v.length > 60) {
      row.alignment = { wrapText: true, vertical: 'top' };
      row.height = Math.min(160, 20 + Math.ceil(v.length / 60) * 16);
      row.getCell('k').alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
    }
  }
  return meta;
};

const saveWorkbook = async (wb, baseName) => {
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 16);
  a.href = url;
  a.download = `${baseName}_${ts}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
};

const exportSingleTheme = async (items, theme, opts, embedImages, onProgress) => {
  const wb = new ExcelJS.Workbook();
  const themeLabel = THEMES.find(t => t.id === theme)?.label ?? theme;
  let processed = 0;
  await buildItemsSheet(wb, '광고 후보', items, theme, embedImages, () => {
    processed++;
    onProgress?.({ phase: 'image', cur: processed, total: items.length });
  });

  const info = [
    ['주제', themeLabel],
    ['실행시각', new Date().toLocaleString('ko-KR')],
  ];
  if (opts.categories?.length > 0) info.push(['카테고리', opts.categories.join(', ')]);
  if (opts.brand) info.push(['브랜드', opts.brand]);
  if (theme === 'newProduct') info.push(['신상품 기준(개월)', opts.newMonths]);
  if (theme === 'steady') {
    info.push(['스테디셀러 최소 개월', opts.minMonths]);
    info.push(['스테디셀러 일평균 임계', `${opts.minAvgDaily} 개/일`]);
  }
  if (opts.seasonFilters?.length > 0) info.push(['시즌 필터', opts.seasonFilters.join(', ')]);
  if (opts.minCurrentStock > 0) info.push(['재고 임계', `${opts.minCurrentStock}개 미만 제외`]);
  if (opts.useDiversity) info.push(['카테고리 다양성', `한 카테고리 최대 ${opts.maxPerCategory}개`]);
  info.push(['선정 사유 브리핑', buildBriefing(theme, opts, opts._mode, items)]);
  buildMetaSheet(wb, info);

  await saveWorkbook(wb, `ADpicker_${themeLabel}`);
};

const exportAllThemes = async (groups, mode, opts, embedImages, brands, categories, selectedThemeIds, onProgress) => {
  const visibleThemes = THEMES.filter(t =>
    t.modes.includes(mode) && (!selectedThemeIds || selectedThemeIds.includes(t.id))
  );
  const sheetPlans = [];

  for (const t of visibleThemes) {
    if (t.id === 'category') {
      for (const cat of categories) {
        const items = pickItems(groups, 'category', { ...opts, categories: [cat] });
        if (items.length > 0) sheetPlans.push({ name: `카테고리·${cat}`, items, theme: 'category' });
      }
    } else if (t.id === 'brand') {
      for (const b of brands) {
        const items = pickItems(groups, 'brand', { ...opts, brand: b });
        if (items.length > 0) sheetPlans.push({ name: `브랜드·${b}`, items, theme: 'brand' });
      }
    } else {
      const items = pickItems(groups, t.id, opts);
      if (items.length > 0) sheetPlans.push({ name: t.label, items, theme: t.id });
    }
  }

  if (sheetPlans.length === 0) throw new Error('생성할 시트가 없습니다.');

  const wb = new ExcelJS.Workbook();
  const totalImages = embedImages ? sheetPlans.reduce((s, p) => s + p.items.length, 0) : 0;
  let processedImages = 0;

  for (let i = 0; i < sheetPlans.length; i++) {
    const plan = sheetPlans[i];
    onProgress?.({ phase: 'sheet', cur: i + 1, total: sheetPlans.length, sheetName: plan.name });
    await buildItemsSheet(wb, plan.name, plan.items, plan.theme, embedImages, () => {
      processedImages++;
      if (totalImages > 0) {
        onProgress?.({ phase: 'image', cur: processedImages, total: totalImages, sheetName: plan.name });
      }
    });
  }

  const info = [
    ['모드', mode.toUpperCase()],
    ['실행시각', new Date().toLocaleString('ko-KR')],
    ['시트 수', sheetPlans.length],
  ];
  if (opts.seasonFilters?.length > 0) info.push(['시즌 필터', opts.seasonFilters.join(', ')]);
  if (opts.minCurrentStock > 0) info.push(['재고 임계', `${opts.minCurrentStock}개 미만 제외`]);
  if (opts.useDiversity) info.push(['카테고리 다양성', `한 카테고리 최대 ${opts.maxPerCategory}개`]);
  for (const plan of sheetPlans) {
    info.push([`📋 ${plan.name}`, buildBriefing(plan.theme, opts, mode, plan.items)]);
  }
  buildMetaSheet(wb, info);

  await saveWorkbook(wb, `ADpicker_전체_${mode.toUpperCase()}`);
};

const App = () => {
  const [skus, setSkus] = useState([]);
  const [groups, setGroups] = useState([]);
  const [campaigns, setCampaigns] = useState([]);
  const [campaignsName, setCampaignsName] = useState('');
  const [adList, setAdList] = useState(null);
  const [adListName, setAdListName] = useState('');
  const [dateLabels, setDateLabels] = useState([]);
  const [fileName, setFileName] = useState('');
  const [mode, setMode] = useState(null);
  const [error, setError] = useState(null);
  const [parsing, setParsing] = useState(false);

  const [theme, setTheme] = useState('bestseller');
  const [opts, setOpts] = useState({
    categories: [],
    brand: '',
    newMonths: 6,
    minMonths: 24,
    minAvgDaily: 1,
    minSales: 2,
    minEarly: 3,
    minStock: 30,
    maxSales: 10,
    minCurrentStock: 5,
    minSalesDays: 3,
    searchQuery: '',
    seasonFilters: [],
    useDiversity: true,
    maxPerCategory: 3,
  });
  const [embedImages, setEmbedImages] = useState(true);
  const [exporting, setExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState(null);
  const [showSelector, setShowSelector] = useState(false);
  const [excluded, setExcluded] = useState([]);
  const [customQuery, setCustomQuery] = useState('');
  const [apiKey, setApiKey] = useState(() =>
    typeof window !== 'undefined' ? (localStorage.getItem('geminiApiKey') || '') : ''
  );
  const [customResults, setCustomResults] = useState([]);
  const [customSpec, setCustomSpec] = useState(null);
  const [customLoading, setCustomLoading] = useState(false);
  const [customError, setCustomError] = useState(null);
  const fileInputRef = useRef(null);

  const categories = useMemo(() => {
    const EXCLUDED = new Set(['기타', '패키지', '미분류']);
    const set = new Set();
    for (const g of groups) if (g.category && !EXCLUDED.has(g.category)) set.add(g.category);
    return Array.from(set).sort();
  }, [groups]);

  const brands = useMemo(() => {
    const set = new Set();
    for (const g of groups) if (g.brand) set.add(g.brand);
    return BRAND_CODES.filter(b => set.has(b));
  }, [groups]);

  const visibleThemes = useMemo(
    () => mode ? THEMES.filter(t => t.modes.includes(mode)) : THEMES,
    [mode]
  );

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setParsing(true);
    setError(null);
    try {
      const buf = await file.arrayBuffer();
      const bytes = new Uint8Array(buf);
      const isXlsx = bytes[0] === 0x50 && bytes[1] === 0x4B;
      if (isXlsx) {
        const zip = await JSZip.loadAsync(buf);
        if (await isAdListXlsx(zip)) {
          const list = await parseAdListXlsx(zip);
          setAdList(list);
          setAdListName(file.name);
          if (skus.length > 0) setMode('adtrack');
          return;
        }
        const parsedCampaigns = await parseAdPerformance(buf);
        setCampaigns(parsedCampaigns);
        setCampaignsName(file.name);
        // 단독으로 광고성과만 올린 경우만 adperf 모드 진입
        if (!adList && skus.length === 0) {
          setSkus([]);
          setGroups([]);
          setDateLabels([]);
          setMode('adperf');
          setFileName(file.name);
        }
        return;
      }
      const decoder = new TextDecoder('euc-kr');
      let text = decoder.decode(buf);
      if (!text.includes('상품명') && !text.includes('<table')) {
        text = new TextDecoder('utf-8').decode(buf);
      }

      if (isAdListCsv(text)) {
        const list = parseAdList(text);
        setAdList(list);
        setAdListName(file.name);
        if (skus.length > 0) setMode('adtrack');
        return;
      }

      const isHtml = text.trimStart().startsWith('<');
      if (isHtml) {
        const { skus: parsedSkus, dateLabels: dl } = parseHtmlXls(text);
        setSkus(parsedSkus);
        setGroups(groupByProduct(parsedSkus));
        setDateLabels(dl);
        setFileName(file.name);
        if (adList) {
          setMode('adtrack');
        } else {
          if (mode !== 'xls') setTheme('bestseller');
          setMode('xls');
        }
      } else {
        const productGroups = parseProductCsv(text);
        setSkus([]);
        setGroups(productGroups);
        setDateLabels([]);
        setFileName(file.name);
        if (mode !== 'csv') setTheme('bestseller');
        setMode('csv');
      }
    } catch (e) {
      setError(`파일 파싱 실패: ${e.message}`);
    } finally {
      setParsing(false);
    }
  }, [adList, skus, mode]);

  const preview = useMemo(() => {
    if (groups.length === 0) return [];
    const exSet = new Set(excluded);
    const filteredGroups = exSet.size > 0 ? groups.filter(g => !exSet.has(g.productName)) : groups;
    const filteredCustom = exSet.size > 0 ? customResults.filter(g => !exSet.has(g.productName)) : customResults;
    return pickItems(filteredGroups, theme, { ...opts, _customResults: filteredCustom, _mode: mode });
  }, [groups, theme, opts, customResults, mode, excluded]);

  const handleExclude = useCallback((name) => {
    setExcluded(prev => prev.includes(name) ? prev : [...prev, name]);
  }, []);

  const handleResetExcluded = useCallback(() => {
    setExcluded([]);
  }, []);

  const saveApiKey = useCallback((k) => {
    setApiKey(k);
    if (typeof window !== 'undefined') {
      if (k) localStorage.setItem('geminiApiKey', k);
      else localStorage.removeItem('geminiApiKey');
    }
  }, []);

  const handleRunCustom = useCallback(async () => {
    if (!customQuery.trim()) {
      setCustomError('조건을 입력해 주세요.');
      return;
    }
    if (!apiKey.trim()) {
      setCustomError('Gemini API 키를 먼저 입력해 주세요.');
      return;
    }
    setCustomLoading(true);
    setCustomError(null);
    setCustomSpec(null);
    try {
      const spec = await callGemini(apiKey, buildCustomPrompt(customQuery, mode));
      const items = applyCustomSpec(groups, spec);
      setCustomSpec(spec);
      setCustomResults(items);
      if (items.length === 0) {
        setCustomError('조건에 맞는 상품이 없어요. 조건을 완화해 보세요.');
      }
    } catch (e) {
      setCustomError(e.message);
      setCustomResults([]);
    } finally {
      setCustomLoading(false);
    }
  }, [apiKey, customQuery, groups, mode]);

  const handleExportSingle = useCallback(async () => {
    if (preview.length === 0) return;
    setExporting('single');
    setExportProgress(null);
    try {
      await exportSingleTheme(preview, theme, { ...opts, _mode: mode }, embedImages, (p) => setExportProgress(p));
    } catch (e) {
      setError(`엑셀 저장 실패: ${e.message}`);
    } finally {
      setExporting(false);
      setExportProgress(null);
    }
  }, [preview, theme, opts, embedImages]);

  const handleExportSelected = useCallback(async (selectedIds) => {
    if (groups.length === 0 || !mode) return;
    setShowSelector(false);
    setExporting('all');
    setExportProgress(null);
    try {
      await exportAllThemes(groups, mode, { ...opts, _mode: mode }, embedImages, brands, categories, selectedIds, (p) => setExportProgress(p));
    } catch (e) {
      setError(`다운 실패: ${e.message}`);
    } finally {
      setExporting(false);
      setExportProgress(null);
    }
  }, [groups, mode, opts, embedImages, brands, categories]);

  const themeIcon = THEMES.find(t => t.id === theme)?.icon || LucideStar;
  const ThemeIcon = themeIcon;

  return (
    <div className="min-h-screen bg-cream-300 text-stone-900">
      <div className="w-full px-8 py-6 bg-cream-100 min-h-screen border-x border-cream-400">
        <header className="mb-6 pb-5 border-b border-cream-400 flex items-end justify-between gap-4 flex-wrap">
          <div>
            <h1 className="text-4xl font-medium flex items-baseline gap-3 tracking-tight">
              ADpicker
              <span className="font-sans text-xs font-normal text-stone-500 tracking-wide uppercase">v1</span>
            </h1>
            <p className="text-sm text-stone-600 mt-2 font-light">인스타 메타광고 아이템 선정기</p>
          </div>
          {fileName && mode !== 'adperf' && (
            <div className="flex items-center gap-3 flex-wrap">
              <div className="flex items-center gap-1.5 bg-cream-50 px-3 py-1.5 border border-cream-400">
                <LucideSearch size={13} className="text-stone-500" />
                <input
                  type="text"
                  value={opts.searchQuery}
                  onChange={e => setOpts({ ...opts, searchQuery: e.target.value })}
                  placeholder="상품명 검색"
                  className="text-sm bg-transparent focus:outline-none w-36 placeholder:text-stone-400"
                />
                {opts.searchQuery && (
                  <button
                    onClick={() => setOpts({ ...opts, searchQuery: '' })}
                    className="text-stone-400 hover:text-stone-700"
                  >
                    <LucideX size={13} />
                  </button>
                )}
              </div>
              <div className="flex items-center gap-1.5 bg-cream-50 px-3 py-1.5 border border-cream-400">
                <label className="text-xs font-medium text-stone-600 mr-1">시즌</label>
                {SEASON_VALUES.map(v => {
                  const active = (opts.seasonFilters || []).includes(v);
                  return (
                    <button
                      key={v}
                      onClick={() => {
                        const cur = opts.seasonFilters || [];
                        const next = cur.includes(v) ? cur.filter(x => x !== v) : [...cur, v];
                        setOpts({ ...opts, seasonFilters: next });
                      }}
                      className={`px-2 py-0.5 text-xs border transition ${
                        active
                          ? 'bg-stone-900 text-cream-50 border-stone-900'
                          : 'bg-cream-50 text-stone-700 border-cream-400 hover:border-stone-700'
                      }`}
                    >
                      {v}
                    </button>
                  );
                })}
              </div>
              <button
                onClick={handleExportSingle}
                disabled={exporting || preview.length === 0}
                className="bg-stone-900 hover:bg-stone-800 disabled:bg-cream-300 disabled:text-stone-400 disabled:cursor-not-allowed text-cream-50 px-4 py-2 font-medium flex items-center gap-2 text-sm transition"
              >
                <LucideDownload size={14} />
                {exporting === 'single'
                  ? exportProgress?.phase === 'image'
                    ? `이미지 ${exportProgress.cur}/${exportProgress.total}...`
                    : '생성 중...'
                  : '현재 다운'}
              </button>
              <button
                onClick={() => setShowSelector(true)}
                disabled={exporting || groups.length === 0}
                className="bg-cream-50 hover:bg-cream-200 disabled:bg-cream-200 disabled:text-stone-400 disabled:cursor-not-allowed text-stone-900 border border-stone-900 px-4 py-2 font-medium flex items-center gap-2 text-sm transition"
              >
                <LucideDownload size={14} />
                {exporting === 'all'
                  ? exportProgress?.phase === 'image'
                    ? `이미지 ${exportProgress.cur}/${exportProgress.total}...`
                    : exportProgress?.phase === 'sheet'
                      ? `시트 ${exportProgress.cur}/${exportProgress.total}...`
                      : '생성 중...'
                  : '선택 다운'}
              </button>
              <div className="text-xs text-stone-600 flex items-center gap-2 bg-cream-50 px-3 py-2 border border-cream-400">
                <LucideFileSpreadsheet size={14} />
                {fileName}
                <span className="text-stone-400">·</span>
                {skus.length > 0
                  ? `${skus.length.toLocaleString()} SKU · ${groups.length.toLocaleString()} 상품`
                  : `${groups.length.toLocaleString()} 상품 (누적)`}
              </div>
            </div>
          )}
        </header>

        {error && (
          <div className="mb-4 bg-rose-50 border border-rose-300 text-rose-900 px-4 py-3 flex items-start gap-2">
            <LucideAlertCircle size={18} className="mt-0.5 shrink-0" />
            <span className="flex-1 text-sm">{error}</span>
            <button onClick={() => setError(null)}><LucideX size={16} /></button>
          </div>
        )}

        {mode === 'adtrack' ? (
          <AdTrackView
            adList={adList}
            adListName={adListName}
            groups={groups}
            dateLabels={dateLabels}
            fileName={fileName}
            campaigns={campaigns}
            campaignsName={campaignsName}
            onReset={() => {
              setAdList(null);
              setAdListName('');
              setSkus([]);
              setGroups([]);
              setDateLabels([]);
              setFileName('');
              setCampaigns([]);
              setCampaignsName('');
              setMode(null);
            }}
          />
        ) : mode === 'adperf' ? (
          <AdPerformanceView
            campaigns={campaigns}
            fileName={fileName}
            onReset={() => {
              setCampaigns([]);
              setFileName('');
              setMode(null);
            }}
          />
        ) : groups.length === 0 ? (
          <UploadArea
            onFile={handleFile}
            parsing={parsing}
            inputRef={fileInputRef}
            adList={adList}
            adListName={adListName}
            onClearAdList={() => { setAdList(null); setAdListName(''); }}
          />
        ) : (
          <div className="grid grid-cols-12 gap-6">
            <aside className="col-span-12 lg:col-span-4 xl:col-span-3 space-y-4">
              <Panel title={`주제 선택 ${mode ? `· ${mode.toUpperCase()} 모드` : ''}`} icon={ThemeIcon}>
                <div className="grid grid-cols-2 gap-0 border border-cream-400">
                  {visibleThemes.map((t, idx) => {
                    const Icon = t.icon;
                    const active = theme === t.id;
                    const col = idx % 2;
                    const row = Math.floor(idx / 2);
                    const totalRows = Math.ceil(visibleThemes.length / 2);
                    return (
                      <button
                        key={t.id}
                        onClick={() => setTheme(t.id)}
                        title={t.desc}
                        className={`flex flex-col items-center justify-center gap-1.5 px-2 py-3 text-center transition ${
                          col === 0 ? 'border-r border-cream-400' : ''
                        } ${row < totalRows - 1 ? 'border-b border-cream-400' : ''} ${
                          active
                            ? 'bg-stone-900 text-cream-50'
                            : 'bg-cream-50 text-stone-800 hover:bg-cream-100'
                        }`}
                      >
                        <Icon size={18} strokeWidth={1.5} className={active ? 'text-cream-50' : 'text-stone-700'} />
                        <div className="text-xs font-medium leading-tight">{t.label}</div>
                      </button>
                    );
                  })}
                </div>
              </Panel>

              <Panel title="주제별 옵션">
                {theme !== 'custom' && theme !== 'package' && mode !== 'csv' && (
                  <div className="mb-3 pb-3 border-b border-cream-300 flex items-center gap-2">
                    <label className="text-xs text-stone-600 whitespace-nowrap">현재 재고</label>
                    <input
                      type="number"
                      min="0"
                      value={opts.minCurrentStock}
                      onChange={e => setOpts({ ...opts, minCurrentStock: parseInt(e.target.value) || 0 })}
                      className="w-16 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
                    />
                    <span className="text-xs text-stone-500">개 미만 제외</span>
                  </div>
                )}
                <ThemeOptions
                  theme={theme}
                  opts={opts}
                  setOpts={setOpts}
                  categories={categories}
                  brands={brands}
                  customQuery={customQuery}
                  setCustomQuery={setCustomQuery}
                  apiKey={apiKey}
                  saveApiKey={saveApiKey}
                  customLoading={customLoading}
                  customError={customError}
                  customSpec={customSpec}
                  customResultsCount={customResults.length}
                  onRunCustom={handleRunCustom}
                />
              </Panel>

              <button
                onClick={() => {
                  setSkus([]);
                  setGroups([]);
                  setFileName('');
                  setMode(null);
                }}
                className="w-full text-stone-500 hover:text-stone-900 text-xs py-2 underline underline-offset-4"
              >
                다른 파일 불러오기
              </button>
            </aside>

            <main className="col-span-12 lg:col-span-8 xl:col-span-9">
              <Preview
                items={preview}
                theme={theme}
                dateLabels={dateLabels}
                mode={mode}
                opts={opts}
                excludedCount={excluded.length}
                onExclude={handleExclude}
                onResetExcluded={handleResetExcluded}
              />
            </main>
          </div>
        )}
      </div>
      {showSelector && (
        <ThemeSelectorModal
          mode={mode}
          categoriesCount={categories.length}
          brandsCount={brands.length}
          embedImages={embedImages}
          setEmbedImages={setEmbedImages}
          opts={opts}
          setOpts={setOpts}
          onCancel={() => setShowSelector(false)}
          onConfirm={handleExportSelected}
        />
      )}
    </div>
  );
};

const ThemeSelectorModal = ({
  mode, categoriesCount, brandsCount, embedImages, setEmbedImages, opts, setOpts, onCancel, onConfirm,
}) => {
  const visibleThemes = THEMES.filter(t => t.modes.includes(mode));
  const [selected, setSelected] = useState(() => visibleThemes.map(t => t.id));

  const toggle = (id) => {
    setSelected(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
  };

  const sheetEstimate = (id) => {
    if (id === 'category') return `시트 ${categoriesCount}개`;
    if (id === 'brand') return `시트 ${brandsCount}개`;
    return '시트 1개';
  };

  const totalSheets = selected.reduce((sum, id) => {
    if (id === 'category') return sum + categoriesCount;
    if (id === 'brand') return sum + brandsCount;
    return sum + 1;
  }, 0);

  return (
    <div className="fixed inset-0 z-50 bg-stone-900/40 flex items-center justify-center p-4" onClick={onCancel}>
      <div
        className="bg-cream-50 shadow-xl w-full max-w-lg max-h-[90vh] flex flex-col overflow-hidden"
        onClick={e => e.stopPropagation()}
      >
        <div className="px-5 py-4 border-b border-cream-400 flex items-center justify-between">
          <div>
            <h2 className="font-semibold">다운받을 주제 선택</h2>
            <p className="text-xs text-stone-500 mt-0.5">{mode?.toUpperCase()} 모드 · 체크된 주제만 시트로 생성</p>
          </div>
          <button onClick={onCancel} className="text-stone-400 hover:text-stone-700">
            <LucideX size={20} />
          </button>
        </div>

        <div className="px-5 py-3 border-b border-cream-300 flex items-center justify-between text-xs">
          <div className="flex gap-2">
            <button
              onClick={() => setSelected(visibleThemes.map(t => t.id))}
              className="text-stone-600 hover:text-stone-900 underline"
            >전체 선택</button>
            <span className="text-stone-300">·</span>
            <button
              onClick={() => setSelected([])}
              className="text-stone-600 hover:text-stone-900 underline"
            >해제</button>
          </div>
          <span className="text-stone-500">총 {totalSheets} 시트 생성 예정</span>
        </div>

        <div className="overflow-y-auto px-5 py-3 space-y-1.5">
          {visibleThemes.map(t => {
            const Icon = t.icon;
            const checked = selected.includes(t.id);
            return (
              <label
                key={t.id}
                className={`flex items-start gap-3 px-3 py-2.5 border cursor-pointer transition ${
                  checked ? 'bg-cream-100 border-cream-400' : 'bg-cream-50 border-cream-400 hover:border-cream-400'
                }`}
              >
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => toggle(t.id)}
                  className="mt-1"
                />
                <Icon size={16} className="mt-0.5 text-stone-500 shrink-0" />
                <div className="flex-1 min-w-0">
                  <div className="flex items-baseline justify-between gap-2">
                    <span className="text-sm font-medium">{t.label}</span>
                    <span className="text-xs text-stone-400 shrink-0">{sheetEstimate(t.id)}</span>
                  </div>
                  <p className="text-xs text-stone-500 mt-0.5">{t.desc}</p>
                </div>
              </label>
            );
          })}
        </div>

        <div className="px-5 py-3 border-t border-cream-300 space-y-2 text-sm bg-cream-100">
          <label className="flex items-center gap-2">
            <input
              type="checkbox"
              checked={opts.useDiversity}
              onChange={e => setOpts({ ...opts, useDiversity: e.target.checked })}
            />
            카테고리 다양성 (한 카테고리 최대
            <input
              type="number"
              min="1"
              max="8"
              value={opts.maxPerCategory}
              onChange={e => setOpts({ ...opts, maxPerCategory: parseInt(e.target.value) || 3 })}
              className="w-12 px-2 py-0.5 text-sm border border-cream-400 rounded mx-1"
            />
            개)
          </label>
          <label className="flex items-center gap-2">
            <input
              type="checkbox"
              checked={embedImages}
              onChange={e => setEmbedImages(e.target.checked)}
            />
            <LucideImage size={14} className="text-stone-500" />
            이미지 셀에 임베드 (느려짐, 끄면 URL 링크만)
          </label>
        </div>

        <div className="px-5 py-3 border-t border-cream-400 flex items-center justify-end gap-2 bg-cream-50">
          <button
            onClick={onCancel}
            className="px-4 py-2 text-sm text-stone-600 hover:text-stone-900"
          >취소</button>
          <button
            onClick={() => onConfirm(selected)}
            disabled={selected.length === 0}
            className="bg-stone-900 hover:bg-stone-800 disabled:bg-cream-300 text-white px-5 py-2 text-sm font-medium flex items-center gap-2"
          >
            <LucideDownload size={14} />
            다운로드
          </button>
        </div>
      </div>
    </div>
  );
};

const UploadArea = ({ onFile, parsing, inputRef, adList, adListName, onClearAdList }) => {
  const [drag, setDrag] = useState(false);
  return (
    <div className="space-y-4">
      {adList && (
        <div className="flex items-center gap-3 bg-cream-200 border border-cream-400 px-4 py-3">
          <LucideMessageSquare size={16} className="text-stone-600 shrink-0" />
          <div className="flex-1 text-sm">
            <span className="font-medium text-stone-800">광고리스트 로드됨</span>
            <span className="text-stone-500"> · {adListName} · 캠페인 {adList.length}개</span>
            <div className="text-xs text-stone-500 mt-0.5">
              이제 <span className="font-medium text-stone-700">상품 매출 .xls</span> 파일을 올리면 광고-판매 추적이 시작됩니다.
            </div>
          </div>
          <button onClick={onClearAdList} className="text-stone-400 hover:text-stone-700">
            <LucideX size={16} />
          </button>
        </div>
      )}
      <div
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDrag(false);
          const f = e.dataTransfer.files[0];
          if (f) onFile(f);
        }}
        className={`border border-dashed p-20 text-center transition ${
          drag ? 'border-stone-900 bg-cream-200' : 'border-cream-400 bg-cream-50'
        }`}
      >
        <LucideUpload size={40} strokeWidth={1.2} className="mx-auto text-stone-500 mb-6" />
        <h2 className="text-3xl font-medium mb-3 tracking-tight">데이터 파일을 올려주세요</h2>
        <p className="text-sm text-stone-600 mb-8 font-light leading-relaxed">
          <span className="font-medium text-stone-800">stk_forOptSalesInfo .xls</span> (SKU 단위 · 이미지 포함) ·{' '}
          <span className="font-medium text-stone-800">sts_prdListStatistics .csv</span> (상품 누적) ·{' '}
          <span className="font-medium text-stone-800">메타 광고 성과 .xlsx</span> ·{' '}
          <span className="font-medium text-stone-800">SNS 광고리스트 .xlsx/.csv</span> — 모두 자동 인식해요.
          <br />광고리스트 + 상품매출 .xls를 함께 올리면 광고-판매 추적 화면이 나옵니다. 메타 광고성과 .xlsx를 추가로 올리면 캠페인 ROAS·매출도 함께 표시됩니다.
          <br />파일을 끌어다 놓거나 아래 버튼으로 선택하세요. 데이터는 브라우저에서만 처리됩니다.
        </p>
        <input
          ref={inputRef}
          type="file"
          accept=".xls,.xlsx,.html,.htm,.csv"
          className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) onFile(f); }}
        />
        <button
          onClick={() => inputRef.current?.click()}
          disabled={parsing}
          className="bg-stone-900 hover:bg-stone-800 disabled:bg-stone-400 text-cream-50 px-8 py-3 font-medium tracking-wide text-sm"
        >
          {parsing ? '파싱 중...' : '파일 선택'}
        </button>
      </div>
    </div>
  );
};

const AD_STATUS_LABEL = { active: '게재중', inactive: '꺼짐', not_delivering: '미게재', paused: '일시중지' };

const TERM_DESC = {
  ROAS: 'Return On Ad Spend — 광고비 대비 매출 비율. 예: ROAS 4 = 광고비 1원당 매출 4원. 광고 효율의 핵심 지표.',
  CPC: 'Cost Per Click — 광고 클릭 1회당 평균 비용. 낮을수록 클릭 효율 좋음.',
  CPM: 'Cost Per Mille — 노출 1,000회당 비용. 도달 단가 지표.',
  CTR: 'Click-Through Rate — 노출 대비 클릭률 (클릭 / 노출). 광고 소재 매력도 지표.',
};

const InfoTerm = ({ term, children }) => (
  <span
    title={TERM_DESC[term] || ''}
    className="cursor-help decoration-dotted underline decoration-stone-400 underline-offset-2"
  >
    {children || term}
  </span>
);
const fmtWon0 = (n) => Math.round(n || 0).toLocaleString();
const fmtPct1 = (n) => `${((n || 0) * 100).toFixed(1)}%`;

const AdPerformanceView = ({ campaigns, fileName, onReset }) => {
  const isProductMode = campaigns.length > 0 && Array.isArray(campaigns[0].products);
  const [sortKey, setSortKey] = useState(isProductMode ? 'spend' : 'roas');
  const [sortDir, setSortDir] = useState('desc');
  const [targetRoas, setTargetRoas] = useState(3);
  const [statusFilter, setStatusFilter] = useState('all');
  const [ageFilter, setAgeFilter] = useState('all');
  const [expanded, setExpanded] = useState(() => new Set());

  const totals = useMemo(() => {
    const sum = (k) => campaigns.reduce((s, c) => s + (c[k] || 0), 0);
    const spend = sum('spend'), revenue = sum('revenue');
    const impressions = sum('impressions'), clicks = sum('clicks');
    const carts = sum('carts'), purchases = sum('purchases');
    return {
      spend, revenue, impressions, clicks, carts, purchases,
      roas: spend > 0 ? revenue / spend : 0,
      ctr: impressions > 0 ? clicks / impressions : 0,
      buyRate: carts > 0 ? purchases / carts : 0,
    };
  }, [campaigns]);

  const totalProducts = useMemo(
    () => campaigns.reduce((s, c) => s + (c.productCount || 0), 0),
    [campaigns]
  );

  const ageStats = useMemo(() => {
    const ids = [...AGE_GROUPS.map(g => g.id), '미상'];
    return ids.map(id => {
      const list = campaigns.filter(c => c.ageGroup === id);
      const spend = list.reduce((s, c) => s + c.spend, 0);
      const revenue = list.reduce((s, c) => s + c.revenue, 0);
      const impressions = list.reduce((s, c) => s + c.impressions, 0);
      const clicks = list.reduce((s, c) => s + c.clicks, 0);
      const meta = AGE_GROUPS.find(g => g.id === id);
      return {
        id, range: meta ? meta.range : '게시일 미상',
        count: list.length, spend, revenue, clicks, impressions,
        roas: spend > 0 ? revenue / spend : 0,
        ctr: impressions > 0 ? clicks / impressions : 0,
      };
    }).filter(s => s.count > 0);
  }, [campaigns]);

  const statuses = useMemo(
    () => ['all', ...Array.from(new Set(campaigns.map(c => c.status).filter(Boolean)))],
    [campaigns]
  );

  const rows = useMemo(() => {
    let list = campaigns;
    if (statusFilter !== 'all') list = list.filter(c => c.status === statusFilter);
    if (ageFilter !== 'all') list = list.filter(c => c.ageGroup === ageFilter);
    return [...list].sort((a, b) => {
      const av = a[sortKey], bv = b[sortKey];
      let cmp;
      if (typeof av === 'string' || typeof bv === 'string') cmp = String(av).localeCompare(String(bv));
      else cmp = (av || 0) - (bv || 0);
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }, [campaigns, sortKey, sortDir, statusFilter, ageFilter]);

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir(d => (d === 'asc' ? 'desc' : 'asc'));
    else { setSortKey(key); setSortDir('desc'); }
  };

  const toggleExpand = (name) => {
    setExpanded(prev => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name); else next.add(name);
      return next;
    });
  };

  const period = campaigns[0] ? `${campaigns[0].reportStart} ~ ${campaigns[0].reportEnd}` : '';
  const belowCount = campaigns.filter(c => c.spend > 0 && c.roas < targetRoas).length;

  const fresh = ageStats.find(s => s.id === '신규');
  const old = ageStats.find(s => s.id === '노후');
  let fatigueMsg = '';
  if (fresh && old) {
    if (isProductMode) {
      if (fresh.ctr > 0 && old.ctr < fresh.ctr * 0.8) {
        fatigueMsg = `노후 캠페인 CTR(${fmtPct1(old.ctr)})이 신규(${fmtPct1(fresh.ctr)}) 대비 낮습니다 — 광고 피로 신호. 소재 교체를 검토하세요.`;
      } else {
        fatigueMsg = `노후 캠페인 CTR(${fmtPct1(old.ctr)})이 신규(${fmtPct1(fresh.ctr)}) 대비 잘 유지되고 있습니다.`;
      }
    } else if (fresh.roas > 0) {
      if (old.roas < fresh.roas * 0.8) {
        fatigueMsg = `노후 캠페인 ROAS(${old.roas.toFixed(2)})가 신규(${fresh.roas.toFixed(2)}) 대비 낮습니다 — 광고 피로 신호. 오래된 캠페인의 소재 교체를 검토하세요.`;
      } else if (old.roas >= fresh.roas) {
        fatigueMsg = `노후 캠페인 ROAS(${old.roas.toFixed(2)})가 신규(${fresh.roas.toFixed(2)}) 이상으로 유지되고 있습니다 — 검증된 캠페인은 계속 운영해도 좋습니다.`;
      } else {
        fatigueMsg = `노후 캠페인 ROAS(${old.roas.toFixed(2)})가 신규(${fresh.roas.toFixed(2)})보다 다소 낮습니다 — 추이를 지켜보세요.`;
      }
    }
  }

  const cols = isProductMode
    ? [
        { key: 'name', label: '캠페인명', align: 'left' },
        { key: 'manager', label: '담당자', align: 'left' },
        { key: 'postCode', label: '게시', align: 'left' },
        { key: 'ageDays', label: '경과일', align: 'right' },
        { key: 'ageGroup', label: '나이', align: 'left' },
        { key: 'productCount', label: '제품수', align: 'right' },
        { key: 'spend', label: '지출', align: 'right' },
        { key: 'impressions', label: '노출', align: 'right' },
        { key: 'clicks', label: '링크클릭', align: 'right' },
        { key: 'ctr', label: 'CTR', align: 'right', term: 'CTR' },
        { key: 'cpc', label: 'CPC', align: 'right', term: 'CPC' },
      ]
    : [
        { key: 'name', label: '캠페인명', align: 'left' },
        { key: 'manager', label: '담당자', align: 'left' },
        { key: 'postCode', label: '게시', align: 'left' },
        { key: 'ageDays', label: '경과일', align: 'right' },
        { key: 'ageGroup', label: '나이', align: 'left' },
        { key: 'status', label: '상태', align: 'left' },
        { key: 'spend', label: '지출', align: 'right' },
        { key: 'revenue', label: '매출', align: 'right' },
        { key: 'roas', label: 'ROAS', align: 'right', term: 'ROAS' },
        { key: 'purchases', label: '구매', align: 'right' },
        { key: 'carts', label: '장바구니', align: 'right' },
        { key: 'buyRate', label: '구매전환', align: 'right' },
        { key: 'cpc', label: 'CPC', align: 'right', term: 'CPC' },
        { key: 'impressions', label: '노출(역산)', align: 'right' },
        { key: 'ctr', label: 'CTR', align: 'right', term: 'CTR' },
      ];

  const cellValue = (c, key) => {
    switch (key) {
      case 'postCode': return c.postCode ? `${c.postCode.slice(0, 2)}/${c.postCode.slice(2)}` : '-';
      case 'ageDays': return c.ageDays == null ? '-' : `${c.ageDays}일`;
      case 'ageGroup': return c.ageGroup || '-';
      case 'status': return AD_STATUS_LABEL[c.status] || c.status || '-';
      case 'productCount': return `${c.productCount || 0}개`;
      case 'spend': case 'revenue': case 'cpc': case 'cpm': case 'impressions': case 'clicks':
        return fmtWon0(c[key]);
      case 'roas': return (c.roas || 0).toFixed(2);
      case 'buyRate': case 'ctr': return fmtPct1(c[key]);
      case 'purchases': case 'carts': return (c[key] || 0).toLocaleString();
      default: return c[key] || '-';
    }
  };

  const summaryCards = isProductMode
    ? [
        { label: '총 광고비', value: `${fmtWon0(totals.spend)}원` },
        { label: '총 노출', value: fmtWon0(totals.impressions) },
        { label: '총 링크클릭', value: fmtWon0(totals.clicks) },
        { label: '평균 CTR', value: fmtPct1(totals.ctr), hi: true },
        { label: '캠페인 / 제품', value: `${campaigns.length} / ${totalProducts}` },
      ]
    : [
        { label: '총 광고비', value: `${fmtWon0(totals.spend)}원` },
        { label: '총 매출', value: `${fmtWon0(totals.revenue)}원` },
        { label: '통합 ROAS', value: totals.roas.toFixed(2), hi: true },
        { label: '총 구매', value: `${totals.purchases.toLocaleString()}건` },
        { label: '장바구니→구매', value: fmtPct1(totals.buyRate) },
      ];

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <h2 className="text-2xl font-medium tracking-tight">메타 광고 성과{isProductMode ? ' · 제품별' : ''}</h2>
          <p className="text-xs text-stone-500 mt-1">
            {fileName} · 기간 {period} · 캠페인 {campaigns.length}개{isProductMode ? ` · 제품 ${totalProducts}개` : ''}
          </p>
        </div>
        <button
          onClick={onReset}
          className="text-sm text-stone-500 hover:text-stone-900 underline underline-offset-2"
        >
          다른 파일 불러오기
        </button>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
        {summaryCards.map(card => (
          <div
            key={card.label}
            className={`p-4 border ${card.hi ? 'bg-stone-900 text-cream-50 border-stone-900' : 'bg-cream-50 border-cream-400'}`}
          >
            <div className={`text-xs ${card.hi ? 'text-cream-300' : 'text-stone-500'}`}>{card.label}</div>
            <div className="text-xl font-medium mt-1">{card.value}</div>
          </div>
        ))}
      </div>

      <div className="border border-cream-400 bg-cream-50">
        <div className="px-4 py-2.5 border-b border-cream-300 flex items-center justify-between">
          <h3 className="text-sm font-medium text-stone-700">광고 나이별 효율 — 게시 후 경과 기간 단위</h3>
          {ageFilter !== 'all' && (
            <button
              onClick={() => setAgeFilter('all')}
              className="text-xs text-stone-500 hover:text-stone-900 underline underline-offset-2"
            >
              필터 해제
            </button>
          )}
        </div>
        <table className="w-full text-sm">
          <thead className="text-stone-500">
            <tr>
              <th className="px-4 py-2 text-left font-medium">그룹</th>
              <th className="px-3 py-2 text-right font-medium">캠페인</th>
              <th className="px-3 py-2 text-right font-medium">지출</th>
              {isProductMode ? (
                <>
                  <th className="px-3 py-2 text-right font-medium">노출</th>
                  <th className="px-3 py-2 text-right font-medium">링크클릭</th>
                </>
              ) : (
                <>
                  <th className="px-3 py-2 text-right font-medium">매출</th>
                  <th className="px-3 py-2 text-right font-medium">통합 ROAS</th>
                </>
              )}
              <th className="px-3 py-2 text-right font-medium">CTR</th>
            </tr>
          </thead>
          <tbody>
            {ageStats.map(s => {
              const sel = ageFilter === s.id;
              return (
                <tr
                  key={s.id}
                  onClick={() => setAgeFilter(sel ? 'all' : s.id)}
                  className={`border-t border-cream-300 cursor-pointer ${sel ? 'bg-stone-900 text-cream-50' : 'hover:bg-cream-100'}`}
                >
                  <td className="px-4 py-2">
                    {s.id} <span className={sel ? 'text-cream-300' : 'text-stone-400'}>{s.range}</span>
                  </td>
                  <td className="px-3 py-2 text-right">{s.count}개</td>
                  <td className="px-3 py-2 text-right">{fmtWon0(s.spend)}</td>
                  {isProductMode ? (
                    <>
                      <td className="px-3 py-2 text-right">{fmtWon0(s.impressions)}</td>
                      <td className="px-3 py-2 text-right">{fmtWon0(s.clicks)}</td>
                    </>
                  ) : (
                    <>
                      <td className="px-3 py-2 text-right">{fmtWon0(s.revenue)}</td>
                      <td className="px-3 py-2 text-right font-medium">{s.roas.toFixed(2)}</td>
                    </>
                  )}
                  <td className="px-3 py-2 text-right">{fmtPct1(s.ctr)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {fatigueMsg && (
          <p className="px-4 py-2.5 text-xs text-stone-600 border-t border-cream-300 leading-relaxed">
            {fatigueMsg}
          </p>
        )}
      </div>

      {!isProductMode && (
        <div className="flex items-center gap-3 flex-wrap bg-cream-50 border border-cream-400 px-4 py-3">
          <label className="text-sm text-stone-700">목표 ROAS</label>
          <input
            type="number"
            min="0"
            step="0.1"
            value={targetRoas}
            onChange={e => setTargetRoas(parseFloat(e.target.value) || 0)}
            className="w-20 px-2 py-1 text-sm border border-cream-400 bg-cream-100"
          />
          <span className="text-sm text-stone-600">
            미만 캠페인 <span className="font-medium text-rose-700">{belowCount}개</span> — 소재 교체·예산 재배분 검토 대상
          </span>
          {statuses.length > 1 && (
            <div className="ml-auto flex items-center gap-1">
              <span className="text-xs text-stone-500 mr-1">상태</span>
              {statuses.map(s => (
                <button
                  key={s}
                  onClick={() => setStatusFilter(s)}
                  className={`px-2 py-0.5 text-xs border transition ${
                    statusFilter === s
                      ? 'bg-stone-900 text-cream-50 border-stone-900'
                      : 'bg-cream-50 text-stone-700 border-cream-400 hover:border-stone-700'
                  }`}
                >
                  {s === 'all' ? '전체' : (AD_STATUS_LABEL[s] || s)}
                </button>
              ))}
            </div>
          )}
        </div>
      )}

      <div className="border border-cream-400 bg-cream-50 overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-cream-200 text-stone-600">
            <tr>
              <th className="px-3 py-2 text-left font-medium whitespace-nowrap">#</th>
              {cols.map(col => (
                <th
                  key={col.key}
                  onClick={() => toggleSort(col.key)}
                  className={`px-3 py-2 font-medium whitespace-nowrap cursor-pointer hover:text-stone-900 ${
                    col.align === 'right' ? 'text-right' : 'text-left'
                  }`}
                >
                  {col.term ? <InfoTerm term={col.term}>{col.label}</InfoTerm> : col.label}{sortKey === col.key ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((c, i) => {
              const below = !isProductMode && c.spend > 0 && c.roas < targetRoas;
              const isOpen = expanded.has(c.name);
              return (
                <React.Fragment key={c.name + i}>
                  <tr
                    className={`border-t border-cream-300 hover:bg-cream-100 ${isProductMode ? 'cursor-pointer' : ''}`}
                    onClick={isProductMode ? () => toggleExpand(c.name) : undefined}
                  >
                    <td className="px-3 py-2 text-stone-400 whitespace-nowrap">{i + 1}</td>
                    {cols.map(col => (
                      <td
                        key={col.key}
                        title={col.key === 'name' ? c.name : undefined}
                        className={`px-3 py-2 whitespace-nowrap ${col.align === 'right' ? 'text-right' : 'text-left'} ${
                          col.key === 'roas'
                            ? (below ? 'text-rose-700 font-medium' : 'font-medium')
                            : 'text-stone-700'
                        } ${col.key === 'name' ? 'max-w-[260px] truncate' : ''}`}
                      >
                        {col.key === 'name' && isProductMode ? (
                          <span>
                            <span className="text-stone-400 mr-1">{isOpen ? '▾' : '▸'}</span>
                            {cellValue(c, col.key)}
                          </span>
                        ) : cellValue(c, col.key)}
                      </td>
                    ))}
                  </tr>
                  {isProductMode && isOpen && (
                    <tr className="border-t border-cream-300">
                      <td colSpan={cols.length + 1} className="p-0 bg-cream-100">
                        <div className="px-6 py-3">
                          <div className="text-xs font-medium text-stone-600 mb-2">
                            {c.name} · 제품 {c.products.length}개 (지출 많은 순)
                          </div>
                          <table className="w-full text-xs">
                            <thead className="text-stone-500">
                              <tr>
                                <th className="px-2 py-1.5 text-left font-medium">제품명</th>
                                <th className="px-2 py-1.5 text-right font-medium">지출</th>
                                <th className="px-2 py-1.5 text-right font-medium">노출</th>
                                <th className="px-2 py-1.5 text-right font-medium">링크클릭</th>
                                <th className="px-2 py-1.5 text-right font-medium">CTR</th>
                                <th className="px-2 py-1.5 text-right font-medium">CPC</th>
                              </tr>
                            </thead>
                            <tbody>
                              {c.products.map((p, pi) => (
                                <tr key={p.productId + pi} className="border-t border-cream-300">
                                  <td className="px-2 py-1.5 text-stone-700">{p.productName}</td>
                                  <td className="px-2 py-1.5 text-right text-stone-600">{fmtWon0(p.spend)}</td>
                                  <td className="px-2 py-1.5 text-right text-stone-600">{fmtWon0(p.impressions)}</td>
                                  <td className="px-2 py-1.5 text-right text-stone-600">{fmtWon0(p.clicks)}</td>
                                  <td className="px-2 py-1.5 text-right text-stone-600">{fmtPct1(p.ctr)}</td>
                                  <td className="px-2 py-1.5 text-right text-stone-600">{fmtWon0(p.cpc)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>

      <p className="text-xs text-stone-500 leading-relaxed">
        {isProductMode
          ? '제품별 지출·노출·링크클릭·CTR은 정확한 값이에요. 메타는 제품 단위로는 ROAS·매출을 제공하지 않아 이 화면엔 노출·클릭 효율 위주로 표시됩니다. 캠페인명을 클릭하면 제품별 상세가 펼쳐집니다.'
          : '노출·클릭은 CPM·CPC로 역산한 추정치예요. ROAS는 게시 직후 지출이 적은 캠페인일수록 변동이 크니, 지출 규모를 함께 보고 판단하세요.'}
      </p>
    </div>
  );
};

const AdTrackView = ({ adList, adListName, groups, dateLabels, fileName, campaigns: perfCampaigns, campaignsName, onReset }) => {
  const [viewMode, setViewMode] = useState('product');
  const [periodMode, setPeriodMode] = useState('day');
  const [expanded, setExpanded] = useState(() => new Set());
  const [sortKey, setSortKey] = useState('adCount');
  const [sortDir, setSortDir] = useState('desc');
  const [query, setQuery] = useState('');

  const codeMap = useMemo(() => {
    const m = new Map();
    for (const g of groups) {
      const code = extractProductCode(g.productName);
      if (code && !m.has(code)) m.set(code, g);
    }
    return m;
  }, [groups]);

  const perfMap = useMemo(() => {
    const m = new Map();
    for (const c of perfCampaigns || []) {
      if (c.name) m.set(c.name, c);
    }
    return m;
  }, [perfCampaigns]);
  const hasPerf = perfMap.size > 0;

  const periods = useMemo(() => {
    if (periodMode === 'day') {
      return dateLabels.map((d, i) => ({ label: d.slice(5).replace('-', '/'), idx: [i], start: d }));
    }
    const out = [];
    for (let i = 0; i < dateLabels.length; i += 7) {
      const idx = [];
      for (let j = i; j < Math.min(i + 7, dateLabels.length); j++) idx.push(j);
      const s = dateLabels[i].slice(5).replace('-', '/');
      const e = dateLabels[Math.min(i + 6, dateLabels.length - 1)].slice(5).replace('-', '/');
      out.push({ label: `${s}~${e}`, idx, start: dateLabels[i] });
    }
    return out;
  }, [periodMode, dateLabels]);

  const dailyOfGroup = (g, len) => {
    const arr = new Array(len).fill(0);
    for (const sku of g.skus || []) {
      for (let i = 0; i < len; i++) arr[i] += sku.sales[i] || 0;
    }
    return arr;
  };

  const postIdxOf = (postCode) => {
    if (!postCode) return -1;
    const md = postCode.slice(0, 2) + '/' + postCode.slice(2);
    return dateLabels.findIndex(d => d.slice(5).replace('-', '/') === md);
  };

  // 상품 중심
  const products = useMemo(() => {
    const pm = new Map();
    for (const camp of adList) {
      for (const p of camp.products) {
        const codes = p.codes && p.codes.length ? p.codes : ['__' + p.raw];
        for (const code of codes) {
          if (!pm.has(code)) pm.set(code, { code, adName: p.raw, campaigns: [] });
          const prod = pm.get(code);
          if (!prod.campaigns.some(c => c.no === camp.no && c.name === camp.name)) {
            prod.campaigns.push({
              no: camp.no, name: camp.name, manager: camp.manager,
              postCode: camp.postCode, thumbUrl: p.thumbUrl || null,
            });
          }
        }
      }
    }
    const len = dateLabels.length;
    return [...pm.values()].map(prod => {
      const realCode = prod.code.startsWith('__') ? '' : prod.code;
      const g = realCode ? codeMap.get(realCode) : null;
      let daily = new Array(len).fill(0), total = 0, imageUrl = null;
      let productName = prod.adName;
      if (g) {
        daily = dailyOfGroup(g, len);
        total = daily.reduce((a, b) => a + b, 0);
        imageUrl = g.imageUrl;
        productName = g.productName;
      }
      const sortedCamps = prod.campaigns
        .slice()
        .sort((a, b) => (b.postCode || '').localeCompare(a.postCode || ''));
      const campWithEffect = sortedCamps.map(c => {
        let nextDaySales = null;
        if (g) {
          const pi = postIdxOf(c.postCode);
          if (pi >= 0 && pi + 1 < len) nextDaySales = daily[pi + 1];
        }
        const perf = perfMap.get(c.name);
        return {
          ...c, nextDaySales,
          roas: perf?.roas ?? null,
          revenue: perf?.revenue ?? null,
          purchases: perf?.purchases ?? null,
          spend: perf?.spend ?? null,
        };
      });
      const validNext = campWithEffect.filter(c => c.nextDaySales != null).map(c => c.nextDaySales);
      const maxNext = validNext.length ? Math.max(...validNext) : 0;
      const validRoas = campWithEffect.filter(c => c.roas != null && c.roas > 0).map(c => c.roas);
      const maxRoas = validRoas.length ? Math.max(...validRoas) : 0;
      campWithEffect.forEach(c => {
        c.isBest = c.nextDaySales != null && c.nextDaySales === maxNext && maxNext > 0;
        c.isBestRoas = c.roas != null && c.roas === maxRoas && maxRoas > 0;
      });
      return {
        code: realCode, adName: prod.adName, productName, imageUrl,
        matched: !!g, daily, total,
        recent1: g && len > 0 ? daily[len - 1] : 0,
        nextDayMax: g ? maxNext : 0,
        maxRoas,
        campaigns: campWithEffect,
        adCount: campWithEffect.length,
      };
    });
  }, [adList, codeMap, dateLabels, perfMap]);

  // 광고 중심
  const campaignRows = useMemo(() => {
    const len = dateLabels.length;
    return adList.map(camp => {
      const pi = postIdxOf(camp.postCode);
      const prods = camp.products.map(p => {
        const code = (p.codes || []).find(c => codeMap.has(c)) || (p.codes || [])[0] || '';
        const g = code ? codeMap.get(code) : null;
        let daily = new Array(len).fill(0), total = 0, nextDaySales = null;
        let productName = p.raw, imageUrl = null;
        if (g) {
          daily = dailyOfGroup(g, len);
          total = daily.reduce((a, b) => a + b, 0);
          productName = g.productName;
          imageUrl = g.imageUrl;
          if (pi >= 0 && pi + 1 < len) nextDaySales = daily[pi + 1];
        }
        return {
          raw: p.raw, code, codes: p.codes, thumbUrl: p.thumbUrl || null,
          matched: !!g, productName, imageUrl, daily, total, nextDaySales,
        };
      });
      const validNext = prods.filter(x => x.nextDaySales != null).map(x => x.nextDaySales);
      const maxNext = validNext.length ? Math.max(...validNext) : 0;
      prods.forEach(x => {
        x.isBest = x.nextDaySales != null && x.nextDaySales === maxNext && maxNext > 0;
      });
      const perf = perfMap.get(camp.name);
      return {
        no: camp.no, name: camp.name, manager: camp.manager,
        postCode: camp.postCode, status: camp.status,
        prods,
        productCount: prods.length,
        matchedCount: prods.filter(x => x.matched).length,
        nextDayTotal: prods.reduce((s, x) => s + (x.nextDaySales || 0), 0),
        bestName: (prods.find(x => x.isBest) || {}).productName || '',
        roas: perf?.roas ?? null,
        revenue: perf?.revenue ?? null,
        purchases: perf?.purchases ?? null,
        spend: perf?.spend ?? null,
      };
    });
  }, [adList, codeMap, dateLabels, perfMap]);

  const sortedProducts = useMemo(() => {
    let list = products;
    if (query.trim()) {
      const q = query.trim().toLowerCase();
      list = list.filter(p =>
        (p.productName || '').toLowerCase().includes(q) ||
        (p.code || '').toLowerCase().includes(q) ||
        (p.adName || '').toLowerCase().includes(q)
      );
    }
    return [...list].sort((a, b) => {
      const av = a[sortKey], bv = b[sortKey];
      let cmp;
      if (typeof av === 'string') cmp = String(av).localeCompare(String(bv));
      else cmp = (av || 0) - (bv || 0);
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }, [products, sortKey, sortDir, query]);

  const sortedCampaigns = useMemo(() => {
    let list = campaignRows;
    if (query.trim()) {
      const q = query.trim().toLowerCase();
      list = list.filter(c =>
        (c.name || '').toLowerCase().includes(q) ||
        (c.manager || '').toLowerCase().includes(q)
      );
    }
    return [...list].sort((a, b) => (b.postCode || '').localeCompare(a.postCode || ''));
  }, [campaignRows, query]);

  const toggleSort = (k) => {
    if (sortKey === k) setSortDir(d => (d === 'asc' ? 'desc' : 'asc'));
    else { setSortKey(k); setSortDir('desc'); }
  };
  const toggleExpand = (key) => {
    setExpanded(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  };

  const period = dateLabels.length ? `${dateLabels[0]} ~ ${dateLabels[dateLabels.length - 1]}` : '';
  const fmtPost = (pc) => pc ? `${pc.slice(0, 2)}/${pc.slice(2)}` : '-';
  const matchedCount = products.filter(p => p.matched).length;
  const hasThumb = adList.some(c => c.products.some(p => p.thumbUrl));

  const ProductCard = ({ thumbUrl, title, sub, nextDaySales, isBest, roas, isBestRoas }) => (
    <div className={`w-40 bg-cream-50 ${isBest ? 'border-4 border-stone-900' : 'border border-cream-400'}`}>
      <div className="w-full h-40 bg-cream-200 overflow-hidden flex items-center justify-center relative">
        {thumbUrl ? (
          <img src={thumbUrl} alt="" className="w-full h-full object-cover" />
        ) : (
          <span className="text-xs text-stone-400">이미지 없음</span>
        )}
        {isBest && (
          <span className="absolute top-0 left-0 bg-stone-900 text-cream-50 text-[10px] px-1.5 py-0.5 font-medium">
            BEST
          </span>
        )}
      </div>
      <div className="px-2 py-1.5">
        <div className="text-xs font-medium text-stone-700 truncate" title={title}>{title}</div>
        <div className="text-[11px] text-stone-500 mt-0.5">{sub}</div>
        <div className={`text-[11px] mt-0.5 ${isBest ? 'text-stone-900 font-medium' : 'text-stone-500'}`}>
          게시 다음날 판매 {nextDaySales == null ? '—' : nextDaySales}
        </div>
        {hasPerf && (
          <div className={`text-[11px] mt-0.5 ${isBestRoas ? 'text-amber-700 font-medium' : 'text-stone-500'}`}>
            <InfoTerm term="ROAS" /> {roas == null ? '—' : roas.toFixed(2)}{isBestRoas ? ' ★' : ''}
          </div>
        )}
      </div>
    </div>
  );

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <h2 className="text-2xl font-medium tracking-tight">
            광고-판매 추적 · {viewMode === 'product' ? '상품별' : '광고별'}
          </h2>
          <p className="text-xs text-stone-500 mt-1">
            광고 {adListName} · 매출 {fileName}{hasPerf ? ` · 성과 ${campaignsName}` : ''} · 기간 {period} · 상품 {products.length}개 (매출매칭 {matchedCount}) · 캠페인 {campaignRows.length}개
          </p>
        </div>
        <button
          onClick={onReset}
          className="text-sm text-stone-500 hover:text-stone-900 underline underline-offset-2"
        >
          다른 파일 불러오기
        </button>
      </div>

      {!hasThumb && (
        <div className="bg-cream-200 border border-cream-400 px-4 py-2.5 text-xs text-stone-600">
          광고리스트가 CSV라 광고 이미지가 없어요. 이미지를 보려면 <span className="font-medium">.xlsx 광고리스트</span>로 올려주세요.
        </div>
      )}

      <div className="flex items-center gap-3 bg-cream-50 border border-cream-400 px-4 py-3 flex-wrap">
        <div className="flex items-center border border-cream-400">
          {[['product', '상품별'], ['campaign', '광고별']].map(([v, label]) => (
            <button
              key={v}
              onClick={() => { setViewMode(v); setExpanded(new Set()); }}
              className={`px-3 py-1.5 text-sm transition ${
                viewMode === v ? 'bg-stone-900 text-cream-50' : 'bg-cream-50 text-stone-700 hover:bg-cream-200'
              }`}
            >
              {label} 보기
            </button>
          ))}
        </div>
        <div className="flex items-center gap-1.5 bg-cream-100 px-2.5 py-1.5 border border-cream-400">
          <LucideSearch size={13} className="text-stone-500" />
          <input
            type="text"
            value={query}
            onChange={e => setQuery(e.target.value)}
            placeholder={viewMode === 'product' ? '상품명·코드 검색' : '캠페인·담당자 검색'}
            className="text-sm bg-transparent focus:outline-none w-40 placeholder:text-stone-400"
          />
          {query && (
            <button onClick={() => setQuery('')} className="text-stone-400 hover:text-stone-700">
              <LucideX size={13} />
            </button>
          )}
        </div>
        {viewMode === 'product' && (
          <div className="flex items-center border border-cream-400">
            <span className="px-2 text-xs text-stone-500">기간 단위</span>
            {[['day', '일별'], ['week', '주별']].map(([v, label]) => (
              <button
                key={v}
                onClick={() => setPeriodMode(v)}
                className={`px-3 py-1.5 text-sm transition ${
                  periodMode === v ? 'bg-stone-900 text-cream-50' : 'bg-cream-50 text-stone-700 hover:bg-cream-200'
                }`}
              >
                {label}
              </button>
            ))}
          </div>
        )}
        <span className="text-xs text-stone-500">
          {viewMode === 'product'
            ? '상품을 클릭하면 그 상품을 쓴 광고 이미지들과 기간별 판매가 펼쳐집니다.'
            : '광고를 클릭하면 그 광고에 담긴 상품들이 펼쳐지고, 게시 다음날 판매가 가장 높은 상품이 강조됩니다.'}
        </span>
      </div>

      {viewMode === 'product' ? (
        <div className="border border-cream-400 bg-cream-50 overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-cream-200 text-stone-600">
              <tr>
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">#</th>
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">이미지</th>
                <th onClick={() => toggleSort('productName')} className="px-3 py-2 text-left font-medium whitespace-nowrap cursor-pointer hover:text-stone-900">
                  상품명{sortKey === 'productName' ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                </th>
                <th onClick={() => toggleSort('code')} className="px-3 py-2 text-left font-medium whitespace-nowrap cursor-pointer hover:text-stone-900">코드</th>
                <th onClick={() => toggleSort('adCount')} className="px-3 py-2 text-right font-medium whitespace-nowrap cursor-pointer hover:text-stone-900">
                  광고 횟수{sortKey === 'adCount' ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                </th>
                <th onClick={() => toggleSort('total')} className="px-3 py-2 text-right font-medium whitespace-nowrap cursor-pointer hover:text-stone-900">
                  기간 판매{sortKey === 'total' ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                </th>
                <th onClick={() => toggleSort('nextDayMax')} className="px-3 py-2 text-right font-medium whitespace-nowrap cursor-pointer hover:text-stone-900" title="이 상품이 들어간 광고들의 게시 다음날 판매 중 최댓값">
                  광고 다음날 최다{sortKey === 'nextDayMax' ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                </th>
              </tr>
            </thead>
            <tbody>
              {sortedProducts.map((p, i) => {
                const key = 'P' + (p.code || p.adName);
                const isOpen = expanded.has(key);
                return (
                  <React.Fragment key={key + i}>
                    <tr className="border-t border-cream-300 hover:bg-cream-100 cursor-pointer" onClick={() => toggleExpand(key)}>
                      <td className="px-3 py-2 text-stone-400 whitespace-nowrap">{i + 1}</td>
                      <td className="px-3 py-2">
                        <div className="w-12 h-12 bg-cream-200 border border-cream-400 overflow-hidden flex items-center justify-center">
                          {p.imageUrl && p.imageUrl !== '이미지없음' ? (
                            <img src={p.imageUrl} alt="" className="w-full h-full object-cover" onError={e => { e.target.style.display = 'none'; }} />
                          ) : (
                            <span className="text-[10px] text-stone-400">없음</span>
                          )}
                        </div>
                      </td>
                      <td className="px-3 py-2 whitespace-nowrap max-w-[340px] truncate" title={p.productName}>
                        <span className="text-stone-400 mr-1">{isOpen ? '▾' : '▸'}</span>
                        {p.productName}
                        {!p.matched && <span className="text-rose-600 text-xs ml-1">(매출 미매칭)</span>}
                      </td>
                      <td className="px-3 py-2 text-stone-500 whitespace-nowrap">{p.code || '-'}</td>
                      <td className="px-3 py-2 text-right whitespace-nowrap">
                        <span className="inline-flex items-center justify-center min-w-[1.6rem] px-1.5 py-0.5 bg-stone-900 text-cream-50 text-xs font-medium">
                          {p.adCount}
                        </span>
                      </td>
                      <td className="px-3 py-2 text-right font-medium whitespace-nowrap">
                        {p.matched ? p.total.toLocaleString() : '—'}
                      </td>
                      <td className="px-3 py-2 text-right whitespace-nowrap text-stone-700">
                        {p.matched && p.nextDayMax > 0 ? p.nextDayMax.toLocaleString() : '—'}
                      </td>
                    </tr>
                    {isOpen && (
                      <tr className="border-t border-cream-300">
                        <td colSpan={7} className="p-0 bg-cream-100">
                          <div className="px-6 py-4 space-y-4">
                            <div>
                              <div className="text-xs font-medium text-stone-600 mb-2">
                                이 상품이 들어간 광고 {p.campaigns.length}개 — 광고별 사용 이미지
                                <span className="text-stone-400 font-normal"> (두꺼운 테두리 = 게시 다음날 판매가 가장 높았던 광고)</span>
                              </div>
                              <div className="flex flex-wrap gap-3">
                                {p.campaigns.map((c, ci) => (
                                  <ProductCard
                                    key={ci}
                                    thumbUrl={c.thumbUrl}
                                    title={c.name}
                                    sub={`게시 ${fmtPost(c.postCode)} · ${c.manager}`}
                                    nextDaySales={c.nextDaySales}
                                    isBest={c.isBest}
                                    roas={c.roas}
                                    isBestRoas={c.isBestRoas}
                                  />
                                ))}
                              </div>
                            </div>
                            {p.matched ? (
                              <div className="overflow-x-auto">
                                <div className="text-xs font-medium text-stone-600 mb-2">
                                  {periodMode === 'day' ? '일별' : '주별'} 판매량
                                  <span className="text-stone-400 font-normal"> (▲ = 광고 게시일)</span>
                                </div>
                                <table className="text-xs">
                                  <thead className="text-stone-500">
                                    <tr>
                                      {periods.map((pd, idx) => {
                                        const marker = p.campaigns.some(c => {
                                          if (!c.postCode) return false;
                                          const md = c.postCode.slice(0, 2) + '/' + c.postCode.slice(2);
                                          return pd.idx.some(di => dateLabels[di].slice(5).replace('-', '/') === md);
                                        });
                                        return (
                                          <th key={idx} className="px-2 py-1.5 text-right font-medium whitespace-nowrap">
                                            {marker ? <span className="text-stone-900">▲ </span> : ''}{pd.label}
                                          </th>
                                        );
                                      })}
                                      <th className="px-2 py-1.5 text-right font-medium">합계</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    <tr className="border-t border-cream-300">
                                      {periods.map((pd, idx) => {
                                        const v = pd.idx.reduce((a, di) => a + (p.daily[di] || 0), 0);
                                        return (
                                          <td key={idx} className={`px-2 py-1.5 text-right whitespace-nowrap ${v ? 'text-stone-700' : 'text-stone-400'}`}>
                                            {v}
                                          </td>
                                        );
                                      })}
                                      <td className="px-2 py-1.5 text-right font-medium text-stone-700">{p.total.toLocaleString()}</td>
                                    </tr>
                                  </tbody>
                                </table>
                              </div>
                            ) : (
                              <div className="text-xs text-stone-500">매출 파일에 이 코드({p.code || '코드 없음'})의 상품이 없어 판매량을 표시할 수 없어요.</div>
                            )}
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="border border-cream-400 bg-cream-50 overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-cream-200 text-stone-600">
              <tr>
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">#</th>
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">캠페인명</th>
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">담당자</th>
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">게시일</th>
                <th className="px-3 py-2 text-right font-medium whitespace-nowrap">상품수</th>
                <th className="px-3 py-2 text-right font-medium whitespace-nowrap">매칭</th>
                {hasPerf && <th className="px-3 py-2 text-right font-medium whitespace-nowrap"><InfoTerm term="ROAS" /></th>}
                {hasPerf && <th className="px-3 py-2 text-right font-medium whitespace-nowrap">매출</th>}
                {hasPerf && <th className="px-3 py-2 text-right font-medium whitespace-nowrap">구매</th>}
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">게시 다음날 효율 1위</th>
              </tr>
            </thead>
            <tbody>
              {sortedCampaigns.map((c, i) => {
                const key = 'C' + c.no + c.name;
                const isOpen = expanded.has(key);
                return (
                  <React.Fragment key={key + i}>
                    <tr className="border-t border-cream-300 hover:bg-cream-100 cursor-pointer" onClick={() => toggleExpand(key)}>
                      <td className="px-3 py-2 text-stone-400 whitespace-nowrap">{i + 1}</td>
                      <td className="px-3 py-2 whitespace-nowrap max-w-[280px] truncate" title={c.name}>
                        <span className="text-stone-400 mr-1">{isOpen ? '▾' : '▸'}</span>{c.name}
                      </td>
                      <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{c.manager}</td>
                      <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{fmtPost(c.postCode)}</td>
                      <td className="px-3 py-2 text-right text-stone-600 whitespace-nowrap">{c.productCount}</td>
                      <td className="px-3 py-2 text-right text-stone-600 whitespace-nowrap">{c.matchedCount}/{c.productCount}</td>
                      {hasPerf && (
                        <td className="px-3 py-2 text-right font-medium text-stone-700 whitespace-nowrap">
                          {c.roas == null ? '—' : c.roas.toFixed(2)}
                        </td>
                      )}
                      {hasPerf && (
                        <td className="px-3 py-2 text-right text-stone-600 whitespace-nowrap">
                          {c.revenue == null ? '—' : Math.round(c.revenue).toLocaleString()}
                        </td>
                      )}
                      {hasPerf && (
                        <td className="px-3 py-2 text-right text-stone-600 whitespace-nowrap">
                          {c.purchases == null ? '—' : c.purchases.toLocaleString()}
                        </td>
                      )}
                      <td className="px-3 py-2 text-stone-700 whitespace-nowrap max-w-[260px] truncate" title={c.bestName}>
                        {c.bestName || '—'}
                      </td>
                    </tr>
                    {isOpen && (
                      <tr className="border-t border-cream-300">
                        <td colSpan={hasPerf ? 10 : 7} className="p-0 bg-cream-100">
                          <div className="px-6 py-4">
                            <div className="text-xs font-medium text-stone-600 mb-2">
                              {c.name} · 담긴 상품 {c.productCount}개
                              <span className="text-stone-400 font-normal"> (두꺼운 테두리 = 이 광고에서 게시 다음날 판매가 가장 높았던 상품)</span>
                            </div>
                            <div className="flex flex-wrap gap-3">
                              {c.prods.map((pr, pi) => (
                                <ProductCard
                                  key={pi}
                                  thumbUrl={pr.thumbUrl}
                                  title={pr.productName}
                                  sub={pr.code || '코드 없음'}
                                  nextDaySales={pr.matched ? pr.nextDaySales : null}
                                  isBest={pr.isBest}
                                />
                              ))}
                            </div>
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      <p className="text-xs text-stone-500 leading-relaxed">
        광고리스트의 상품과 매출 파일을 상품코드(영문+숫자)로 매칭했어요.
        {viewMode === 'product'
          ? ' 상품을 펼치면 광고마다 사용한 이미지를 비교할 수 있고, 게시 다음날 판매가 가장 높았던 광고가 강조됩니다.'
          : ' 광고를 펼치면 담긴 상품들이 나오고, 그 광고 게시 다음날 판매가 가장 높았던 상품이 강조됩니다.'}
      </p>
    </div>
  );
};



const Panel = ({ title, icon: Icon, children }) => (
  <div className="bg-cream-50 border border-cream-400">
    <div className="px-4 py-3 border-b border-cream-400 flex items-center gap-2">
      {Icon && <Icon size={14} strokeWidth={1.5} className="text-stone-600" />}
      <h3 className="text-sm font-medium text-stone-800 tracking-tight">{title}</h3>
    </div>
    <div className="p-4">{children}</div>
  </div>
);

const ThemeOptions = ({
  theme, opts, setOpts, categories, brands,
  customQuery, setCustomQuery, apiKey, saveApiKey,
  customLoading, customError, customSpec, customResultsCount, onRunCustom,
}) => {
  const set = (k, v) => setOpts({ ...opts, [k]: v });

  if (theme === 'custom') {
    return (
      <div className="space-y-3">
        <div>
          <label className="text-xs text-stone-600 block mb-1">조건 (자연어)</label>
          <textarea
            value={customQuery}
            onChange={e => setCustomQuery(e.target.value)}
            rows={4}
            placeholder={'예: 반바지 카테고리 중 재고 50개 이상이고\n판매량 5개 이상인 S/S 상품을\n매출액 높은 순으로'}
            className="w-full px-2 py-2 text-sm border border-cream-400 bg-cream-50 focus:outline-none focus:border-stone-700 leading-snug resize-y"
          />
        </div>
        <details>
          <summary className="text-xs text-stone-600 cursor-pointer hover:text-stone-900 flex items-center gap-1">
            <LucideKey size={11} /> Gemini API 키 {apiKey ? '(저장됨)' : '(필요)'}
          </summary>
          <input
            type="password"
            value={apiKey}
            onChange={e => saveApiKey(e.target.value)}
            placeholder="AIzaSy..."
            className="w-full mt-2 px-2 py-1.5 text-xs border border-cream-400 bg-cream-50 focus:outline-none focus:border-stone-700"
          />
          <p className="text-xs text-stone-500 mt-1 leading-snug">
            <a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer" className="underline">aistudio.google.com</a>에서 무료 발급. 브라우저에만 저장돼.
          </p>
        </details>
        <button
          onClick={onRunCustom}
          disabled={customLoading || !customQuery.trim() || !apiKey.trim()}
          className="w-full bg-stone-900 hover:bg-stone-800 disabled:bg-cream-300 disabled:text-stone-400 text-cream-50 py-2 text-sm font-medium flex items-center justify-center gap-2"
        >
          {customLoading ? (
            <><LucideLoader2 size={14} className="animate-spin" /> 분석 중...</>
          ) : (
            <>조건으로 추출</>
          )}
        </button>
        {customError && (
          <p className="text-xs text-rose-700">{customError}</p>
        )}
        {customSpec && !customError && (
          <div className="bg-cream-200 border border-cream-400 px-3 py-2 text-xs text-stone-700 leading-relaxed">
            <div className="font-medium text-stone-900 mb-1">{customSpec.summary || '추출 완료'}</div>
            <div>총 {customResultsCount}개 결과</div>
          </div>
        )}
      </div>
    );
  }

  if (theme === 'category') {
    const selected = opts.categories || [];
    const toggle = (c) => {
      const next = selected.includes(c) ? selected.filter(x => x !== c) : [...selected, c];
      set('categories', next);
    };
    return (
      <div>
        <div className="flex items-center justify-between mb-2">
          <label className="text-xs text-stone-600">
            카테고리 ({categories.length}개 중 {selected.length})
          </label>
          <div className="flex gap-1.5">
            <button
              onClick={() => set('categories', categories)}
              className="text-xs text-stone-600 hover:text-stone-900 underline underline-offset-2"
            >전체</button>
            <span className="text-stone-300">·</span>
            <button
              onClick={() => set('categories', [])}
              className="text-xs text-stone-600 hover:text-stone-900 underline underline-offset-2"
            >해제</button>
          </div>
        </div>
        <div className="flex flex-wrap gap-1 max-h-60 overflow-y-auto py-1">
          {categories.map(c => {
            const active = selected.includes(c);
            return (
              <button
                key={c}
                onClick={() => toggle(c)}
                className={`px-2.5 py-1 text-xs border transition ${
                  active
                    ? 'bg-stone-900 text-cream-50 border-stone-900'
                    : 'bg-cream-50 text-stone-700 border-cream-400 hover:border-stone-700'
                }`}
              >
                {c}
              </button>
            );
          })}
        </div>
        {selected.length === 0 && (
          <p className="text-xs text-rose-700 mt-2">한 개 이상 선택해야 결과가 나와요</p>
        )}
      </div>
    );
  }
  if (theme === 'brand') {
    return (
      <div>
        <label className="text-xs text-stone-600 block mb-1 font-medium">
          브랜드 선택 ({brands.length}개)
        </label>
        <select
          value={opts.brand}
          onChange={e => set('brand', e.target.value)}
          className="w-full px-2 py-2 text-sm border border-cream-400 bg-cream-50 focus:outline-none focus:border-stone-700"
        >
          <option value="">— 브랜드를 선택하세요 —</option>
          {brands.map(b => <option key={b} value={b}>{b}</option>)}
        </select>
        <p className="text-xs text-stone-500 mt-2">
          상품명 앞 prefix(FP·JM·WV·PS·EZ·TWN·PL·DY) 기준
        </p>
        {!opts.brand && (
          <p className="text-xs text-rose-700 mt-2">브랜드를 선택해야 결과가 나와요</p>
        )}
      </div>
    );
  }
  if (theme === 'newProduct') {
    return (
      <div>
        <label className="text-xs text-stone-600 block mb-1">최근 N개월 이내 등록</label>
        <input
          type="number"
          min="1"
          max="60"
          value={opts.newMonths}
          onChange={e => set('newMonths', parseInt(e.target.value) || 6)}
          className="w-24 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
        />
        <span className="text-xs text-stone-500 ml-1">개월</span>
      </div>
    );
  }
  if (theme === 'steady') {
    return (
      <div className="space-y-3">
        <div>
          <label className="text-xs text-stone-600 block mb-1">등록 후 최소 N개월 경과</label>
          <input
            type="number"
            min="6"
            max="120"
            value={opts.minMonths}
            onChange={e => set('minMonths', parseInt(e.target.value) || 24)}
            className="w-24 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
          />
          <span className="text-xs text-stone-500 ml-1">개월</span>
        </div>
        <div>
          <label className="text-xs text-stone-600 block mb-1">일평균 판매량 임계치</label>
          <input
            type="number"
            min="0"
            step="0.1"
            value={opts.minAvgDaily}
            onChange={e => set('minAvgDaily', parseFloat(e.target.value) || 0)}
            className="w-24 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
          />
          <span className="text-xs text-stone-500 ml-1">개/일 이상</span>
        </div>
        <p className="text-xs text-stone-500 leading-relaxed">
          점수 = 일평균수량 × log(1 + 햇수). .xls는 기간 판매합계를 일수로 나눠 계산하고,
          누적 .csv는 파일의 일평균수량을 그대로 써요.
        </p>
      </div>
    );
  }
  if (theme === 'rising') {
    return (
      <div>
        <label className="text-xs text-stone-600 block mb-1">최소 판매량</label>
        <input
          type="number"
          min="0"
          value={opts.minSales}
          onChange={e => set('minSales', parseInt(e.target.value) || 0)}
          className="w-24 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
        />
        <span className="text-xs text-stone-500 ml-1">개 이상</span>
      </div>
    );
  }
  if (theme === 'frequency') {
    return (
      <div className="space-y-3">
        <div>
          <label className="text-xs text-stone-600 block mb-1">최소 판매일수 (8일 중)</label>
          <input
            type="number"
            min="1"
            max="8"
            value={opts.minSalesDays}
            onChange={e => set('minSalesDays', parseInt(e.target.value) || 3)}
            className="w-24 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
          />
          <span className="text-xs text-stone-500 ml-1">일 이상</span>
        </div>
        <p className="text-xs text-stone-500 leading-relaxed">
          점수 = 판매일수 × log(1 + 총판매). 며칠 내내 꾸준히 팔린 상품일수록 상위 → 광고 ROAS 예측이 안정적.
        </p>
      </div>
    );
  }
  if (theme === 'declining') {
    return (
      <div className="space-y-3">
        <div>
          <label className="text-xs text-stone-600 block mb-1">초반 구간 최소 판매량</label>
          <input
            type="number"
            min="1"
            value={opts.minEarly}
            onChange={e => set('minEarly', parseInt(e.target.value) || 3)}
            className="w-24 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
          />
          <span className="text-xs text-stone-500 ml-1">개 이상</span>
        </div>
        <p className="text-xs text-stone-500 leading-relaxed">
          데이터 기간 초반 10% 구간엔 잘 팔렸는데 후반 10% 구간에 판매가 떨어진 상품을 우선. 점수 = 감소율 × log(1+초반판매)
        </p>
      </div>
    );
  }
  if (theme === 'overstock') {
    return (
      <div className="space-y-3">
        <div>
          <label className="text-xs text-stone-600 block mb-1">최소 현재재고</label>
          <input
            type="number"
            min="1"
            value={opts.minStock}
            onChange={e => set('minStock', parseInt(e.target.value) || 30)}
            className="w-24 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
          />
          <span className="text-xs text-stone-500 ml-1">개 이상</span>
        </div>
        <div>
          <label className="text-xs text-stone-600 block mb-1">최대 판매량</label>
          <input
            type="number"
            min="0"
            value={opts.maxSales}
            onChange={e => set('maxSales', parseInt(e.target.value) || 10)}
            className="w-24 px-2 py-1 text-sm border border-cream-400 bg-cream-50"
          />
          <span className="text-xs text-stone-500 ml-1">개 이하</span>
        </div>
        <p className="text-xs text-stone-500 leading-relaxed">
          점수 = 재고 ÷ (판매+1). 재고 많고 판매 적을수록 상위.
        </p>
      </div>
    );
  }
  return <p className="text-xs text-stone-500">이 주제는 별도 옵션이 없어요.</p>;
};

const Preview = ({ items, theme, dateLabels, mode, opts, excludedCount, onExclude, onResetExcluded }) => {
  if (items.length === 0) {
    return (
      <div className="bg-cream-50 border border-cream-400 p-16 text-center">
        <LucideAlertCircle strokeWidth={1.2} className="mx-auto text-stone-400 mb-4" size={32} />
        <p className="text-xl text-stone-800">선정할 상품이 없어요</p>
        <p className="text-sm text-stone-500 mt-2 font-light">필터/옵션을 조정해 보세요.</p>
      </div>
    );
  }
  return (
    <div className="bg-cream-50 border border-cream-400">
      <div className="px-6 py-5 border-b border-cream-400 flex items-center justify-between gap-4 flex-wrap">
        <div>
          <h2 className="text-2xl font-medium tracking-tight">미리보기</h2>
          <p className="text-xs text-stone-600 mt-1 font-light">
            {items.length}개 상품 · {dateLabels.length > 0
              ? `데이터 기간: ${dateLabels[0]} ~ ${dateLabels[dateLabels.length - 1]}`
              : '상품별 누적 데이터 (옵션·이미지 정보 없음)'}
          </p>
        </div>
        {excludedCount > 0 && (
          <div className="flex items-center gap-2 text-xs">
            <span className="text-stone-600">제외 {excludedCount}개 (다음 후보로 자동 보충됨)</span>
            <button
              onClick={onResetExcluded}
              className="text-stone-700 hover:text-stone-900 underline underline-offset-2"
            >
              초기화
            </button>
          </div>
        )}
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-cream-100 text-stone-600 border-b border-cream-400">
            <tr>
              <th className="px-3 py-2 text-left font-medium w-10 whitespace-nowrap">#</th>
              <th className="px-3 py-2 text-left font-medium w-16 whitespace-nowrap">이미지</th>
              <th className="px-3 py-2 text-left font-medium whitespace-nowrap">상품명</th>
              <th className="px-3 py-2 text-left font-medium whitespace-nowrap">카테고리</th>
              <th className="px-3 py-2 text-left font-medium whitespace-nowrap">시즌</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">총 판매</th>
              <th className="px-3 py-2 text-left font-medium whitespace-nowrap">베스트 SKU 1위</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">1위 판매</th>
              <th className="px-3 py-2 text-left font-medium whitespace-nowrap">베스트 SKU 2위</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">2위 판매</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">현재 재고</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">매출액</th>
              {theme === 'recommend' && (
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">선정 사유</th>
              )}
              <th className="px-2 py-2 text-center font-medium whitespace-nowrap w-10">제외</th>
            </tr>
          </thead>
          <tbody>
            {items.map((it, i) => (
              <tr key={it.productName} className="border-t border-cream-300 hover:bg-cream-100 h-[68px]">
                <td className="px-3 py-2 font-medium text-stone-500 whitespace-nowrap">{i + 1}</td>
                <td className="px-3 py-2 whitespace-nowrap">
                  <div className="w-12 h-12 bg-cream-200 border border-cream-400 overflow-hidden flex items-center justify-center">
                    {it.imageUrl && it.imageUrl !== '이미지없음' ? (
                      <img
                        src={it.imageUrl}
                        alt=""
                        className="w-full h-full object-cover"
                        onError={(e) => { e.target.style.display = 'none'; }}
                      />
                    ) : (
                      <span className="text-xs text-stone-400">없음</span>
                    )}
                  </div>
                </td>
                <td className="px-3 py-2 font-medium whitespace-nowrap">{it.productName}</td>
                <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{it.category}</td>
                <td className="px-3 py-2 text-stone-600 text-xs whitespace-nowrap">{it.season || '-'}</td>
                <td className="px-3 py-2 text-right font-medium whitespace-nowrap">{it.totalSales}</td>
                <td className="px-3 py-2 text-stone-600 text-xs whitespace-nowrap">{it.bestSku?.optionName || '-'}</td>
                <td className="px-3 py-2 text-right whitespace-nowrap">{it.bestSku?.salesTotal || 0}</td>
                <td className="px-3 py-2 text-stone-600 text-xs whitespace-nowrap">{it.bestSku2?.optionName || '-'}</td>
                <td className="px-3 py-2 text-right whitespace-nowrap">{it.bestSku2?.salesTotal || 0}</td>
                <td className="px-3 py-2 text-right text-stone-600 whitespace-nowrap">{it.totalCurrentStock.toLocaleString()}</td>
                <td className="px-3 py-2 text-right text-stone-600 whitespace-nowrap">{it.totalRevenue.toLocaleString()}</td>
                {theme === 'recommend' && (
                  <td className="px-3 py-2 text-xs whitespace-nowrap">
                    <span className="inline-flex px-2 py-0.5 bg-stone-900 text-cream-50 font-medium">
                      {reasonLabel(it.pickReason)}
                    </span>
                  </td>
                )}
                <td className="px-2 py-2 text-center whitespace-nowrap">
                  <button
                    onClick={() => onExclude(it.productName)}
                    title="이 상품 빼고 다음 후보로 채우기"
                    className="w-7 h-7 inline-flex items-center justify-center text-stone-400 hover:bg-stone-900 hover:text-cream-50 border border-cream-400 hover:border-stone-900 transition"
                  >
                    <LucideX size={14} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="px-6 py-5 border-t border-cream-300 bg-cream-100">
        <h3 className="text-sm font-semibold text-stone-700 mb-2 tracking-tight">선정 사유 브리핑</h3>
        <p className="text-sm text-stone-700 leading-relaxed">
          {buildBriefing(theme, opts, mode, items)}
        </p>
      </div>
    </div>
  );
};

export default App;
