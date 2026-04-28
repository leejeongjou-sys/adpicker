import React, { useState, useMemo, useCallback, useRef } from 'react';
import {
  LucideUpload, LucideFileSpreadsheet, LucideDownload, LucideSparkles,
  LucideTrendingUp, LucideTrendingDown, LucideStar, LucideCalendar, LucideShirt, LucideTag,
  LucidePackage, LucideAlertCircle, LucideX, LucideImage, LucideBox, LucideArchive, LucideAward
} from 'lucide-react';
import ExcelJS from 'exceljs';

const THEMES = [
  { id: 'recommend', label: '추천 (통합 8개)', desc: '베스트·급상승·신상·재활성·재고소진 분산 픽', icon: LucideAward, modes: ['xls'] },
  { id: 'bestseller', label: '베스트셀러', desc: '7일 총 판매량 Top 8', icon: LucideStar, modes: ['xls', 'csv'] },
  { id: 'rising', label: '급상승(라이징)', desc: '후반 4일 vs 전반 4일 증가율', icon: LucideTrendingUp, modes: ['xls'] },
  { id: 'declining', label: '판매 감소', desc: '후반에 판매가 떨어진 상품 (재활성용)', icon: LucideTrendingDown, modes: ['xls'] },
  { id: 'package', label: '패키지 베스트', desc: '상품명에 PACK 포함된 상품 Top 8', icon: LucideBox, modes: ['csv'] },
  { id: 'newProduct', label: '신상품 베스트', desc: '최근 N개월 등록 + 판매량', icon: LucideSparkles, modes: ['xls', 'csv'] },
  { id: 'category', label: '카테고리별 베스트', desc: '카테고리 선택 → Top 8', icon: LucideShirt, modes: ['xls', 'csv'] },
  { id: 'brand', label: '브랜드별 베스트', desc: '상품명 prefix 코드로 분류', icon: LucideTag, modes: ['xls', 'csv'] },
  { id: 'steady', label: '스테디셀러', desc: '오래됐지만 꾸준한 상품', icon: LucidePackage, modes: ['xls', 'csv'] },
  { id: 'overstock', label: '재고 과다', desc: '재고 많고 안 팔리는 상품 (재고 소진용)', icon: LucideArchive, modes: ['xls'] },
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
    const half = Math.floor(s.sales.length / 2);
    for (let i = 0; i < s.sales.length; i++) {
      if (i < half) g.earlySales += s.sales[i];
      else g.lateSales += s.sales[i];
    }
  }
  for (const g of map.values()) {
    const sortedSkus = [...g.skus].sort((a, b) => b.salesTotal - a.salesTotal);
    g.bestSku = sortedSkus[0] || null;
    g.bestSku2 = sortedSkus[1] && sortedSkus[1].salesTotal > 0 ? sortedSkus[1] : null;
    g.growthRate = g.earlySales > 0 ? g.lateSales / g.earlySales : (g.lateSales > 0 ? 99 : 0);
    g.cancelRate = g.totalSales > 0 ? g.totalCanceled / g.totalSales : 0;
    const numDays = g.skus[0]?.sales.length || 1;
    g.avgDaily = g.totalSales / numDays;
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
      if (sales > (opts.maxSales ?? 5)) return -1;
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
  { theme: 'bestseller', n: 3 },
  { theme: 'rising', n: 2 },
  { theme: 'newProduct', n: 1 },
  { theme: 'declining', n: 1 },
  { theme: 'overstock', n: 1 },
];

const pickRecommendation = (groups, opts) => {
  const seasonFiltered = opts.seasonFilters?.length > 0
    ? groups.filter(g => opts.seasonFilters.includes(g.season))
    : groups;

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

const pickItems = (groups, theme, opts) => {
  if (theme === 'recommend') return pickRecommendation(groups, opts);

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
    { header: '7일 판매량', key: 'totalSales', width: 12 },
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
    row.alignment = { vertical: 'middle', wrapText: true };
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
  meta.columns = [{ key: 'k', width: 18 }, { key: 'v', width: 50 }];
  for (const [k, v] of info) meta.addRow({ k, v });
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
  if (opts.useDiversity) info.push(['카테고리 다양성', `한 카테고리 최대 ${opts.maxPerCategory}개`]);
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
  if (opts.useDiversity) info.push(['카테고리 다양성', `한 카테고리 최대 ${opts.maxPerCategory}개`]);
  buildMetaSheet(wb, info);

  await saveWorkbook(wb, `ADpicker_전체_${mode.toUpperCase()}`);
};

const App = () => {
  const [skus, setSkus] = useState([]);
  const [groups, setGroups] = useState([]);
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
    maxSales: 5,
    seasonFilters: [],
    useDiversity: true,
    maxPerCategory: 3,
  });
  const [embedImages, setEmbedImages] = useState(true);
  const [exporting, setExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState(null);
  const [showSelector, setShowSelector] = useState(false);
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
      const decoder = new TextDecoder('euc-kr');
      let text = decoder.decode(buf);
      if (!text.includes('상품명') && !text.includes('<table')) {
        text = new TextDecoder('utf-8').decode(buf);
      }
      const isHtml = text.trimStart().startsWith('<');
      const newMode = isHtml ? 'xls' : 'csv';
      if (isHtml) {
        const { skus: parsedSkus, dateLabels: dl } = parseHtmlXls(text);
        setSkus(parsedSkus);
        setGroups(groupByProduct(parsedSkus));
        setDateLabels(dl);
      } else {
        const productGroups = parseProductCsv(text);
        setSkus([]);
        setGroups(productGroups);
        setDateLabels([]);
      }
      if (newMode !== mode) setTheme('bestseller');
      setMode(newMode);
      setFileName(file.name);
    } catch (e) {
      setError(`파일 파싱 실패: ${e.message}`);
    } finally {
      setParsing(false);
    }
  }, []);

  const preview = useMemo(() => {
    if (groups.length === 0) return [];
    return pickItems(groups, theme, opts);
  }, [groups, theme, opts]);

  const handleExportSingle = useCallback(async () => {
    if (preview.length === 0) return;
    setExporting('single');
    setExportProgress(null);
    try {
      await exportSingleTheme(preview, theme, opts, embedImages, (p) => setExportProgress(p));
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
      await exportAllThemes(groups, mode, opts, embedImages, brands, categories, selectedIds, (p) => setExportProgress(p));
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
            <h1 className="font-serif text-4xl font-medium flex items-baseline gap-3 tracking-tight">
              ADpicker
              <span className="font-sans text-xs font-normal text-stone-500 tracking-wide uppercase">v1</span>
            </h1>
            <p className="text-sm text-stone-600 mt-2 font-light">인스타 메타광고 아이템 선정기</p>
          </div>
          {fileName && (
            <div className="flex items-center gap-3 flex-wrap">
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

        {groups.length === 0 ? (
          <UploadArea
            onFile={handleFile}
            parsing={parsing}
            inputRef={fileInputRef}
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
                <ThemeOptions
                  theme={theme}
                  opts={opts}
                  setOpts={setOpts}
                  categories={categories}
                  brands={brands}
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
              <Preview items={preview} theme={theme} dateLabels={dateLabels} />
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

const UploadArea = ({ onFile, parsing, inputRef }) => {
  const [drag, setDrag] = useState(false);
  return (
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
      <h2 className="font-serif text-3xl font-medium mb-3 tracking-tight">판매 데이터 파일을 올려주세요</h2>
      <p className="text-sm text-stone-600 mb-8 font-light leading-relaxed">
        <span className="font-medium text-stone-800">stk_forOptSalesInfo .xls</span> (SKU 단위 · 이미지 포함) 또는{' '}
        <span className="font-medium text-stone-800">sts_prdListStatistics .csv</span> (상품 누적) — 둘 다 자동 인식해요.
        <br />파일을 끌어다 놓거나 아래 버튼으로 선택하세요. 데이터는 브라우저에서만 처리됩니다.
      </p>
      <input
        ref={inputRef}
        type="file"
        accept=".xls,.html,.htm,.csv"
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
  );
};

const Panel = ({ title, icon: Icon, children }) => (
  <div className="bg-cream-50 border border-cream-400">
    <div className="px-4 py-3 border-b border-cream-400 flex items-center gap-2">
      {Icon && <Icon size={14} strokeWidth={1.5} className="text-stone-600" />}
      <h3 className="font-serif text-sm font-medium text-stone-800 tracking-tight">{title}</h3>
    </div>
    <div className="p-4">{children}</div>
  </div>
);

const ThemeOptions = ({ theme, opts, setOpts, categories, brands }) => {
  const set = (k, v) => setOpts({ ...opts, [k]: v });

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
          점수 = 일평균수량 × log(1 + 햇수). 기간이 긴 .xls는 7일 합계를 일수로 나눠 계산하고,
          누적 .csv는 파일의 일평균수량을 그대로 써요.
        </p>
      </div>
    );
  }
  if (theme === 'rising') {
    return (
      <div>
        <label className="text-xs text-stone-600 block mb-1">최소 7일 판매량</label>
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
  if (theme === 'declining') {
    return (
      <div className="space-y-3">
        <div>
          <label className="text-xs text-stone-600 block mb-1">전반 4일 최소 판매량</label>
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
          원래 잘 팔리던 상품 중 후반 4일에 판매가 떨어진 것을 우선. 점수 = 감소율 × log(1+전반판매)
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
          <label className="text-xs text-stone-600 block mb-1">최대 7일 판매량</label>
          <input
            type="number"
            min="0"
            value={opts.maxSales}
            onChange={e => set('maxSales', parseInt(e.target.value) || 5)}
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

const Preview = ({ items, theme, dateLabels }) => {
  if (items.length === 0) {
    return (
      <div className="bg-cream-50 border border-cream-400 p-16 text-center">
        <LucideAlertCircle strokeWidth={1.2} className="mx-auto text-stone-400 mb-4" size={32} />
        <p className="font-serif text-xl text-stone-800">선정할 상품이 없어요</p>
        <p className="text-sm text-stone-500 mt-2 font-light">필터/옵션을 조정해 보세요.</p>
      </div>
    );
  }
  return (
    <div className="bg-cream-50 border border-cream-400">
      <div className="px-6 py-5 border-b border-cream-400 flex items-center justify-between">
        <div>
          <h2 className="font-serif text-2xl font-medium tracking-tight">미리보기</h2>
          <p className="text-xs text-stone-600 mt-1 font-light">
            {items.length}개 상품 · {dateLabels.length > 0
              ? `데이터 기간: ${dateLabels[0]} ~ ${dateLabels[dateLabels.length - 1]}`
              : '상품별 누적 데이터 (옵션·이미지 정보 없음)'}
          </p>
        </div>
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
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">7일 판매</th>
              <th className="px-3 py-2 text-left font-medium whitespace-nowrap">베스트 SKU 1위</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">1위 판매</th>
              <th className="px-3 py-2 text-left font-medium whitespace-nowrap">베스트 SKU 2위</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">2위 판매</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">현재 재고</th>
              <th className="px-3 py-2 text-right font-medium whitespace-nowrap">매출액</th>
              {theme === 'recommend' && (
                <th className="px-3 py-2 text-left font-medium whitespace-nowrap">선정 사유</th>
              )}
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
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default App;
