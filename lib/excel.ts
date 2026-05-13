import * as XLSX from "xlsx";

export interface LabelData {
  box: string;
  date: string;
  location: string;
  productName: string;
  optionName: string;
  barcode: string;
  quantity: number;
  quantityAdjusted?: boolean;
}

/** 연결 테스트·프린터 시험용 (엑셀 없이 1장 출력) */
export const SAMPLE_LABEL_DATA: LabelData = {
  box: "TEST",
  date: new Date().toISOString().slice(0, 10),
  location: "LOC-01",
  productName: "연결 테스트 상품",
  optionName: "샘플 옵션",
  barcode: "TEST12345678",
  quantity: 1,
};

/** 출력 단위 53 등에서 맨 끝에 붙이는 빈 라벨(바코드 없음) */
export const BLANK_LABEL_DATA: LabelData = {
  box: "",
  date: "",
  location: "",
  productName: "",
  optionName: "",
  barcode: "",
  quantity: 1,
};

function cellString(v: unknown): string {
  if (v == null || v === "") return "";
  if (typeof v === "number" && Number.isFinite(v)) return String(v);
  return String(v).trim();
}

/** Collapse duplicate trimmed keys (last wins). */
function normalizeRowKeys(row: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(row)) {
    const key = k.trim();
    if (!key) continue;
    out[key] = v;
  }
  return out;
}

function pickFirst(row: Record<string, unknown>, keys: string[]): string {
  for (const key of keys) {
    const v = row[key];
    const s = cellString(v);
    if (s) return s;
  }
  for (const [k, v] of Object.entries(row)) {
    const kl = k.toLowerCase();
    for (const key of keys) {
      if (kl === key.toLowerCase()) {
        const s = cellString(v);
        if (s) return s;
      }
    }
  }
  return "";
}

const BARCODE_HEADER_HINTS = /바코드|barcode|sku|품번|item\s*code|^code$/i;
const QUANTITY_HEADERS = ["갯수", "개수", "수량", "Qty", "QTY", "quantity", "Quantity", "매수"];
const MAX_QUANTITY = 999;

function findBarcode(row: Record<string, unknown>): string {
  const direct = pickFirst(row, [
    "바코드",
    "Barcode",
    "BARCODE",
    "바코드번호",
    "코드",
    "SKU",
    "품번",
    "ItemCode",
    "item_code",
    "상품코드",
  ]);
  if (direct) return direct;
  for (const [k, v] of Object.entries(row)) {
    if (BARCODE_HEADER_HINTS.test(k)) {
      const s = cellString(v);
      if (s) return s;
    }
  }
  return "";
}

function pickQuantityRaw(row: Record<string, unknown>): string {
  // 수량은 명시된 수량 컬럼에서만 읽는다. (재고 등 다른 숫자 컬럼 무시)
  for (const header of QUANTITY_HEADERS) {
    const exact = cellString(row[header]);
    if (exact) return exact;
  }
  const lowerToValue = new Map<string, string>();
  for (const [k, v] of Object.entries(row)) {
    const normalizedKey = k.trim().toLowerCase();
    if (!normalizedKey) continue;
    const value = cellString(v);
    if (value) lowerToValue.set(normalizedKey, value);
  }
  for (const header of QUANTITY_HEADERS) {
    const value = lowerToValue.get(header.toLowerCase());
    if (value) return value;
  }
  return "";
}

function sanitizeQuantity(raw: string): { quantity: number; adjusted: boolean } {
  const n = Number(raw);
  if (!Number.isFinite(n) || n <= 0) return { quantity: 1, adjusted: raw !== "" };

  const whole = Math.floor(n);
  // 날짜 직렬 값이 수량으로 잘못 매핑되는 경우 보호.
  if (whole >= 40000 && whole <= 60000) {
    return { quantity: 1, adjusted: true };
  }
  if (whole > MAX_QUANTITY) {
    return { quantity: MAX_QUANTITY, adjusted: true };
  }
  return { quantity: whole, adjusted: whole !== n };
}

function findQuantity(row: Record<string, unknown>): { quantity: number; adjusted: boolean } {
  const raw = pickQuantityRaw(row);
  if (!raw) return { quantity: 1, adjusted: false };
  return sanitizeQuantity(raw);
}

function extractLabelsFromObjects(rows: Record<string, unknown>[]): LabelData[] {
  const labels: LabelData[] = [];
  for (const raw of rows) {
    const row = normalizeRowKeys(raw);
    const barcode = findBarcode(row);
    if (!barcode) continue;

    const quantityResult = findQuantity(row);
    labels.push({
      box: pickFirst(row, ["box", "Box", "BOX", "박스", "BoxNo"]),
      date: pickFirst(row, ["date", "Date", "DATE", "일자", "날짜", "입고일"]),
      location: pickFirst(row, ["location", "Location", "로케이션", "위치", "창고"]),
      productName: pickFirst(row, [
        "등록상품명",
        "입고상품명",
        "상품명",
        "제품명",
        "품명",
        "product",
        "ProductName",
      ]),
      optionName: pickFirst(row, ["옵션명", "옵션", "Option", "option", "규격"]),
      barcode,
      quantity: quantityResult.quantity,
      quantityAdjusted: quantityResult.adjusted,
    });
  }
  return labels;
}

function sheetRows(
  sheet: XLSX.WorkSheet,
  opts?: { headerStartRow0?: number }
): Record<string, unknown>[] {
  if (opts?.headerStartRow0 != null) {
    return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
      range: opts.headerStartRow0,
      defval: "",
    });
  }
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
}

/** When default JSON keys are wrong (e.g. __EMPTY), find a header row and build objects. */
function extractLabelsViaHeaderScan(sheet: XLSX.WorkSheet): LabelData[] {
  const aoa = XLSX.utils.sheet_to_json<(string | number | boolean | null | undefined)[]>(
    sheet,
    { header: 1, defval: "", raw: false }
  );
  const maxHeaderScan = Math.min(40, aoa.length);
  for (let r = 0; r < maxHeaderScan; r++) {
    const headerCells = aoa[r];
    if (!Array.isArray(headerCells)) continue;
    const hasBarcodeHeader = headerCells.some((cell) => {
      const s = cellString(cell);
      if (!s || s.length > 40) return false;
      return (
        s === "바코드" ||
        /^barcode$/i.test(s) ||
        /^바코드번호$/i.test(s) ||
        /^sku$/i.test(s) ||
        /^품번$/i.test(s) ||
        /^상품코드$/i.test(s) ||
        BARCODE_HEADER_HINTS.test(s)
      );
    });
    if (!hasBarcodeHeader) continue;

    const headers = headerCells.map((c) => cellString(c));
    const objects: Record<string, unknown>[] = [];
    for (let dataR = r + 1; dataR < aoa.length; dataR++) {
      const dataRow = aoa[dataR];
      if (!Array.isArray(dataRow)) continue;
      const obj: Record<string, unknown> = {};
      let any = false;
      for (let c = 0; c < headers.length; c++) {
        const h = headers[c];
        if (!h) continue;
        const v = dataRow[c];
        if (v != null && v !== "") any = true;
        obj[h] = v;
      }
      if (any) objects.push(obj);
    }
    const labels = extractLabelsFromObjects(objects);
    if (labels.length > 0) return labels;
  }
  return [];
}

function parseSheetToLabels(sheet: XLSX.WorkSheet): LabelData[] {
  const starts = [undefined, 1, 2] as const;
  for (const start of starts) {
    const rows =
      start === undefined
        ? sheetRows(sheet)
        : sheetRows(sheet, { headerStartRow0: start });
    const labels = extractLabelsFromObjects(rows);
    if (labels.length > 0) return labels;
  }
  return extractLabelsViaHeaderScan(sheet);
}

function readWorkbook(file: File): Promise<XLSX.WorkBook> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        resolve(XLSX.read(data, { type: "array" }));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export async function listExcelSheets(file: File): Promise<string[]> {
  const workbook = await readWorkbook(file);
  return workbook.SheetNames;
}

export async function parseExcel(
  file: File,
  opts?: { sheetName?: string }
): Promise<LabelData[]> {
  const workbook = await readWorkbook(file);
  const defaultSheetName = workbook.SheetNames[0];
  const targetSheetName = opts?.sheetName && workbook.SheetNames.includes(opts.sheetName)
    ? opts.sheetName
    : defaultSheetName;
  const sheet = workbook.Sheets[targetSheetName];
  return parseSheetToLabels(sheet);
}
