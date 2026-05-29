"use client";
import { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { parseExcel, listExcelSheets, LabelData, SAMPLE_LABEL_DATA, BLANK_LABEL_DATA } from "@/lib/excel";
import Label from "@/components/Label";

type NavTab = "print" | "jobs" | "test";

type LabelDataWithSpan = LabelData & {
  rowIndex: number;
  startLabel: number;
  endLabel: number;
};

interface JobRecord {
  id: string;
  createdAt: string;
  fileName: string;
  sheetName: string;
  totalLabels: number;
  rowCount: number;
  adjustedQuantityCount: number;
  importHint: string | null;
  /** 이전 버전 작업에는 없을 수 있음 */
  labelsSnapshot?: { data: LabelData; key: number }[];
}

const JOBS_STORAGE_KEY = "label-printer-jobs";
/** 작업 이력 보관 기간(용량·할당량 완화) */
const JOB_RETENTION_MS = 7 * 24 * 60 * 60 * 1000;

function pruneJobsByAge(jobs: JobRecord[], nowMs: number = Date.now()): JobRecord[] {
  const cutoff = nowMs - JOB_RETENTION_MS;
  return jobs.filter((j) => {
    const t = new Date(j.createdAt).getTime();
    return Number.isFinite(t) && t >= cutoff;
  });
}

const BATCH_OPTIONS = [53, 20, 10, 1] as const;
const MAX_BATCH_SIZE_WITHOUT_ROWS = 999;

function clampBatchSize(size: number, rowCount: number): number {
  const floored = Math.floor(size);
  const min = Math.max(1, floored);
  const max = rowCount > 0 ? rowCount : MAX_BATCH_SIZE_WITHOUT_ROWS;
  return Math.min(min, max);
}

function toLocalDateKey(iso: string): string {
  const d = new Date(iso);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/** 연속 클릭·이중 이벤트로 인쇄 대화상자가 두 번 뜨는 것 방지 */
let lastPrintStartedAt = 0;

function printWithTarget(target: "all" | "first" | "test" | "batch") {
  const now = Date.now();
  if (now - lastPrintStartedAt < 900) return;
  lastPrintStartedAt = now;

  const html = document.documentElement;
  html.dataset.printTarget = target;
  const cleanup = () => {
    delete html.dataset.printTarget;
  };
  window.addEventListener("afterprint", cleanup, { once: true });
  setTimeout(cleanup, 4000);
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      window.print();
    });
  });
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<NavTab>("print");
  const [labels, setLabels] = useState<{ data: LabelData; key: number }[]>([]);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState("");
  const [importHint, setImportHint] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [rows, setRows] = useState<LabelDataWithSpan[]>([]);
  const [jobHistory, setJobHistory] = useState<JobRecord[]>([]);
  const [selectedJobId, setSelectedJobId] = useState<string | null>(null);
  const [batchSize, setBatchSize] = useState<number>(53);
  const [batchSizeDraft, setBatchSizeDraft] = useState("53");
  const [batchSizeHint, setBatchSizeHint] = useState<string | null>(null);
  const [selectedBatchIndex, setSelectedBatchIndex] = useState(0);
  const [printedBatchIndexes, setPrintedBatchIndexes] = useState<number[]>([]);
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const [isVerifyListOpen, setIsVerifyListOpen] = useState(false);
  const [verifyListOverride, setVerifyListOverride] = useState<{ data: LabelData; key: number }[] | null>(null);
  const [jobDateFilter, setJobDateFilter] = useState("");
  const [printPortalsReady, setPrintPortalsReady] = useState(false);

  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(JOBS_STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        setJobHistory([]);
        return;
      }
      const jobs = parsed as JobRecord[];
      const pruned = pruneJobsByAge(jobs);
      if (pruned.length !== jobs.length) {
        window.localStorage.setItem(JOBS_STORAGE_KEY, JSON.stringify(pruned));
      }
      setJobHistory(pruned);
    } catch {
      setJobHistory([]);
    }
  }, []);

  useEffect(() => {
    setPrintPortalsReady(true);
  }, []);

  const selectedJob = useMemo(
    () => jobHistory.find((job) => job.id === selectedJobId) ?? null,
    [jobHistory, selectedJobId]
  );

  const filteredJobs = useMemo(() => {
    if (!jobDateFilter) return jobHistory;
    return jobHistory.filter((job) => toLocalDateKey(job.createdAt) === jobDateFilter);
  }, [jobHistory, jobDateFilter]);

  useEffect(() => {
    if (!jobDateFilter || !selectedJobId) return;
    const sel = jobHistory.find((j) => j.id === selectedJobId);
    if (!sel || toLocalDateKey(sel.createdAt) !== jobDateFilter) {
      setSelectedJobId(null);
    }
  }, [jobDateFilter, jobHistory, selectedJobId]);

  const rowBatchCards = useMemo(() => {
    const cards: Array<{
      index: number;
      rowStart: number;
      rowEnd: number;
      rows: LabelDataWithSpan[];
      startLabel: number;
      endLabel: number;
      totalLabels: number;
    }> = [];
    for (let start = 0, idx = 0; start < rows.length; start += batchSize, idx += 1) {
      const rowsInCard = rows.slice(start, start + batchSize);
      if (rowsInCard.length === 0) continue;
      const startLabel = rowsInCard[0].startLabel;
      const endLabel = rowsInCard[rowsInCard.length - 1].endLabel;
      const totalLabels = rowsInCard.reduce((sum, r) => sum + r.quantity, 0);
      cards.push({
        index: idx,
        rowStart: start + 1,
        rowEnd: start + rowsInCard.length,
        rows: rowsInCard,
        startLabel,
        endLabel,
        totalLabels,
      });
    }
    return cards;
  }, [rows, batchSize]);

  const selectedBatch = useMemo(
    () => rowBatchCards.find((card) => card.index === selectedBatchIndex) ?? null,
    [rowBatchCards, selectedBatchIndex]
  );

  const printedLabelCount = useMemo(
    () =>
      rowBatchCards
        .filter((card) => printedBatchIndexes.includes(card.index))
        .reduce((sum, card) => sum + card.totalLabels, 0),
    [rowBatchCards, printedBatchIndexes]
  );

  const verifyDisplayRows = useMemo(
    () => verifyListOverride ?? labels,
    [verifyListOverride, labels]
  );

  function saveJob(record: JobRecord) {
    setJobHistory((prev) => {
      const next = pruneJobsByAge([record, ...prev]).slice(0, 100);
      window.localStorage.setItem(JOBS_STORAGE_KEY, JSON.stringify(next));
      return next;
    });
    setSelectedJobId(record.id);
  }

  async function parseAndExpand(file: File, sheetName: string) {
    const parsedRows = await parseExcel(file, { sheetName });
    const adjustedQuantityCount = parsedRows.filter((row) => row.quantityAdjusted).length;
    const expanded: { data: LabelData; key: number }[] = [];
    const rowsWithSpan: LabelDataWithSpan[] = [];
    let key = 0;
    let cursor = 0;
    parsedRows.forEach((row, rowIndex) => {
      const start = cursor;
      const quantity = row.quantity ?? 1;
      for (let i = 0; i < quantity; i += 1) {
        expanded.push({ data: row, key: key++ });
        cursor += 1;
      }
      const end = cursor;
      rowsWithSpan.push({ ...row, rowIndex, startLabel: start, endLabel: end });
    });
    setLabels(expanded);
    setRows(rowsWithSpan);

    let hint: string | null = null;
    if (expanded.length === 0) {
      hint =
        "파일은 읽혔지만 바코드가 있는 데이터 행이 없습니다. 첫 줄에 열 이름(바코드, 수량, 등록상품명 등)이 오도록 하거나, 시트 선택이 올바른지 확인해 주세요.";
    } else if (adjustedQuantityCount > 0) {
      hint = `일부 행의 수량 값이 비정상적으로 보여 ${adjustedQuantityCount}개 행을 자동 보정했습니다. 수량 컬럼(수량/개수/Qty) 값을 확인해 주세요.`;
    }
    setImportHint(hint);
    setBatchSizeHint(null);
    let nextBatch = 53;
    setBatchSize((prev) => {
      nextBatch = clampBatchSize(prev, rowsWithSpan.length);
      return nextBatch;
    });
    setBatchSizeDraft(String(nextBatch));
    setSelectedBatchIndex(0);
    setPrintedBatchIndexes([]);

    saveJob({
      id: crypto.randomUUID(),
      createdAt: new Date().toISOString(),
      fileName: file.name,
      sheetName,
      totalLabels: expanded.length,
      rowCount: parsedRows.length,
      adjustedQuantityCount,
      labelsSnapshot: expanded,
      importHint:
        expanded.length === 0
          ? "파일은 읽혔지만 바코드가 있는 데이터 행이 없습니다."
          : adjustedQuantityCount > 0
            ? `${adjustedQuantityCount}개 행 수량 자동 보정`
            : null,
    });
  }

  async function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const nextFile = e.target.files?.[0];
    if (!nextFile) return;
    setLoading(true);
    setFileName(nextFile.name);
    setImportHint(null);
    try {
      const names = await listExcelSheets(nextFile);
      const firstSheet = names[0] ?? "";
      setUploadedFile(nextFile);
      setSheetNames(names);
      setSelectedSheet(firstSheet);
      setActiveTab("print");
      if (firstSheet) {
        await parseAndExpand(nextFile, firstSheet);
      } else {
        setLabels([]);
        setImportHint("엑셀 시트를 찾지 못했습니다. 파일 형식을 확인해 주세요.");
      }
    } catch (err) {
      setImportHint(null);
      alert("엑셀 파일을 읽는 중 오류가 발생했습니다.");
      console.error(err);
    } finally {
      setLoading(false);
    }
  }

  async function handleSheetChange(e: React.ChangeEvent<HTMLSelectElement>) {
    const nextSheet = e.target.value;
    setSelectedSheet(nextSheet);
    if (!uploadedFile || !nextSheet) return;
    setLoading(true);
    setImportHint(null);
    try {
      await parseAndExpand(uploadedFile, nextSheet);
    } catch (err) {
      setImportHint(null);
      alert("선택한 시트를 읽는 중 오류가 발생했습니다.");
      console.error(err);
    } finally {
      setLoading(false);
    }
  }

  function resetBatchProgress() {
    setSelectedBatchIndex(0);
    setPrintedBatchIndexes([]);
  }

  /** 프리셋 버튼 — 입력값만 변경 (적용 전) */
  function pickPresetBatchSize(size: number) {
    setBatchSizeDraft(String(size));
    setBatchSizeHint(null);
  }

  /** 적용 — 프리셋·직접 입력 공통 */
  function applyBatchSize() {
    const raw = batchSizeDraft.trim();
    if (!raw) {
      setBatchSizeHint("1 이상의 숫자를 입력해 주세요.");
      return;
    }
    const parsed = Number(raw);
    if (!Number.isFinite(parsed) || parsed < 1) {
      setBatchSizeHint("1 이상의 숫자를 입력해 주세요.");
      return;
    }
    const clamped = clampBatchSize(parsed, rows.length);
    setBatchSize(clamped);
    setBatchSizeDraft(String(clamped));
    setBatchSizeHint(null);
    resetBatchProgress();
  }

  const draftBatchNumber = Number(batchSizeDraft.trim());
  const hasPendingBatchSize =
    batchSizeDraft.trim() !== "" &&
    Number.isFinite(draftBatchNumber) &&
    draftBatchNumber >= 1 &&
    clampBatchSize(draftBatchNumber, rows.length) !== batchSize;

  function handlePrintBatch() {
    if (!selectedBatch || selectedBatch.totalLabels === 0) return;
    printWithTarget("batch");
    setPrintedBatchIndexes((prev) => (prev.includes(selectedBatch.index) ? prev : [...prev, selectedBatch.index]));
  }

  function handleOpenBatchPreview() {
    if (!selectedBatch || selectedBatch.totalLabels === 0) return;
    setIsPreviewOpen(true);
  }

  return (
    <>
    <main className="min-h-screen bg-gray-100 p-8">
      <div className="mx-auto flex w-full max-w-6xl gap-6">
        <aside className="w-full max-w-[240px] rounded-xl border border-gray-200 bg-white shadow-sm print:hidden">
          <div className="border-b border-gray-200 px-4 py-4">
            <p className="text-sm font-bold text-gray-900">미즈코스 관리자</p>
          </div>
          <div className="px-3 py-3">
            <p className="px-2 pb-2 text-xs font-semibold uppercase tracking-wide text-gray-400">작업 메뉴</p>
            <nav className="flex flex-col gap-1">
              {[
                { id: "print", label: "바코드 출력하기" },
                { id: "jobs", label: "작업 조회" },
                { id: "test", label: "테스트" },
              ].map((tab) => (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() => setActiveTab(tab.id as NavTab)}
                  className={`rounded-md px-3 py-2 text-left text-sm font-medium transition ${
                    activeTab === tab.id
                      ? "bg-blue-50 text-blue-700"
                      : "text-gray-700 hover:bg-gray-100"
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </nav>
          </div>
        </aside>

        <div className="min-w-0 flex-1">
          <div className="mb-8">
            <h1 className="text-2xl font-bold text-gray-800">라벨 출력 시스템</h1>
            <p className="mt-1 text-sm text-gray-500">엑셀 파일을 업로드하면 입고 바코드 라벨을 출력합니다</p>
          </div>

        {activeTab === "print" && (
          <>
            <label className="relative mb-6 block cursor-pointer rounded-xl border-2 border-dashed border-blue-300 bg-white p-10 text-center transition hover:border-blue-500">
              <div className="pointer-events-none">
                <p className="text-base font-semibold text-gray-700">엑셀 파일 업로드</p>
                <p className="mt-1 font-medium text-gray-600">클릭하여 파일 선택</p>
                <p className="mt-1 text-sm text-gray-400">.xlsx / .xls 파일 지원</p>
                {fileName && (
                  <p className="mt-3 text-sm font-medium text-blue-600">선택됨: {fileName}</p>
                )}
              </div>
              <input
                type="file"
                accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
                className="absolute inset-0 z-10 h-full w-full cursor-pointer opacity-0"
                onClick={(e) => {
                  e.currentTarget.value = "";
                }}
                onChange={handleFile}
              />
            </label>

            {sheetNames.length > 1 && (
              <div className="mb-6 rounded-xl border border-gray-200 bg-white p-4 shadow-sm print:hidden">
                <label htmlFor="sheet-select" className="mb-2 block text-sm font-semibold text-gray-700">
                  시트 선택
                </label>
                <div className="flex flex-wrap items-center gap-3">
                  <select
                    id="sheet-select"
                    value={selectedSheet}
                    onChange={handleSheetChange}
                    className="min-w-[220px] rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm"
                  >
                    {sheetNames.map((name) => (
                      <option key={name} value={name}>
                        {name}
                      </option>
                    ))}
                  </select>
                  <span className="text-xs text-gray-500">현재 시트: {selectedSheet}</span>
                </div>
              </div>
            )}

            {importHint && !loading && (
              <div
                className="mb-6 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900 print:hidden"
                role="status"
              >
                {importHint}
              </div>
            )}

            {labels.length > 0 && (
              <div className="mb-6 rounded-xl bg-white p-4 shadow-sm print:hidden">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <span className="font-medium text-gray-700">총 라벨 수: </span>
                    <span className="text-lg font-bold text-blue-600">{labels.length}장</span>
                    <span className="ml-3 text-sm text-gray-500">
                      출력됨: {printedLabelCount}장 / 남음: {Math.max(labels.length - printedLabelCount, 0)}장
                    </span>
                  </div>
                  <span className="text-sm text-gray-500">
                    선택 그룹:{" "}
                    {selectedBatch
                      ? `${selectedBatch.rowStart}~${selectedBatch.rowEnd}행 (${selectedBatch.totalLabels}장)`
                      : "-"}
                  </span>
                </div>

                <div className="mt-4 flex flex-col gap-2">
                  <div className="flex flex-wrap items-center gap-2 text-sm text-gray-500">
                    <span className="font-medium text-gray-700">출력 단위</span>
                    <span>
                      적용됨: <strong className="text-gray-800">{batchSize}행씩</strong> · {rowBatchCards.length}그룹
                    </span>
                    {hasPendingBatchSize && (
                      <span className="text-amber-700">→ 적용 대기: {batchSizeDraft}행씩</span>
                    )}
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    {BATCH_OPTIONS.map((size) => (
                      <button
                        key={size}
                        type="button"
                        onClick={() => pickPresetBatchSize(size)}
                        className={`rounded-md px-3 py-1.5 text-sm font-semibold ${
                          batchSizeDraft === String(size)
                            ? "bg-blue-600 text-white"
                            : "border border-gray-300 bg-white text-gray-700 hover:bg-gray-100"
                        }`}
                      >
                        {size}개
                      </button>
                    ))}
                    <span className="mx-1 text-sm text-gray-400">|</span>
                    <input
                      type="number"
                      min={1}
                      max={rows.length > 0 ? rows.length : MAX_BATCH_SIZE_WITHOUT_ROWS}
                      inputMode="numeric"
                      value={batchSizeDraft}
                      onChange={(e) => {
                        setBatchSizeDraft(e.target.value);
                        setBatchSizeHint(null);
                      }}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          e.preventDefault();
                          applyBatchSize();
                        }
                      }}
                      placeholder={rows.length > 0 ? `1~${rows.length}` : "행 수"}
                      aria-label="출력 단위 행 수"
                      className="w-24 rounded-md border border-gray-300 px-2 py-1.5 text-sm"
                    />
                    <span className="text-sm text-gray-600">행씩</span>
                    <button
                      type="button"
                      onClick={applyBatchSize}
                      className={`rounded-md px-4 py-1.5 text-sm font-bold ${
                        hasPendingBatchSize
                          ? "bg-blue-600 text-white hover:bg-blue-700"
                          : "border border-gray-300 bg-white text-gray-700 hover:bg-gray-100"
                      }`}
                    >
                      적용
                    </button>
                    {batchSizeHint && (
                      <span className="text-sm text-amber-700">{batchSizeHint}</span>
                    )}
                  </div>
                </div>

                <div className="mt-4 flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={handlePrintBatch}
                    disabled={!selectedBatch || selectedBatch.totalLabels === 0}
                    className="rounded-lg bg-blue-600 px-5 py-2 text-sm font-bold text-white hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-gray-300"
                  >
                    선택 그룹 출력
                  </button>
                  <button
                    type="button"
                    onClick={handleOpenBatchPreview}
                    disabled={!selectedBatch || selectedBatch.totalLabels === 0}
                    className="rounded-lg border border-gray-300 bg-white px-5 py-2 text-sm font-bold text-gray-700 hover:bg-gray-100 disabled:cursor-not-allowed disabled:bg-gray-100 disabled:text-gray-400"
                  >
                    선택 그룹 보기
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setVerifyListOverride(null);
                      setIsVerifyListOpen(true);
                    }}
                    className="rounded-lg border border-gray-300 bg-white px-5 py-2 text-sm font-bold text-gray-700 hover:bg-gray-100"
                  >
                    검증하기
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedBatchIndex(0);
                      setPrintedBatchIndexes([]);
                    }}
                    className="rounded-lg border border-gray-300 bg-white px-5 py-2 text-sm font-bold text-gray-700 hover:bg-gray-100"
                  >
                    출력 상태 초기화
                  </button>
                </div>

                <div className="mt-4 grid gap-3 sm:grid-cols-2 xl:grid-cols-3">
                  {rowBatchCards.map((card) => {
                    const printed = printedBatchIndexes.includes(card.index);
                    const selected = selectedBatchIndex === card.index;
                    return (
                      <button
                        key={`batch-card-${card.index}`}
                        type="button"
                        onClick={() => setSelectedBatchIndex(card.index)}
                        className={`rounded-lg border p-3 text-left transition ${
                          printed
                            ? "border-gray-200 bg-gray-100 text-gray-400"
                            : selected
                              ? "border-blue-500 bg-blue-50 text-blue-900"
                              : "border-gray-200 bg-white text-gray-700 hover:border-blue-300 hover:bg-blue-50/40"
                        }`}
                      >
                        <p className="text-xs font-semibold">그룹 {card.index + 1}</p>
                        <p className="mt-1 text-sm font-bold">
                          {card.rowStart}~{card.rowEnd}행 ({card.totalLabels}장)
                        </p>
                        <p className="mt-2 truncate text-xs">
                          {card.rows[0]?.productName ?? "-"}
                        </p>
                        <p className="truncate text-xs">{card.rows[0]?.optionName ?? "-"}</p>
                        {printed && <p className="mt-2 text-[11px] font-semibold">출력 완료</p>}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}
          </>
        )}

        {activeTab === "jobs" && (
          <section className="rounded-xl border border-gray-200 bg-white p-5 shadow-sm print:hidden">
            <h2 className="text-lg font-semibold text-gray-800">작업 조회</h2>
            <p className="mt-1 text-sm text-gray-500">이전에 처리한 엑셀 작업 목록과 상세 정보를 확인합니다.</p>
            <p className="mt-2 text-xs text-gray-400">작업 이력은 최근 7일치만 이 기기에 보관됩니다.</p>
            {jobHistory.length === 0 ? (
              <p className="mt-4 rounded-lg bg-gray-50 p-4 text-sm text-gray-500">저장된 작업 이력이 없습니다.</p>
            ) : (
              <div className="mt-4 grid gap-4 md:grid-cols-[1.2fr_1fr]">
                <div className="flex min-h-0 flex-col gap-2">
                  <div className="flex flex-wrap items-center gap-2 rounded-lg border border-gray-200 bg-gray-50 px-3 py-2">
                    <label htmlFor="job-date-filter" className="text-sm font-medium text-gray-700">
                      날짜 필터
                    </label>
                    <input
                      id="job-date-filter"
                      type="date"
                      value={jobDateFilter}
                      onChange={(e) => setJobDateFilter(e.target.value)}
                      className="rounded-md border border-gray-300 bg-white px-2 py-1.5 text-sm"
                    />
                    <button
                      type="button"
                      onClick={() => setJobDateFilter("")}
                      className="rounded-md border border-gray-300 bg-white px-2 py-1.5 text-xs font-semibold text-gray-700 hover:bg-gray-100"
                    >
                      전체
                    </button>
                    <span className="text-xs text-gray-500">
                      {jobDateFilter ? `${filteredJobs.length}건` : `전체 ${jobHistory.length}건`}
                    </span>
                  </div>
                  <div className="max-h-[420px] overflow-auto rounded-lg border border-gray-200">
                    {filteredJobs.length === 0 ? (
                      <p className="p-4 text-sm text-gray-500">해당 날짜에 저장된 작업이 없습니다.</p>
                    ) : (
                      filteredJobs.map((job) => (
                        <button
                          key={job.id}
                          type="button"
                          onClick={() => setSelectedJobId(job.id)}
                          className={`w-full border-b border-gray-100 px-4 py-3 text-left last:border-b-0 ${
                            selectedJobId === job.id ? "bg-blue-50" : "hover:bg-gray-50"
                          }`}
                        >
                          <p className="text-sm font-semibold text-gray-800">{job.fileName}</p>
                          <p className="mt-1 text-xs text-gray-500">
                            {new Date(job.createdAt).toLocaleString()} · 시트: {job.sheetName} · 라벨:{" "}
                            {job.totalLabels}장
                          </p>
                        </button>
                      ))
                    )}
                  </div>
                </div>
                <div className="rounded-lg border border-gray-200 bg-gray-50 p-4">
                  {selectedJob ? (
                    <>
                      <h3 className="text-sm font-semibold text-gray-800">작업 상세</h3>
                      <dl className="mt-3 space-y-2 text-sm text-gray-700">
                        <div>
                          <dt className="font-medium text-gray-500">파일명</dt>
                          <dd>{selectedJob.fileName}</dd>
                        </div>
                        <div>
                          <dt className="font-medium text-gray-500">시트명</dt>
                          <dd>{selectedJob.sheetName}</dd>
                        </div>
                        <div>
                          <dt className="font-medium text-gray-500">처리 시각</dt>
                          <dd>{new Date(selectedJob.createdAt).toLocaleString()}</dd>
                        </div>
                        <div>
                          <dt className="font-medium text-gray-500">데이터 행 수</dt>
                          <dd>{selectedJob.rowCount}</dd>
                        </div>
                        <div>
                          <dt className="font-medium text-gray-500">총 라벨 수</dt>
                          <dd>{selectedJob.totalLabels}</dd>
                        </div>
                        <div>
                          <dt className="font-medium text-gray-500">수량 자동 보정 행</dt>
                          <dd>{selectedJob.adjustedQuantityCount}</dd>
                        </div>
                        {selectedJob.importHint && (
                          <div>
                            <dt className="font-medium text-gray-500">메모</dt>
                            <dd>{selectedJob.importHint}</dd>
                          </div>
                        )}
                      </dl>
                      <div className="mt-4 border-t border-gray-200 pt-4">
                        <button
                          type="button"
                          disabled={
                            !selectedJob.labelsSnapshot || selectedJob.labelsSnapshot.length === 0
                          }
                          title={
                            !selectedJob.labelsSnapshot?.length
                              ? "이 작업은 저장 시점에 목록이 없습니다. 동일 파일을 다시 업로드해 주세요."
                              : undefined
                          }
                          onClick={() => {
                            if (!selectedJob.labelsSnapshot?.length) return;
                            setVerifyListOverride(selectedJob.labelsSnapshot);
                            setIsVerifyListOpen(true);
                          }}
                          className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm font-bold text-gray-800 hover:bg-gray-100 disabled:cursor-not-allowed disabled:bg-gray-100 disabled:text-gray-400"
                        >
                          출력 목록 검증
                        </button>
                        {!selectedJob.labelsSnapshot?.length && (
                          <p className="mt-2 text-xs text-gray-500">
                            이 작업은 저장 시점에 목록이 없습니다. 동일 파일을 다시 업로드해 주세요.
                          </p>
                        )}
                      </div>
                    </>
                  ) : (
                    <p className="text-sm text-gray-500">왼쪽 목록에서 작업을 선택하세요.</p>
                  )}
                </div>
              </div>
            )}
          </section>
        )}

        {activeTab === "test" && (
          <section className="mb-6 rounded-xl border border-gray-200 bg-white p-5 shadow-sm print:hidden">
            <h2 className="text-lg font-semibold text-gray-800">테스트</h2>
            <p className="mt-2 text-sm leading-relaxed text-gray-600">
              아래 샘플 라벨 1장만 인쇄합니다. 엑셀 없이 프린터·용지가 맞는지 확인할 때 사용하세요.
            </p>
            <p className="mt-2 text-xs leading-relaxed text-gray-500">
              별도 설치 프로그램은 없습니다. 인쇄는 브라우저 인쇄 창(Ctrl+P)으로 열리며, 여기서 고른
              프린터와 드라이버 설정이 그대로 적용됩니다. 라벨 크기는{" "}
              <code className="rounded bg-gray-100 px-1">globals.css</code>의{" "}
              <code className="rounded bg-gray-100 px-1">@page</code> 기준{" "}
              <strong>45mm × 35mm</strong>입니다.
            </p>
            <div className="mt-4 flex flex-wrap items-center gap-3">
              <div className="rounded border border-dashed border-gray-300 bg-gray-50 p-3">
                <Label data={SAMPLE_LABEL_DATA} />
              </div>
              <button
                type="button"
                onClick={() => printWithTarget("test")}
                className="rounded-lg bg-slate-700 px-5 py-2.5 text-sm font-semibold text-white hover:bg-slate-800"
              >
                샘플 1장만 출력
              </button>
            </div>
          </section>
        )}

        {isPreviewOpen && selectedBatch && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4 print:hidden">
            <div className="max-h-[90vh] w-full max-w-5xl overflow-hidden rounded-xl bg-white shadow-xl">
              <div className="flex items-center justify-between border-b border-gray-200 px-5 py-3">
                <div>
                  <h3 className="text-base font-bold text-gray-800">선택 그룹 보기</h3>
                  <p className="text-xs text-gray-500">
                    그룹 {selectedBatch.index + 1} · {selectedBatch.rowStart}~{selectedBatch.rowEnd}행 ·{" "}
                    {selectedBatch.totalLabels + (batchSize === 53 ? 1 : 0)}장
                    {batchSize === 53 ? " (빈 라벨 1장 포함)" : ""}
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setIsPreviewOpen(false)}
                  className="rounded-md border border-gray-300 px-3 py-1.5 text-sm font-semibold text-gray-700 hover:bg-gray-100"
                >
                  닫기
                </button>
              </div>
              <div className="grid max-h-[80vh] gap-0 overflow-auto md:grid-cols-[1fr_1fr]">
                <div className="border-r border-gray-200">
                  {labels
                    .slice(selectedBatch.startLabel, selectedBatch.endLabel)
                    .map(({ data, key }, idx) => (
                      <div key={`preview-item-${key}`} className="border-b border-gray-100 px-4 py-2">
                      <p className="text-xs font-semibold text-gray-500">
                        {selectedBatch.rowStart + idx}행
                      </p>
                      <p className="truncate text-sm font-semibold text-gray-800">{data.productName}</p>
                      <p className="truncate text-xs text-gray-600">{data.optionName}</p>
                      <p className="truncate text-xs text-gray-500">{data.location}</p>
                      </div>
                    ))}
                </div>
                <div className="bg-gray-50 p-4">
                  <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                    {labels
                      .slice(selectedBatch.startLabel, selectedBatch.endLabel)
                      .map(({ data, key }) => (
                        <div
                          key={`preview-label-${key}`}
                          className="rounded border border-dashed border-gray-300 bg-white p-2"
                        >
                          <Label data={data} />
                        </div>
                      ))}
                    {batchSize === 53 && (
                      <div
                        key="preview-label-blank-trailer"
                        className="rounded border border-dashed border-gray-300 bg-white p-2"
                      >
                        <Label data={BLANK_LABEL_DATA} />
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {isVerifyListOpen && verifyDisplayRows.length > 0 && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4 print:hidden">
            <div className="flex max-h-[90vh] w-full max-w-6xl flex-col overflow-hidden rounded-xl bg-white shadow-xl">
              <div className="flex shrink-0 items-center justify-between border-b border-gray-200 px-5 py-3">
                <div>
                  <h3 className="text-base font-bold text-gray-800">출력 목록 검증</h3>
                  <p className="text-xs text-gray-500">
                    인쇄 순서와 동일 · 총 {verifyDisplayRows.length}장 (한 행 = 출력 1장)
                    {verifyListOverride && (
                      <span className="ml-1 font-medium text-gray-600">· 저장된 작업 기준</span>
                    )}
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    setIsVerifyListOpen(false);
                    setVerifyListOverride(null);
                  }}
                  className="rounded-md border border-gray-300 px-3 py-1.5 text-sm font-semibold text-gray-700 hover:bg-gray-100"
                >
                  닫기
                </button>
              </div>
              <div className="min-h-0 flex-1 overflow-auto">
                <table className="w-max min-w-full border-collapse text-left text-sm">
                  <thead className="sticky top-0 z-10 border-b border-gray-300 bg-gray-200 shadow-sm">
                    <tr>
                      {[
                        ["No", "w-14 text-center"],
                        ["바코드", "min-w-[120px] max-w-[160px]"],
                        ["등록상품명", "min-w-[160px] max-w-[240px]"],
                        ["옵션명", "min-w-[120px] max-w-[200px]"],
                        ["로케이션", "min-w-[88px] max-w-[120px]"],
                        ["박스", "min-w-[72px] max-w-[100px]"],
                        ["날짜", "min-w-[96px] max-w-[120px]"],
                      ].map(([label, cls]) => (
                        <th
                          key={label}
                          scope="col"
                          className={`whitespace-nowrap border border-gray-300 px-2 py-2 text-xs font-semibold text-gray-700 ${cls}`}
                        >
                          {label}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {verifyDisplayRows.map(({ data, key }, rowIdx) => (
                      <tr
                        key={`verify-row-${key}`}
                        className={rowIdx % 2 === 0 ? "bg-white" : "bg-gray-50/80"}
                      >
                        <td className="border border-gray-200 px-2 py-1.5 text-center tabular-nums text-gray-600">
                          {rowIdx + 1}
                        </td>
                        <td
                          className="max-w-[160px] truncate border border-gray-200 px-2 py-1.5 font-mono text-xs text-gray-900"
                          title={data.barcode}
                        >
                          {data.barcode}
                        </td>
                        <td
                          className="max-w-[240px] truncate border border-gray-200 px-2 py-1.5 text-gray-900"
                          title={data.productName}
                        >
                          {data.productName}
                        </td>
                        <td
                          className="max-w-[200px] truncate border border-gray-200 px-2 py-1.5 text-gray-800"
                          title={data.optionName}
                        >
                          {data.optionName}
                        </td>
                        <td
                          className="max-w-[120px] truncate border border-gray-200 px-2 py-1.5 text-gray-800"
                          title={data.location}
                        >
                          {data.location}
                        </td>
                        <td
                          className="max-w-[100px] truncate border border-gray-200 px-2 py-1.5 text-gray-700"
                          title={data.box}
                        >
                          {data.box}
                        </td>
                        <td
                          className="max-w-[120px] truncate border border-gray-200 px-2 py-1.5 text-gray-700"
                          title={data.date}
                        >
                          {data.date}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {loading && <div className="py-10 text-center text-gray-500">파일 읽는 중...</div>}

        </div>
      </div>
    </main>

    {printPortalsReady &&
      createPortal(
        <>
          <div id="print-test" className="print-buffer">
            <Label data={SAMPLE_LABEL_DATA} />
          </div>

          <div id="print-batch" className="print-buffer">
            <div id="print-batch-inner" className="flex flex-col gap-4">
              {selectedBatch && (
                <>
                  {labels
                    .slice(selectedBatch.startLabel, selectedBatch.endLabel)
                    .map(({ data, key }) => (
                      <Label key={`batch-${key}`} data={data} />
                    ))}
                  {batchSize === 53 && (
                    <Label key="batch-blank-trailer" data={BLANK_LABEL_DATA} />
                  )}
                </>
              )}
            </div>
          </div>

          {labels.length > 0 ? (
            <div id="print-area" className="fixed -left-[9999px] top-0">
              <div id="print-area-inner" className="flex flex-col gap-4">
                {labels.map(({ data, key }) => (
                  <Label key={key} data={data} />
                ))}
              </div>
            </div>
          ) : null}
        </>,
        document.body
      )}
    </>
  );
}
