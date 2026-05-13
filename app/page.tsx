"use client";
import { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { parseExcel, listExcelSheets, LabelData, SAMPLE_LABEL_DATA } from "@/lib/excel";
import Label from "@/components/Label";

type NavTab = "print" | "jobs" | "test";

interface JobRecord {
  id: string;
  createdAt: string;
  fileName: string;
  sheetName: string;
  totalLabels: number;
  rowCount: number;
  adjustedQuantityCount: number;
  importHint: string | null;
}

const JOBS_STORAGE_KEY = "label-printer-jobs";
const BATCH_OPTIONS = [50, 20, 10, 1] as const;

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
      setTimeout(() => {
        window.print();
      }, 0);
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
  const [jobHistory, setJobHistory] = useState<JobRecord[]>([]);
  const [selectedJobId, setSelectedJobId] = useState<string | null>(null);
  const [batchSize, setBatchSize] = useState<(typeof BATCH_OPTIONS)[number]>(20);
  const [selectedBatchIndex, setSelectedBatchIndex] = useState(0);
  const [printedBatchIndexes, setPrintedBatchIndexes] = useState<number[]>([]);
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const [printPortalsReady, setPrintPortalsReady] = useState(false);

  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(JOBS_STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw) as JobRecord[];
      setJobHistory(Array.isArray(parsed) ? parsed : []);
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

  const batchCards = useMemo(() => {
    const cards: Array<{ index: number; start: number; end: number; labels: { data: LabelData; key: number }[] }> =
      [];
    for (let start = 0, idx = 0; start < labels.length; start += batchSize, idx += 1) {
      const labelsInCard = labels.slice(start, start + batchSize);
      cards.push({
        index: idx,
        start: start + 1,
        end: start + labelsInCard.length,
        labels: labelsInCard,
      });
    }
    return cards;
  }, [labels, batchSize]);

  const selectedBatch = useMemo(
    () => batchCards.find((card) => card.index === selectedBatchIndex) ?? null,
    [batchCards, selectedBatchIndex]
  );

  const printedLabelCount = useMemo(
    () =>
      batchCards
        .filter((card) => printedBatchIndexes.includes(card.index))
        .reduce((sum, card) => sum + card.labels.length, 0),
    [batchCards, printedBatchIndexes]
  );

  function saveJob(record: JobRecord) {
    setJobHistory((prev) => {
      const next = [record, ...prev].slice(0, 100);
      window.localStorage.setItem(JOBS_STORAGE_KEY, JSON.stringify(next));
      return next;
    });
    setSelectedJobId(record.id);
  }

  async function parseAndExpand(file: File, sheetName: string) {
    const rows = await parseExcel(file, { sheetName });
    const adjustedQuantityCount = rows.filter((row) => row.quantityAdjusted).length;
    const expanded: { data: LabelData; key: number }[] = [];
    let key = 0;
    for (const row of rows) {
      for (let i = 0; i < row.quantity; i++) {
        expanded.push({ data: row, key: key++ });
      }
    }
    setLabels(expanded);

    let hint: string | null = null;
    if (expanded.length === 0) {
      hint =
        "파일은 읽혔지만 바코드가 있는 데이터 행이 없습니다. 첫 줄에 열 이름(바코드, 수량, 등록상품명 등)이 오도록 하거나, 시트 선택이 올바른지 확인해 주세요.";
    } else if (adjustedQuantityCount > 0) {
      hint = `일부 행의 수량 값이 비정상적으로 보여 ${adjustedQuantityCount}개 행을 자동 보정했습니다. 수량 컬럼(수량/개수/Qty) 값을 확인해 주세요.`;
    }
    setImportHint(hint);
    setSelectedBatchIndex(0);
    setPrintedBatchIndexes([]);

    saveJob({
      id: crypto.randomUUID(),
      createdAt: new Date().toISOString(),
      fileName: file.name,
      sheetName,
      totalLabels: expanded.length,
      rowCount: rows.length,
      adjustedQuantityCount,
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

  function handleBatchSizeChange(size: (typeof BATCH_OPTIONS)[number]) {
    setBatchSize(size);
    // 그룹 단위가 바뀌면 출력 진행 상태를 초기화한다.
    setSelectedBatchIndex(0);
    setPrintedBatchIndexes([]);
  }

  function handlePrintBatch() {
    if (!selectedBatch || selectedBatch.labels.length === 0) return;
    printWithTarget("batch");
    setPrintedBatchIndexes((prev) => (prev.includes(selectedBatch.index) ? prev : [...prev, selectedBatch.index]));
  }

  function handleOpenBatchPreview() {
    if (!selectedBatch || selectedBatch.labels.length === 0) return;
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
                    선택 그룹: {selectedBatch ? `${selectedBatch.start}~${selectedBatch.end} (${selectedBatch.labels.length}장)` : "-"}
                  </span>
                </div>

                <div className="mt-4 flex flex-wrap items-center gap-2">
                  <span className="mr-1 text-sm font-medium text-gray-700">출력 단위</span>
                  {BATCH_OPTIONS.map((size) => (
                    <button
                      key={size}
                      type="button"
                      onClick={() => handleBatchSizeChange(size)}
                      className={`rounded-md px-3 py-1.5 text-sm font-semibold ${
                        batchSize === size
                          ? "bg-blue-600 text-white"
                          : "border border-gray-300 bg-white text-gray-700 hover:bg-gray-100"
                      }`}
                    >
                      {size}개
                    </button>
                  ))}
                </div>

                <div className="mt-4 flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={handlePrintBatch}
                    disabled={!selectedBatch || selectedBatch.labels.length === 0}
                    className="rounded-lg bg-blue-600 px-5 py-2 text-sm font-bold text-white hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-gray-300"
                  >
                    선택 그룹 출력
                  </button>
                  <button
                    type="button"
                    onClick={handleOpenBatchPreview}
                    disabled={!selectedBatch || selectedBatch.labels.length === 0}
                    className="rounded-lg border border-gray-300 bg-white px-5 py-2 text-sm font-bold text-gray-700 hover:bg-gray-100 disabled:cursor-not-allowed disabled:bg-gray-100 disabled:text-gray-400"
                  >
                    선택 그룹 보기
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
                  {batchCards.map((card) => {
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
                          {card.start}~{card.end}번 ({card.labels.length}장)
                        </p>
                        <p className="mt-2 truncate text-xs">
                          {card.labels[0]?.data.productName ?? "-"}
                        </p>
                        <p className="truncate text-xs">{card.labels[0]?.data.optionName ?? "-"}</p>
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
            {jobHistory.length === 0 ? (
              <p className="mt-4 rounded-lg bg-gray-50 p-4 text-sm text-gray-500">저장된 작업 이력이 없습니다.</p>
            ) : (
              <div className="mt-4 grid gap-4 md:grid-cols-[1.2fr_1fr]">
                <div className="max-h-[420px] overflow-auto rounded-lg border border-gray-200">
                  {jobHistory.map((job) => (
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
                  ))}
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
                    그룹 {selectedBatch.index + 1} · {selectedBatch.start}~{selectedBatch.end}번 · {selectedBatch.labels.length}장
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
                  {selectedBatch.labels.map(({ data, key }, idx) => (
                    <div key={`preview-item-${key}`} className="border-b border-gray-100 px-4 py-2">
                      <p className="text-xs font-semibold text-gray-500">
                        {selectedBatch.start + idx}번
                      </p>
                      <p className="truncate text-sm font-semibold text-gray-800">{data.productName}</p>
                      <p className="truncate text-xs text-gray-600">{data.optionName}</p>
                      <p className="truncate text-xs text-gray-500">{data.location}</p>
                    </div>
                  ))}
                </div>
                <div className="bg-gray-50 p-4">
                  <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                    {selectedBatch.labels.map(({ data, key }) => (
                      <div key={`preview-label-${key}`} className="rounded border border-dashed border-gray-300 bg-white p-2">
                        <Label data={data} />
                      </div>
                    ))}
                  </div>
                </div>
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
              {(selectedBatch?.labels ?? []).map(({ data, key }) => (
                <Label key={`batch-${key}`} data={data} />
              ))}
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
