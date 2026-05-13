"use client";
import { useLayoutEffect, useRef } from "react";
import JsBarcode from "jsbarcode";
import { LabelData } from "@/lib/excel";

interface Props {
  data: LabelData;
}

/** 긴 CODE128은 모듈 width를 줄여 라벨 너비 안에 들어가게 함 (최소는 스캔 가능 수준으로 유지) */
function barcodeModuleWidth(barcodeLength: number): number {
  const maxW = 0.95;
  const minW = 0.52;
  if (barcodeLength <= 10) return maxW;
  if (barcodeLength >= 30) return minW;
  return maxW - ((barcodeLength - 10) / 20) * (maxW - minW);
}

export default function Label({ data }: Props) {
  const barcodeRef = useRef<SVGSVGElement>(null);

  useLayoutEffect(() => {
    const el = barcodeRef.current;
    if (!el || !data.barcode) return;
    el.innerHTML = "";
    try {
      const len = data.barcode.length;
      JsBarcode(el, data.barcode, {
        format: "CODE128",
        width: barcodeModuleWidth(len),
        height: 21,
        displayValue: true,
        fontSize: 5.5,
        margin: 0,
        lineColor: "#000000",
        background: "#ffffff",
      });
    } catch {
      // invalid barcode
    }
  }, [data]);

  return (
    <div className="label-box">
      {/* 텍스트 영역 */}
      <div className="label-info">
        <p className="label-product">{data.productName}</p>
        <p className="label-option">{data.optionName}</p>
        <p className="label-location">{data.location}</p>
      </div>

      {/* 바코드 */}
      <div className="label-barcode">
        <svg ref={barcodeRef} />
      </div>

      {/* 실제 라벨 하단 고정문구(스티커 인쇄물) 영역 여백 */}
      <div className="label-footer-space" />
    </div>
  );
}
