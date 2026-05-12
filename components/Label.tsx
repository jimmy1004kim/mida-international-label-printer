"use client";
import { useEffect, useRef } from "react";
import JsBarcode from "jsbarcode";
import { LabelData } from "@/lib/excel";

interface Props {
  data: LabelData;
}

export default function Label({ data }: Props) {
  const barcodeRef = useRef<SVGSVGElement>(null);

  useEffect(() => {
    const el = barcodeRef.current;
    if (!el || !data.barcode) return;
    el.innerHTML = "";
    try {
      JsBarcode(el, data.barcode, {
        format: "CODE128",
        width: 0.95,
        height: 21,
        displayValue: true,
        fontSize: 5.5,
        margin: 0,
      });
    } catch {
      // invalid barcode
    }
    // 바코드 문자열이 같아도 행(상품명 등)이 다르면 다시 그려야 함
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
