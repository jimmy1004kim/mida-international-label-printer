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
    if (barcodeRef.current && data.barcode) {
      try {
        JsBarcode(barcodeRef.current, data.barcode, {
          format: "CODE128",
          width: 1,
          height: 22,
          displayValue: true,
          fontSize: 5.5,
          margin: 0,
        });
      } catch {
        // invalid barcode
      }
    }
  }, [data.barcode]);

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
