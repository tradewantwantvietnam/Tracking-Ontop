"use client";

import {
  useEffect,
  useMemo,
  useState
} from "react";

import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis
} from "recharts";

import {
  parseActual,
  parseAllocation,
  type Actual,
  type Allocation
} from "../lib/parse";


type Summary = {
  regionName: string;
  areaName: string;
  programCode: string;
  programName: string;
  structure: string;
  allocation: number;
  used: number;
  discount: number;
  remaining: number;
  percent: number;
};

type NppSummary = {
  regionName: string;
  areaName: string;
  programCode: string;
  npp: string;
  used: number;
  discount: number;
};

type Warning = {
  regionName: string;
  areaName: string;
  programCode: string;
  npp: string;
  orderCode: string;
  discount: number;
  reason: string;
};

type Report = {
  summaries: Summary[];
  npps: NppSummary[];
  warnings: Warning[];
};


// =============================================
// ĐỊNH DẠNG SỐ VÀ TIỀN
// =============================================

function money(value: number): string {
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0
  }).format(value);
}

function fmt(value: number): string {
  return new Intl.NumberFormat("vi-VN", {
    maximumFractionDigits: 2
  }).format(value);
}

function normalizeText(value: string): string {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeKey(value: string): string {
  return normalizeText(value).toLocaleLowerCase("vi");
}

function dataKey(
  areaName: string,
  programCode: string
): string {
  return (
    normalizeKey(areaName) +
    "|||" +
    normalizeKey(programCode)
  );
}

function uniqueValues(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "vi"));
}


// =============================================
// XỬ LÝ VÀ ĐỐI CHIẾU DỮ LIỆU
// =============================================

function buildReport(
  allocations: Allocation[],
  actuals: Actual[]
): Report {

  // -------------------------------------------
  // A. TẠO DANH SÁCH PHÂN BỔ
  // Khóa mapping: Tên vùng + Mã CT
  // -------------------------------------------

  const allocationMap = new Map<
    string,
    Allocation
  >();

  for (const allocation of allocations) {
    const key = dataKey(
      allocation.areaName,
      allocation.programCode
    );

    if (allocationMap.has(key)) {
      throw new Error(
        "File phân bổ có Tên vùng + Mã CT bị trùng: " +
        allocation.areaName +
        " / " +
        allocation.programCode +
        ". Hãy kiểm tra và tổng hợp thành một dòng."
      );
    }

    allocationMap.set(key, allocation);
  }


  // -------------------------------------------
  // B. KHỞI TẠO TỔNG HỢP VÙNG + MÃ CT
  // -------------------------------------------

  const summaryMap = new Map<
