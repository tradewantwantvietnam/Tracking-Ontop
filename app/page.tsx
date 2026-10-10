"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

import {
  parseActual,
  parseAllocation,
  parseProgramMapping,
  type Actual,
  type Allocation,
  type ProgramMapping
} from "../lib/parse";


// ==================================================
// KIỂU DỮ LIỆU
// ==================================================

type StatusName =
  | "Bình thường"
  | "Sắp đầy"
  | "Vượt định biên";

type Summary = {
  regionName: string;
  areaName: string;
  programName: string;
  allocation: number;
  used: number;
  discount: number;
  percent: number;
  status: StatusName;
};

type NppSummary = {
  regionName: string;
  areaName: string;
  programName: string;
  npp: string;
  used: number;
  discount: number;
};

type Warning = {
  regionName: string;
  areaName: string;
  programName: string;
  npp: string;
  orderCode: string;
  discount: number;
  reason: string;
};

type UsedOrder = {
  regionName: string;
  areaName: string;
  programName: string;
  orderCode: string;
  slotOwner: string;
  npps: string[];
  discount: number;
};

type Report = {
  summaries: Summary[];
  npps: NppSummary[];
  warnings: Warning[];
  usedOrders: UsedOrder[];
};

type ProgramGroup = {
  programName: string;
  rows: Summary[];
  allocation: number;
  used: number;
  discount: number;
  percent: number;
  status: StatusName;
};

type UsedScope = {
  programName: string;
  regionName?: string;
  areaName?: string;
};


// ==================================================
// HÀM HỖ TRỢ
// ==================================================

function getCurrentMonth(): string {
  const date = new Date();

  return (
    date.getFullYear() +
    "-" +
    String(date.getMonth() + 1).padStart(2, "0")
  );
}

function cleanText(value: unknown): string {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeKey(value: string): string {
  return cleanText(value).toLocaleLowerCase("vi");
}

function mappingKey(
  areaName: string,
  programName: string
): string {
  return (
    normalizeKey(areaName) +
    "|||" +
    normalizeKey(programName)
  );
}

function detailKey(
  regionName: string,
  areaName: string,
  programName: string
): string {
  return [
    normalizeKey(regionName),
    normalizeKey(areaName),
    normalizeKey(programName)
  ].join("|||");
}

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

function uniqueInOrder(values: string[]): string[] {
  const result: string[] = [];
  const seen = new Set<string>();

  for (const value of values) {
    const name = cleanText(value);

    if (!name) continue;

    const key = normalizeKey(name);

    if (!seen.has(key)) {
      seen.add(key);
      result.push(name);
    }
  }

  return result;
}


// ==================================================
// TRẠNG THÁI TIẾN ĐỘ
// ==================================================

function getPercent(
  allocation: number,
  used: number
): number {
  if (allocation > 0) {
    return (used / allocation) * 100;
  }

  return used > 0 ? 100 : 0;
}

function getStatus(
  allocation: number,
  used: number
): StatusName {
  const percent = getPercent(allocation, used);

  if (percent >= 100) {
    return "Vượt định biên";
  }

  if (percent >= 80) {
    return "Sắp đầy";
  }

  return "Bình thường";
}

function statusClass(status: string): string {
  if (status === "Vượt định biên") {
    return "statusOver";
  }

  if (status === "Sắp đầy") {
    return "statusNear";
  }

  return "statusNormal";
}


// ==================================================
// XỬ LÝ DỮ LIỆU
// ==================================================

function buildReport(
  allocations: Allocation[],
  mappings: ProgramMapping[],
  actuals: Actual[]
): Report {

  // ----------------------------------------------
  // A. TẠO BẢNG QUY ƯỚC MÃ CT -> TÊN CT
  // ----------------------------------------------

  const programByCode: Map<string, string> =
    new Map();

  for (const item of mappings) {
    const code = normalizeKey(item.programCode);
    const title = cleanText(item.programName);

    if (!code || !title) continue;

    if (programByCode.has(code)) {
      const oldTitle = programByCode.get(code)!;

      if (normalizeKey(oldTitle) !== normalizeKey(title)) {
        throw new Error(
          "Một mã CT đang được gán cho nhiều Tên CT. " +
          "Hãy kiểm tra file Tên CT."
        );
      }

      continue;
    }

    programByCode.set(code, title);
  }

  if (programByCode.size === 0) {
    throw new Error(
      "File Tên CT chưa có quy ước Mã CT / Tên CT hợp lệ."
    );
  }


  // ----------------------------------------------
  // B. PHÂN BỔ THEO TÊN VÙNG + TÊN CT
  //
  // File 5.1 không có cột miền được phép đọc,
  // do đó Tên vùng + Tên CT phải xác định duy nhất
  // một dòng phân bổ.
  // ----------------------------------------------

  const allocationByAreaAndProgram: Map<
    string,
    Allocation
  > = new Map();

  const summaryMap: Map<string, Summary> =
    new Map();

  for (const item of allocations) {
    const programName = cleanText(item.programName);

    const key = mappingKey(
      item.areaName,
      programName
    );

    if (allocationByAreaAndProgram.has(key)) {
      const existing =
        allocationByAreaAndProgram.get(key)!;

      throw new Error(
        "Tên vùng + Tên CT bị trùng hoặc không xác định được miền: " +
        item.areaName +
        " / " +
        programName +
        ". Vì file 5.1 không có cột miền trong phạm vi dữ liệu, " +
        "cặp này phải xác định duy nhất một dòng phân bổ. " +
        "Hãy kiểm tra file phân bổ."
      );
    }

    allocationByAreaAndProgram.set(key, item);

    const parentKey = detailKey(
      item.regionName,
      item.areaName,
      programName
    );

    summaryMap.set(parentKey, {
      regionName: item.regionName,
      areaName: item.areaName,
      programName,
      allocation: Number(item.allocation) || 0,
      used: 0,
      discount: 0,
      percent: 0,
      status: "Bình thường"
    });
  }


  // ----------------------------------------------
  // C. NHÓM ĐƠN HÀNG
  //
  // Khử trùng theo:
  // Tên CT + Tên vùng + Mã đơn hàng.
  //
  // Nhiều mã CT cùng Tên CT sẽ chỉ tạo tối đa
  // một suất cho cùng một đơn hàng trong vùng.
  // ----------------------------------------------

  type OrderRecord = {
    row: Actual;
    allocation: Allocation;
    programName: string;
  };

  type OrderGroup = {
    regionName: string;
    areaName: string;
    programName: string;
    records: OrderRecord[];
    discount: number;
    hasPositive: boolean;
  };

  const orderGroups: Map<string, OrderGroup> =
    new Map();


  // ----------------------------------------------
  // D. TỔNG HỢP THEO NPP
  // ----------------------------------------------

  const nppMap: Map<string, NppSummary> =
    new Map();

  function addNpp(
    regionName: string,
    areaName: string,
    programName: string,
    nppName: string,
    discountToAdd: number,
    usedToAdd: number
  ): void {
    const displayName =
      cleanText(nppName) ||
      "(Không có tên NPP)";

    const key =
      detailKey(
        regionName,
        areaName,
        programName
      ) +
      "|||" +
      normalizeKey(displayName);

    if (!nppMap.has(key)) {
      nppMap.set(key, {
        regionName,
        areaName,
        programName,
        npp: displayName,
        used: 0,
        discount: 0
      });
    }

    const item = nppMap.get(key)!;

    item.discount += discountToAdd;
    item.used += usedToAdd;
  }


  // ----------------------------------------------
  // E. CẢNH BÁO
  // ----------------------------------------------

  const unmatchedWarnings: Map<string, Warning> =
    new Map();

  const missingOrderWarnings: Warning[] = [];


  // ----------------------------------------------
  // F. ĐỌC TỪNG DÒNG FILE 5.1
  // ----------------------------------------------

  actuals.forEach((row, rowIndex) => {
    const programCode = normalizeKey(row.programCode);

    // Tìm Tên CT từ file quy ước.
    const programName =
      programByCode.get(programCode);

    // --------------------------------------------
    // F1. MÃ CT CHƯA CÓ TRONG FILE TÊN CT
    // --------------------------------------------

    if (!programName) {
      const orderCode = cleanText(row.orderCode);

      const warningKey = [
        "UNKNOWN_PROGRAM",
        normalizeKey(row.areaName),
        programCode,
        normalizeKey(orderCode || String(rowIndex))
      ].join("|||");

      if (!unmatchedWarnings.has(warningKey)) {
        unmatchedWarnings.set(warningKey, {
          regionName: "",
          areaName: row.areaName,
          programName: "Chưa xác định",
          npp: row.npp,
          orderCode,
          discount: 0,
          reason:
            "Mã CT chưa được khai báo trong file Tên CT."
        });
      }

      unmatchedWarnings.get(warningKey)!.discount +=
        Number(row.discount) || 0;

      return;
    }


    // --------------------------------------------
    // F2. TÊN VÙNG + TÊN CT KHÔNG CÓ PHÂN BỔ
    // --------------------------------------------

    const allocation =
      allocationByAreaAndProgram.get(
        mappingKey(row.areaName, programName)
      );

    if (!allocation) {
      const orderCode = cleanText(row.orderCode);

      const warningKey = [
        "NO_ALLOCATION",
        normalizeKey(row.areaName),
        normalizeKey(programName),
        normalizeKey(orderCode || String(rowIndex))
      ].join("|||");

      if (!unmatchedWarnings.has(warningKey)) {
        unmatchedWarnings.set(warningKey, {
          regionName: "",
          areaName: row.areaName,
          programName,
          npp: row.npp,
          orderCode,
          discount: 0,
          reason:
            "Không có cặp Tên vùng + Tên CT trong file phân bổ."
        });
      }

      unmatchedWarnings.get(warningKey)!.discount +=
        Number(row.discount) || 0;

      return;
    }


    const parentKey = detailKey(
      allocation.regionName,
      allocation.areaName,
      programName
    );

    const summary = summaryMap.get(parentKey)!;


    // --------------------------------------------
    // F3. THIẾU MÃ ĐƠN HÀNG
    //
    // Có phân bổ: vẫn cộng AK nhưng không tính suất.
    // --------------------------------------------

    if (!cleanText(row.orderCode)) {
      const discount = Number(row.discount) || 0;

      summary.discount += discount;

      addNpp(
        allocation.regionName,
        allocation.areaName,
        programName,
        row.npp,
        discount,
        0
      );

      missingOrderWarnings.push({
        regionName: allocation.regionName,
        areaName: allocation.areaName,
        programName,
        npp: row.npp,
        orderCode: "",
        discount,
        reason:
          "Thiếu mã đơn hàng. Tiền AK vẫn được cộng; " +
          "không tính suất để tránh đếm trùng sai."
      });

      return;
    }


    // --------------------------------------------
    // F4. NHÓM ĐƠN HÀNG
    // --------------------------------------------

    const orderKey = [
      normalizeKey(programName),
      normalizeKey(allocation.areaName),
      normalizeKey(row.orderCode)
    ].join("|||");

    if (!orderGroups.has(orderKey)) {
      orderGroups.set(orderKey, {
        regionName: allocation.regionName,
        areaName: allocation.areaName,
        programName,
        records: [],
        discount: 0,
        hasPositive: false
      });
    }

    const group = orderGroups.get(orderKey)!;

    group.records.push({
      row,
      allocation,
      programName
    });

    // Cộng tất cả AK của các dòng trong đơn hàng.
    group.discount += Number(row.discount) || 0;

    // Chỉ cần một dòng AK > 0.
    if (Number(row.discount) > 0) {
      group.hasPositive = true;
    }
  });


  // ----------------------------------------------
  // G. TÍNH SUẤT ĐÃ DÙNG VÀ TIỀN CHIẾT KHẤU
  // ----------------------------------------------

  const usedOrders: UsedOrder[] = [];

  for (const group of orderGroups.values()) {
    const parentKey = detailKey(
      group.regionName,
      group.areaName,
      group.programName
    );

    const summary = summaryMap.get(parentKey);

    if (!summary) continue;

    // Tổng tiền AK của mọi dòng thuộc đơn hàng.
    summary.discount += group.discount;

    // Một đơn hàng đủ điều kiện chỉ được tính một suất.
    if (group.hasPositive) {
      summary.used += 1;
    }


    // --------------------------------------------
    // G1. GỘP TIỀN THEO NPP
    // --------------------------------------------

    type NppDiscount = {
      name: string;
      discount: number;
    };

    const nppDiscounts: Map<string, NppDiscount> =
      new Map();

    for (const record of group.records) {
      const name =
        cleanText(record.row.npp) ||
        "(Không có tên NPP)";

      const key = normalizeKey(name);

      if (!nppDiscounts.has(key)) {
        nppDiscounts.set(key, {
          name,
          discount: 0
        });
      }

      nppDiscounts.get(key)!.discount +=
        Number(record.row.discount) || 0;
    }


    // --------------------------------------------
    // G2. GÁN SUẤT CHO NPP
    //
    // Nếu một đơn hàng có nhiều NPP:
    // - Chỉ giao một suất cho NPP đầu tiên có AK > 0.
    // - Tiền AK cộng riêng theo từng NPP.
    // --------------------------------------------

    const firstPositiveRecord = group.records.find(
      record => Number(record.row.discount) > 0
    );

    const slotOwner =
      firstPositiveRecord
        ? (
            cleanText(firstPositiveRecord.row.npp) ||
            "(Không có tên NPP)"
          )
        : "";

    for (const item of nppDiscounts.values()) {
      const usedToAdd =
        group.hasPositive &&
        normalizeKey(item.name) === normalizeKey(slotOwner)
          ? 1
          : 0;

      addNpp(
        group.regionName,
        group.areaName,
        group.programName,
        item.name,
        item.discount,
        usedToAdd
      );
    }


    // --------------------------------------------
    // G3. LƯU CHI TIẾT ĐỂ TRA CỨU SUẤT
    // --------------------------------------------

    if (group.hasPositive) {
      usedOrders.push({
        regionName: group.regionName,
        areaName: group.areaName,
        programName: group.programName,
        orderCode: cleanText(
          group.records[0].row.orderCode
        ),
        slotOwner,
        npps: [...nppDiscounts.values()].map(
          item => item.name
        ),
        discount: group.discount
      });
    }
  }


  // ----------------------------------------------
  // H. TỶ LỆ SỬ DỤNG VÀ TRẠNG THÁI
  // ----------------------------------------------

  const summaries = [...summaryMap.values()].map(
    item => {
      const percent = getPercent(
        item.allocation,
        item.used
      );

      return {
        ...item,
        percent,
        status: getStatus(
          item.allocation,
          item.used
        )
      };
    }
  );


  return {
    summaries,
    npps: [...nppMap.values()],
    warnings: [
      ...unmatchedWarnings.values(),
      ...missingOrderWarnings
    ],
    usedOrders
  };
}


// ==================================================
// TRANG DASHBOARD
// ==================================================

export default function Home() {
  const [month, setMonth] = useState(getCurrentMonth());

  const [availableMonths, setAvailableMonths] =
    useState<string[]>([]);

  const [allocation, setAllocation] =
    useState<Allocation[] | null>(null);

  const [programMap, setProgramMap] =
    useState<ProgramMapping[] | null>(null);

  const [actual, setActual] =
    useState<Actual[] | null>(null);

  const [report, setReport] =
    useState<Report | null>(null);

  const [allocationName, setAllocationName] =
    useState("");

  const [programMapName, setProgramMapName] =
    useState("");

  const [actualName, setActualName] =
    useState("");

  const [busy, setBusy] = useState(false);

  const [historyLoading, setHistoryLoading] =
    useState(false);

  const [uploadOpen, setUploadOpen] =
    useState(false);

  const [status, setStatus] = useState("");

  const [statusType, setStatusType] =
    useState<"success" | "error" | "info">("info");

  const [savedAt, setSavedAt] = useState("");

  const [region, setRegion] = useState("ALL");

  const [area, setArea] = useState("ALL");

  const [selectedProgram, setSelectedProgram] =
    useState("ALL");

  const [search, setSearch] = useState("");

  const [expandedZones, setExpandedZones] =
    useState<string[]>([]);

  const [selectedUsedScope, setSelectedUsedScope] =
    useState<UsedScope | null>(null);


  // ----------------------------------------------
  // TẢI LỊCH SỬ THEO THÁNG
  // ----------------------------------------------

  useEffect(() => {
    let cancelled = false;

    async function loadMonth() {
      setHistoryLoading(true);
      setReport(null);
      setSavedAt("");
      setStatus("");
      setExpandedZones([]);
      setSelectedUsedScope(null);

      try {
        const [monthsResponse, reportResponse] =
          await Promise.all([
            fetch("/api/save?action=months", {
              cache: "no-store"
            }),

            fetch(
              "/api/save?action=report&month=" +
              encodeURIComponent(month),
              {
                cache: "no-store"
              }
            )
          ]);

        const monthsData =
          await monthsResponse.json();

        const reportData =
          await reportResponse.json();

        if (!monthsResponse.ok || !monthsData.ok) {
          throw new Error(
            monthsData.error ||
            "Không tải được danh sách tháng."
          );
        }

        if (!reportResponse.ok || !reportData.ok) {
          throw new Error(
            reportData.error ||
            "Không tải được dữ liệu lịch sử."
          );
        }

        if (cancelled) return;

        setAvailableMonths(
          uniqueInOrder([
            month,
            ...(monthsData.months || [])
          ]).sort((a, b) => b.localeCompare(a))
        );

        if (
          reportData.found &&
          reportData.report
        ) {
          const loadedReport =
            reportData.report as Report;

          // Dữ liệu lịch sử từ phiên bản cũ
          // có thể chưa lưu chi tiết đơn hàng.
          if (!Array.isArray(loadedReport.usedOrders)) {
            loadedReport.usedOrders = [];
          }

          setReport(loadedReport);

          setSavedAt(
            reportData.updatedAt
              ? new Date(
                  reportData.updatedAt
                ).toLocaleString("vi-VN")
              : ""
          );

          setStatus(
            "Đã tải dữ liệu tháng " + month + "."
          );

          setStatusType("success");

        } else {
          setReport(null);

          setStatus(
            "Tháng " + month +
            " chưa có dữ liệu. Hãy mở khu vực Tải dữ liệu."
          );

          setStatusType("info");
        }

      } catch (error: any) {
        if (cancelled) return;

        setReport(null);

        setStatus(
          error?.message ||
          "Không tải được dữ liệu lịch sử."
        );

        setStatusType("error");

      } finally {
        if (!cancelled) {
          setHistoryLoading(false);
        }
      }
    }

    void loadMonth();

    return () => {
      cancelled = true;
    };
  }, [month]);


  // ----------------------------------------------
  // DANH SÁCH CHƯƠNG TRÌNH, MIỀN, VÙNG
  // ----------------------------------------------

  const programs = useMemo(
    () =>
      uniqueInOrder(
        (report?.summaries || []).map(
          item => item.programName
        )
      ),
    [report]
  );

  const regions = useMemo(
    () =>
      uniqueInOrder(
        (report?.summaries || [])
          .filter(
            item =>
              selectedProgram === "ALL" ||
              normalizeKey(item.programName) ===
                normalizeKey(selectedProgram)
          )
          .map(item => item.regionName)
      ),
    [report, selectedProgram]
  );

  const areas = useMemo(
    () =>
      uniqueInOrder(
        (report?.summaries || [])
          .filter(
            item =>
              (
                selectedProgram === "ALL" ||
                normalizeKey(item.programName) ===
                  normalizeKey(selectedProgram)
              ) &&
              (
                region === "ALL" ||
                normalizeKey(item.regionName) ===
                  normalizeKey(region)
              )
          )
          .map(item => item.areaName)
      ),
    [report, selectedProgram, region]
  );


  useEffect(() => {
    if (
      region !== "ALL" &&
      !regions.some(
        item => normalizeKey(item) === normalizeKey(region)
      )
    ) {
      setRegion("ALL");
    }
  }, [regions, region]);


  useEffect(() => {
    if (
      area !== "ALL" &&
      !areas.some(
        item => normalizeKey(item) === normalizeKey(area)
      )
    ) {
      setArea("ALL");
    }
  }, [areas, area]);


  // ----------------------------------------------
  // DỮ LIỆU CƠ SỞ THEO BỘ LỌC
  // ----------------------------------------------

  const baseFiltered = useMemo(
    () =>
      (report?.summaries || []).filter(
        item =>
          (
            selectedProgram === "ALL" ||
            normalizeKey(item.programName) ===
              normalizeKey(selectedProgram)
          ) &&
          (
            region === "ALL" ||
            normalizeKey(item.regionName) ===
              normalizeKey(region)
          ) &&
          (
            area === "ALL" ||
            normalizeKey(item.areaName) ===
              normalizeKey(area)
          )
      ),
    [report, selectedProgram, region, area]
  );

  const baseNpps = useMemo(
    () =>
      (report?.npps || []).filter(
        item =>
          (
            selectedProgram === "ALL" ||
            normalizeKey(item.programName) ===
              normalizeKey(selectedProgram)
          ) &&
          (
            region === "ALL" ||
            normalizeKey(item.regionName) ===
              normalizeKey(region)
          ) &&
          (
            area === "ALL" ||
            normalizeKey(item.areaName) ===
              normalizeKey(area)
          )
      ),
    [report, selectedProgram, region, area]
  );


  // ----------------------------------------------
  // TÌM KIẾM THEO MIỀN / VÙNG / NPP / TÊN CT
  // ----------------------------------------------

  const query = normalizeKey(search);

  function matchesContext(item: {
    regionName: string;
    areaName: string;
    programName: string;
  }): boolean {
    if (!query) return true;

    return [
      item.regionName,
      item.areaName,
      item.programName
    ].some(
      value =>
        normalizeKey(value).includes(query)
    );
  }


  const parentsMatchedByNpp = useMemo(() => {
    const result = new Set<string>();

    if (!query) return result;

    for (const item of baseNpps) {
      if (normalizeKey(item.npp).includes(query)) {
        result.add(
          detailKey(
            item.regionName,
            item.areaName,
            item.programName
          )
        );
      }
    }

    return result;
  }, [baseNpps, query]);


  const filtered = useMemo(
    () =>
      baseFiltered.filter(item => {
        if (!query) return true;

        return (
          matchesContext(item) ||
          parentsMatchedByNpp.has(
            detailKey(
              item.regionName,
              item.areaName,
              item.programName
            )
          )
        );
      }),
    [baseFiltered, query, parentsMatchedByNpp]
  );


  const npps = useMemo(
    () =>
      baseNpps.filter(item => {
        if (!query) return true;

        return (
          normalizeKey(item.npp).includes(query) ||
          matchesContext(item)
        );
      }),
    [baseNpps, query]
  );


  // ----------------------------------------------
  // KPI THEO DỮ LIỆU ĐÃ LỌC
  // ----------------------------------------------

  const total = useMemo(
    () =>
      filtered.reduce(
        (sum, item) => ({
          allocation: sum.allocation + item.allocation,
          used: sum.used + item.used,
          discount: sum.discount + item.discount
        }),
        {
          allocation: 0,
          used: 0,
          discount: 0
        }
      ),
    [filtered]
  );

  const totalPercent = getPercent(
    total.allocation,
    total.used
  );


  // ----------------------------------------------
  // NHÓM TỔNG CHƯƠNG TRÌNH
  // ----------------------------------------------

  const programGroups = useMemo(() => {
    const map: Map<string, ProgramGroup> =
      new Map();

    for (const item of filtered) {
      const key = normalizeKey(item.programName);

      if (!map.has(key)) {
        map.set(key, {
          programName: item.programName,
          rows: [],
          allocation: 0,
          used: 0,
          discount: 0,
          percent: 0,
          status: "Bình thường"
        });
      }

      const group = map.get(key)!;

      group.rows.push(item);
      group.allocation += item.allocation;
      group.used += item.used;
      group.discount += item.discount;
    }

    return [...map.values()].map(group => {
      const percent = getPercent(
        group.allocation,
        group.used
      );

      return {
        ...group,
        percent,
        status: getStatus(
          group.allocation,
          group.used
        )
      };
    });
  }, [filtered]);


  // ----------------------------------------------
  // TRA CỨU CÁC SUẤT ĐÃ DÙNG
  // ----------------------------------------------

  const activeUsedOrders = useMemo(() => {
    if (!report || !selectedUsedScope) {
      return [];
    }

    const allOrders = report.usedOrders || [];

    return allOrders.filter(item => {
      if (
        normalizeKey(item.programName) !==
        normalizeKey(selectedUsedScope.programName)
      ) {
        return false;
      }

      if (
        selectedUsedScope.regionName &&
        normalizeKey(item.regionName) !==
          normalizeKey(selectedUsedScope.regionName)
      ) {
        return false;
      }

      if (
        selectedUsedScope.areaName &&
        normalizeKey(item.areaName) !==
          normalizeKey(selectedUsedScope.areaName)
      ) {
        return false;
      }

      if (
        region !== "ALL" &&
        normalizeKey(item.regionName) !== normalizeKey(region)
      ) {
        return false;
      }

      if (
        area !== "ALL" &&
        normalizeKey(item.areaName) !== normalizeKey(area)
      ) {
        return false;
      }

      if (query) {
        const fields = [
          item.regionName,
          item.areaName,
          item.programName,
          item.orderCode,
          item.slotOwner,
          ...(item.npps || [])
        ];

        if (
          !fields.some(
            value =>
              normalizeKey(value).includes(query)
          )
        ) {
          return false;
        }
      }

      return true;
    });
  }, [
    report,
    selectedUsedScope,
    region,
    area,
    query
  ]);


  // ----------------------------------------------
  // CẢNH BÁO
  // ----------------------------------------------

  const warnings = useMemo(
    () =>
      (report?.warnings || []).filter(item => {
        if (
          region !== "ALL" &&
          item.regionName &&
          normalizeKey(item.regionName) !== normalizeKey(region)
        ) {
          return false;
        }

        if (
          area !== "ALL" &&
          normalizeKey(item.areaName) !== normalizeKey(area)
        ) {
          return false;
        }

        if (query) {
          const fields = [
            item.regionName,
            item.areaName,
            item.programName,
            item.npp,
            item.orderCode,
            item.reason
          ];

          if (
            !fields.some(
              value =>
                normalizeKey(value).includes(query)
            )
          ) {
            return false;
          }
        }

        return true;
      }),
    [report, region, area, query]
  );


  // ----------------------------------------------
  // THAY THÁNG
  // ----------------------------------------------

  function changeMonth(value: string) {
    if (!value || value === month) return;

    // Không giữ các file tháng trước để tránh
    // vô tình ghi nhầm sang tháng mới.
    setAllocation(null);
    setProgramMap(null);
    setActual(null);

    setAllocationName("");
    setProgramMapName("");
    setActualName("");

    setExpandedZones([]);
    setSelectedUsedScope(null);

    setSavedAt("");
    setStatus("");

    setMonth(value);
  }


  // ----------------------------------------------
  // UPLOAD MỘT FILE
  // ----------------------------------------------

  async function loadFile(
    type: "allocation" | "programMap" | "actual",
    file?: File
  ): Promise<void> {
    if (!file) return;

    setBusy(true);
    setStatus("");
    setSavedAt("");

    try {
      if (type === "allocation") {
        const data = await parseAllocation(file);

        setAllocation(data);
        setAllocationName(file.name);

        setStatus(
          `Đã đọc ${data.length} dòng phân bổ.`
        );

      } else if (type === "programMap") {
        const data = await parseProgramMapping(file);

        setProgramMap(data);
        setProgramMapName(file.name);

        setStatus(
          `Đã đọc ${data.length} dòng quy ước Tên CT.`
        );

      } else {
        const data = await parseActual(file);

        setActual(data);
        setActualName(file.name);

        setStatus(
          "Đã đọc file 5.1 từ các cột E, G, N, W, AK."
        );
      }

      // Sau khi đổi một nguồn dữ liệu, phải xử lý lại.
      setReport(null);
      setStatusType("success");

    } catch (error: any) {
      setStatus(
        error?.message ||
        "Không đọc được file."
      );

      setStatusType("error");

    } finally {
      setBusy(false);
    }
  }


  // ----------------------------------------------
  // LƯU DỮ LIỆU CỦA THÁNG
  // ----------------------------------------------

  async function processData(): Promise<void> {
    if (!allocation || !programMap || !actual) {
      setStatus(
        "Vui lòng upload đủ 3 file: Phân bổ, Tên CT và 5.1."
      );

      setStatusType("error");
      return;
    }

    setBusy(true);
    setStatus("");
    setSavedAt("");

    try {
      const result = buildReport(
        allocation,
        programMap,
        actual
      );

      // Hiển thị kết quả đã tính.
      setReport(result);

      // Gửi dữ liệu đến API để lưu lịch sử tháng.
      const response = await fetch(
        "/api/save",
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            month,
            allocation,
            programMap,
            report: result
          })
        }
      );

      const serverResult =
        await response.json();

      if (!response.ok || !serverResult.ok) {
        throw new Error(
          serverResult.error ||
          "Không lưu được dữ liệu vào Google Sheets."
        );
      }

      setSavedAt(
        serverResult.updatedAt
          ? new Date(
              serverResult.updatedAt
            ).toLocaleString("vi-VN")
          : new Date().toLocaleString("vi-VN")
      );

      setAvailableMonths(
        uniqueInOrder([
          month,
          ...(serverResult.months || [])
        ]).sort((a, b) => b.localeCompare(a))
      );

      setStatus(
        "Đã xử lý và lưu dữ liệu tháng " +
        month +
        ". Các tháng khác được giữ nguyên."
      );

      setStatusType("success");

    } catch (error: any) {
      setStatus(
        error?.message ||
        "Có lỗi khi xử lý hoặc lưu dữ liệu."
      );

      setStatusType("error");

    } finally {
      setBusy(false);
    }
  }


  // ----------------------------------------------
  // MỞ / ĐÓNG CHI TIẾT
  // ----------------------------------------------

  function toggleZone(key: string) {
    setExpandedZones(previous =>
      previous.includes(key)
        ? previous.filter(item => item !== key)
        : [...previous, key]
    );
  }

  function openUsedDetails(scope: UsedScope) {
    setSelectedUsedScope(scope);
  }

  function closeUsedDetails() {
    setSelectedUsedScope(null);
  }


  // ----------------------------------------------
  // XUẤT EXCEL THEO BỘ LỌC
  // ----------------------------------------------

  function exportExcel() {
    if (!report) return;

    const workbook = XLSX.utils.book_new();

    const trackingRows: unknown[][] = [[
      "Loại dòng",
      "Tên Miền",
      "Tên vùng",
      "Tên CT",
      "Suất phân bổ",
      "Suất đã dùng",
      "% sử dụng",
      "Tổng chiết khấu",
      "Trạng thái"
    ]];

    for (const group of programGroups) {
      trackingRows.push([
        "Tổng chương trình",
        "",
        "",
        group.programName,
        group.allocation,
        group.used,
        group.percent,
        group.discount,
        group.status
      ]);

      for (const item of group.rows) {
        trackingRows.push([
          "Chi tiết vùng",
          item.regionName,
          item.areaName,
          item.programName,
          item.allocation,
          item.used,
          item.percent,
          item.discount,
          item.status
        ]);
      }
    }

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(trackingRows),
      "Tracking"
    );


    const nppRows: unknown[][] = [[
      "Tên Miền",
      "Tên vùng",
      "Tên CT",
      "Tên NPP",
      "Suất đã dùng",
      "Tổng chiết khấu"
    ]];

    for (const item of npps) {
      nppRows.push([
        item.regionName,
        item.areaName,
        item.programName,
        item.npp,
        item.used,
        item.discount
      ]);
    }

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(nppRows),
      "Chi tiết NPP"
    );


    const usedRows: unknown[][] = [[
      "Tên Miền",
      "Tên vùng",
      "Tên CT",
      "Mã đơn hàng",
      "NPP nhận suất",
      "Danh sách NPP",
      "Tổng chiết khấu"
    ]];

    for (const item of report.usedOrders || []) {
      usedRows.push([
        item.regionName,
        item.areaName,
        item.programName,
        item.orderCode,
        item.slotOwner,
        (item.npps || []).join(", "),
        item.discount
      ]);
    }

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(usedRows),
      "Chi tiết suất đã dùng"
    );


    const warningRows: unknown[][] = [[
      "Tên Miền",
      "Tên vùng",
      "Tên CT",
      "Tên NPP",
      "Mã đơn hàng",
      "Tổng AK",
      "Lý do"
    ]];

    for (const item of warnings) {
      warningRows.push([
        item.regionName,
        item.areaName,
        item.programName,
        item.npp,
        item.orderCode,
        item.discount,
        item.reason
      ]);
    }

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(warningRows),
      "Cảnh báo"
    );


    XLSX.writeFile(
      workbook,
      `Tracking_phan_bo_${month}.xlsx`
    );
  }


  // ----------------------------------------------
  // ĐẶT LẠI BỘ LỌC
  // ----------------------------------------------

  function resetFilters() {
    setSelectedProgram("ALL");
    setRegion("ALL");
    setArea("ALL");
    setSearch("");
    setExpandedZones([]);
    setSelectedUsedScope(null);
  }


  // ==================================================
  // GIAO DIỆN
  // ==================================================

  return (
    <main className="wrap">

      {/* TIÊU ĐỀ */}

      <header className="header">
        <div>
          <h1 className="pageTitle">
            Theo dõi phân bổ chương trình
          </h1>

          <p className="headerDescription">
            Theo dõi suất phân bổ, tiến độ sử dụng,
            tiền chiết khấu và NPP.
          </p>

          <div className="headerMeta">
            <span>
              Tháng: <strong>{month}</strong>
            </span>

            {savedAt && (
              <span>
                Cập nhật lần cuối: {savedAt}
              </span>
            )}
          </div>
        </div>
      </header>


      {/* UPLOAD BA FILE */}

      <section className="panel uploadPanel">
        <button
          type="button"
          className="collapseHeader"
          onClick={() => setUploadOpen(!uploadOpen)}
          aria-expanded={uploadOpen}
        >
          <span className="collapseTitle">
            <span className="collapseArrow">
              {uploadOpen ? "▾" : "▸"}
            </span>

            Tải dữ liệu
          </span>

          <span className="collapseHint">
            {uploadOpen ? "Thu gọn" : "Mở rộng"}
          </span>
        </button>


        {uploadOpen && (
          <div className="uploadContent">

            <div className="uploadMonth">
              <label htmlFor="uploadMonth">
                Tháng dữ liệu
              </label>

              <input
                id="uploadMonth"
                type="month"
                value={month}
                disabled={busy || historyLoading}
                onChange={event =>
                  changeMonth(event.target.value)
                }
              />

              <p className="muted">
                Upload lại sẽ thay dữ liệu của tháng đang chọn.
                Các tháng khác được giữ nguyên.
              </p>
            </div>


            <div className="uploadgrid">

              <div className="uploadbox">
                <h3>1. File phân bổ</h3>

                <p className="muted">
                  5 cột: Tên Miền, Tên vùng, Tên CT,
                  Cơ cấu CT, Suất phân bổ.
                </p>

                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  disabled={busy}
                  onChange={event => {
                    const file = event.currentTarget.files?.[0];
                    event.currentTarget.value = "";

                    if (file) {
                      void loadFile("allocation", file);
                    }
                  }}
                />

                {allocationName && (
                  <p className="selectedFile">
                    Đã chọn: {allocationName}
                  </p>
                )}
              </div>


              <div className="uploadbox">
                <h3>2. File Tên CT</h3>

                <p className="muted">
                  2 cột: Mã CT, Tên CT.
                  Một tên có thể chứa nhiều mã CT khác nhau.
                </p>

                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  disabled={busy}
                  onChange={event => {
                    const file = event.currentTarget.files?.[0];
                    event.currentTarget.value = "";

                    if (file) {
                      void loadFile("programMap", file);
                    }
                  }}
                />

                {programMapName && (
                  <p className="selectedFile">
                    Đã chọn: {programMapName}
                  </p>
                )}
              </div>


              <div className="uploadbox">
                <h3>3. File 5.1</h3>

                <p className="muted">
                  Chỉ đọc các cột E, G, N, W và AK.
                </p>

                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  disabled={busy}
                  onChange={event => {
                    const file = event.currentTarget.files?.[0];
                    event.currentTarget.value = "";

                    if (file) {
                      void loadFile("actual", file);
                    }
                  }}
                />

                {actualName && (
                  <p className="selectedFile">
                    Đã chọn: {actualName}
                  </p>
                )}
              </div>

            </div>


            <div className="actions uploadActions">
              <button
                type="button"
                className="btn"
                disabled={
                  busy ||
                  !allocation ||
                  !programMap ||
                  !actual
                }
                onClick={processData}
              >
                {busy
                  ? "Đang xử lý..."
                  : "Kiểm tra và lưu dữ liệu"}
              </button>
            </div>


            {status && (
              <div className={`status ${statusType}`}>
                {status}
              </div>
            )}


            <div className="notice">
              <strong>Quy tắc xử lý:</strong>

              <ul>
                <li>
                  Mã CT ở cột N được lấy phần đầu trước dấu phẩy.
                  Mã đó được tra cứu trong file Tên CT.
                </li>

                <li>
                  Các mã CT cùng Tên CT được tổng hợp chung.
                </li>

                <li>
                  Một Tên CT + Vùng + Mã đơn hàng chỉ tính
                  tối đa một suất nếu có ít nhất một dòng AK &gt; 0.
                </li>

                <li>
                  Tổng tiền chiết khấu cộng tất cả các dòng AK hợp lệ.
                </li>

                <li>
                  Dữ liệu không có mapping hoặc không có phân bổ
                  được đưa vào cảnh báo, không cộng KPI.
                </li>
              </ul>
            </div>
          </div>
        )}
      </section>


      {/* BỘ LỌC */}

      <section className="panel filterPanel">
        <div className="filterBar">

          <label className="filterField">
            <span>Tháng</span>

            <input
              type="month"
              value={month}
              disabled={historyLoading || busy}
              onChange={event =>
                changeMonth(event.target.value)
              }
            />
          </label>


          <label className="filterField">
            <span>Miền</span>

            <select
              value={region}
              onChange={event =>
                setRegion(event.target.value)
              }
            >
              <option value="ALL">Tất cả miền</option>

              {regions.map(item => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </select>
          </label>


          <label className="filterField">
            <span>Vùng</span>

            <select
              value={area}
              onChange={event =>
                setArea(event.target.value)
              }
            >
              <option value="ALL">Tất cả vùng</option>

              {areas.map(item => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </select>
          </label>


          <label className="searchField">
            <span>Tìm kiếm</span>

            <input
              type="search"
              value={search}
              onChange={event =>
                setSearch(event.target.value)
              }
              placeholder="Miền, vùng, NPP, Tên CT..."
            />
          </label>


          <button
            type="button"
            className="btn secondary exportButton"
            disabled={!report || historyLoading}
            onClick={exportExcel}
          >
            ↓ Xuất Excel
          </button>

        </div>


        <div className="historyBar">
          <span className="historyLabel">
            Lịch sử tháng:
          </span>

          {availableMonths.length === 0 && (
            <span className="muted">
              Chưa có dữ liệu lịch sử
            </span>
          )}

          {availableMonths.map(item => (
            <button
              type="button"
              key={item}
              className={
                item === month
                  ? "monthChip active"
                  : "monthChip"
              }
              onClick={() => changeMonth(item)}
              disabled={historyLoading || busy}
            >
              {item}
            </button>
          ))}

          {historyLoading && (
            <span className="muted">
              Đang tải lịch sử...
            </span>
          )}
        </div>
      </section>


      {/* TAB TÊN CT */}

      {report && (
        <section className="panel programPanel">
          <div className="sectionHeadingRow">
            <div>
              <h2 className="sectionTitle">
                Danh mục chương trình
              </h2>

              <p className="muted">
                Các mã CT đã được quy về Tên CT.
                Dashboard không hiển thị mã CT.
              </p>
            </div>

            <button
              type="button"
              className="btn secondary"
              onClick={resetFilters}
            >
              Xóa bộ lọc
            </button>
          </div>


          <div className="programTabs">
            <button
              type="button"
              className={
                selectedProgram === "ALL"
                  ? "programTab active"
                  : "programTab"
              }
              onClick={() => setSelectedProgram("ALL")}
            >
              Tất cả
            </button>

            {programs.map(item => (
              <button
                type="button"
                key={item}
                className={
                  normalizeKey(selectedProgram) ===
                  normalizeKey(item)
                    ? "programTab active"
                    : "programTab"
                }
                onClick={() => setSelectedProgram(item)}
                title={item}
              >
                {item}
              </button>
            ))}
          </div>
        </section>
      )}


      {/* KPI */}

      {report && (
        <section className="kpiGrid">

          <div className="metric">
            <div className="metricLabel">
              Tổng suất phân bổ
            </div>

            <div className="metricValue">
              {fmt(total.allocation)}
            </div>

            <div className="metricHint">
              Theo bộ lọc đang chọn
            </div>
          </div>


          <div className="metric">
            <div className="metricLabel">
              Tổng suất đã dùng
            </div>

            <div className="metricValue greenValue">
              {fmt(total.used)}
            </div>

            <div className="metricHint">
              Bấm số trong bảng để tra cứu đơn hàng
            </div>
          </div>


          <div className="metric">
            <div className="metricLabel">
              Tiến độ sử dụng
            </div>

            <div className="metricValue">
              {fmt(totalPercent)}%
            </div>

            <div className="mainProgress">
              <div
                className={
                  "mainProgressFill " +
                  (
                    totalPercent >= 100
                      ? "progressRed"
                      : totalPercent >= 80
                      ? "progressYellow"
                      : ""
                  )
                }
                style={{
                  width: `${Math.min(
                    100,
                    Math.max(0, totalPercent)
                  )}%`
                }}
              />
            </div>

            <div className="metricHint">
              {getStatus(total.allocation, total.used)}
            </div>
          </div>


          <div className="metric">
            <div className="metricLabel">
              Tổng tiền chiết khấu
            </div>

            <div className="moneyValue">
              {money(total.discount)}
            </div>

            <div className="metricHint">
              Tổng AK thuộc dữ liệu hợp lệ
            </div>
          </div>

        </section>
      )}


      {/* BẢNG TRACKING */}

      {report && (
        <section className="panel trackingPanel">

          <div className="sectionHeadingRow">
            <div>
              <h2 className="sectionTitle">
                Tiến độ sử dụng theo chương trình
              </h2>

              <p className="muted">
                Dòng đậm là tổng Tên CT; dòng dưới là từng vùng.
                Bấm số đã dùng để tra cứu đơn hàng và NPP.
              </p>
            </div>

            <div className="resultCount">
              {filtered.length} dòng vùng
            </div>
          </div>


          <div className="tablewrap">
            <table className="trackingTable">

              <thead>
                <tr>
                  <th>Miền</th>
                  <th>Vùng / chương trình</th>
                  <th>Tên CT</th>
                  <th className="num">Phân bổ</th>
                  <th className="num">Đã dùng</th>
                  <th className="progressColumn">Tiến độ</th>
                  <th className="num">Tổng chiết khấu</th>
                  <th>Trạng thái</th>
                </tr>
              </thead>


              <tbody>
                {programGroups.map(group => (
                  <FragmentProgramGroup
                    key={normalizeKey(group.programName)}
                    group={group}
                    expandedZones={expandedZones}
                    toggleZone={toggleZone}
                    fmt={fmt}
                    money={money}

                    onClickProgramUsed={() => {
                      openUsedDetails({
                        programName: group.programName
                      });
                    }}

                    onClickZoneUsed={(item: Summary) => {
                      openUsedDetails({
                        programName: item.programName,
                        regionName: item.regionName,
                        areaName: item.areaName
                      });
                    }}

                    getNppsForZone={(item: Summary) => {
                      const key = detailKey(
                        item.regionName,
                        item.areaName,
                        item.programName
                      );

                      return npps.filter(
                        npp =>
                          detailKey(
                            npp.regionName,
                            npp.areaName,
                            npp.programName
                          ) === key
                      );
                    }}
                  />
                ))}


                {programGroups.length === 0 && (
                  <tr>
                    <td colSpan={8} className="noData">
                      Không có dữ liệu phù hợp với bộ lọc.
                    </td>
                  </tr>
                )}
              </tbody>


              {filtered.length > 0 && (
                <tfoot>
                  <tr>
                    <td colSpan={3}>
                      Tổng cộng
                    </td>

                    <td className="num">
                      {fmt(total.allocation)}
                    </td>

                    <td className="num">
                      {fmt(total.used)}
                    </td>

                    <td>
                      <strong>
                        {fmt(totalPercent)}%
                      </strong>
                    </td>

                    <td className="num">
                      {money(total.discount)}
                    </td>

                    <td>
                      <span
                        className={
                          "statusBadge " +
                          statusClass(
                            getStatus(
                              total.allocation,
                              total.used
                            )
                          )
                        }
                      >
                        {getStatus(
                          total.allocation,
                          total.used
                        )}
                      </span>
                    </td>
                  </tr>
                </tfoot>
              )}

            </table>
          </div>
        </section>
      )}


      {/* CẢNH BÁO */}

      {report && (
        <section className="panel">

          <div className="sectionHeadingRow">
            <div>
              <h2 className="sectionTitle">
                Cảnh báo dữ liệu
              </h2>

              <p className="muted">
                Bao gồm mã CT chưa có trong file Tên CT,
                vùng không có phân bổ và đơn hàng thiếu mã.
              </p>
            </div>

            <span className="warningCount">
              {warnings.length} cảnh báo
            </span>
          </div>


          <div className="tablewrap">
            <table>
              <thead>
                <tr>
                  <th>Miền</th>
                  <th>Vùng</th>
                  <th>Tên CT</th>
                  <th>Mã đơn hàng</th>
                  <th>NPP</th>
                  <th className="num">Tổng AK</th>
                  <th>Lý do</th>
                </tr>
              </thead>

              <tbody>
                {warnings.map((item, index) => (
                  <tr
                    key={
                      item.areaName +
                      item.programName +
                      item.orderCode +
                      index
                    }
                  >
                    <td>
                      {item.regionName || "Chưa xác định"}
                    </td>

                    <td>{item.areaName}</td>

                    <td>{item.programName}</td>

                    <td>
                      {item.orderCode || "(Trống)"}
                    </td>

                    <td>
                      {item.npp || "(Trống)"}
                    </td>

                    <td className="num">
                      {money(item.discount)}
                    </td>

                    <td>{item.reason}</td>
                  </tr>
                ))}

                {warnings.length === 0 && (
                  <tr>
                    <td colSpan={7} className="noData">
                      Không có dữ liệu cảnh báo.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      )}


      {/* CỬA SỔ TRA CỨU SUẤT ĐÃ DÙNG */}

      {selectedUsedScope && (
        <div
          className="modalBackdrop"
          onMouseDown={event => {
            if (event.target === event.currentTarget) {
              setSelectedUsedScope(null);
            }
          }}
        >
          <section
            className="usedModal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="usedModalTitle"
          >

            <div className="usedModalHeader">
              <div>
                <h2
                  className="sectionTitle"
                  id="usedModalTitle"
                >
                  Chi tiết suất đã dùng
                </h2>

                <p className="muted">
                  {selectedUsedScope.programName}

                  {selectedUsedScope.areaName
                    ? " · " + selectedUsedScope.areaName
                    : ""}
                </p>
              </div>

              <button
                type="button"
                className="modalCloseButton"
                onClick={closeUsedDetails}
                aria-label="Đóng cửa sổ"
              >
                ×
              </button>
            </div>


            <div className="usedModalSummary">
              <div>
                <span>Số đơn hàng đã dùng suất</span>

                <strong>
                  {activeUsedOrders.length}
                </strong>
              </div>

              <div>
                <span>Tổng chiết khấu các đơn hàng</span>

                <strong>
                  {money(
                    activeUsedOrders.reduce(
                      (sum, item) => sum + item.discount,
                      0
                    )
                  )}
                </strong>
              </div>
            </div>


            <p className="muted modalExplanation">
              Mỗi dòng là một mã đơn hàng được tính một suất.
              Khi một đơn có nhiều NPP, cột NPP nhận suất
              cho biết đơn vị được gán suất; tổng chiết khấu
              vẫn cộng theo toàn bộ dòng AK của đơn hàng.
            </p>


            <div className="tablewrap modalTableWrap">
              <table>
                <thead>
                  <tr>
                    <th>Miền</th>
                    <th>Vùng</th>
                    <th>Mã đơn hàng</th>
                    <th>NPP nhận suất</th>
                    <th>Các NPP trên đơn</th>
                    <th className="num">Tổng chiết khấu</th>
                  </tr>
                </thead>

                <tbody>
                  {activeUsedOrders.map((item, index) => (
                    <tr
                      key={
                        item.areaName +
                        item.programName +
                        item.orderCode +
                        index
                      }
                    >
                      <td>{item.regionName}</td>
                      <td>{item.areaName}</td>
                      <td>{item.orderCode}</td>
                      <td>{item.slotOwner}</td>
                      <td>
                        {(item.npps || []).join(", ")}
                      </td>
                      <td className="num">
                        {money(item.discount)}
                      </td>
                    </tr>
                  ))}

                  {activeUsedOrders.length === 0 && (
                    <tr>
                      <td colSpan={6} className="noData">
                        Không tìm thấy chi tiết đơn hàng.
                        Nếu tháng này được lưu trước khi có
                        chức năng tra cứu, hãy upload lại
                        dữ liệu tháng đó để lưu chi tiết.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>


            <div className="usedModalFooter">
              <button
                type="button"
                className="btn secondary"
                onClick={closeUsedDetails}
              >
                Đóng
              </button>
            </div>
          </section>
        </div>
      )}


      {/* CHƯA CÓ DỮ LIỆU */}

      {!report && !historyLoading && (
        <section className="panel emptyState">
          <h2>
            Chưa có dữ liệu tháng {month}
          </h2>

          <p className="muted">
            Mở khu vực Tải dữ liệu và upload đủ ba file.
          </p>
        </section>
      )}


      <footer className="footer">
        Dashboard theo dõi phân bổ CTKM / CTTL.
        Kiểm tra cảnh báo và đối chiếu số liệu trước khi
        sử dụng làm báo cáo chính thức.
      </footer>

    </main>
  );
}


// ==================================================
// DÒNG TỔNG CHƯƠNG TRÌNH
// ==================================================

function FragmentProgramGroup(props: {
  group: ProgramGroup;
  expandedZones: string[];
  toggleZone: (key: string) => void;
  fmt: (value: number) => string;
  money: (value: number) => string;
  onClickProgramUsed: () => void;
  onClickZoneUsed: (item: Summary) => void;
  getNppsForZone: (item: Summary) => NppSummary[];
}) {
  const group = props.group;

  const groupWidth = Math.max(
    0,
    Math.min(100, group.percent)
  );

  return (
    <>
      <tr className="programTotalRow">
        <td colSpan={3}>
          <div className="programTotalName">
            {group.programName}
          </div>

          <div className="programTotalSub">
            Tổng chương trình
          </div>
        </td>

        <td className="num">
          {props.fmt(group.allocation)}
        </td>

        <td className="num">
          <button
            type="button"
            className="usedNumberButton"
            onClick={props.onClickProgramUsed}
            disabled={group.used === 0}
            title="Bấm để tra cứu các đơn hàng đã dùng suất"
          >
            {props.fmt(group.used)}
          </button>
        </td>

        <td className="progressCell">
          <div className="progressInfo">
            <strong>{props.fmt(group.percent)}%</strong>
          </div>

          <div className="rowProgress">
            <div
              className={
                "rowProgressFill " +
                (
                  group.percent >= 100
                    ? "progressRed"
                    : group.percent >= 80
                    ? "progressYellow"
                    : ""
                )
              }
              style={{
                width: `${groupWidth}%`
              }}
            />
          </div>
        </td>

        <td className="num discountValue">
          {props.money(group.discount)}
        </td>

        <td>
          <span
            className={
              "statusBadge " +
              statusClass(group.status)
            }
          >
            {group.status}
          </span>
        </td>
      </tr>


      {group.rows.map((item, index) => {
        const key = detailKey(
          item.regionName,
          item.areaName,
          item.programName
        );

        const expanded =
          props.expandedZones.includes(key);

        const npps = props.getNppsForZone(item);

        return (
          <FragmentZoneRow
            key={key + index}
            item={item}
            expanded={expanded}
            npps={npps}
            toggle={() => props.toggleZone(key)}
            onClickUsed={() =>
              props.onClickZoneUsed(item)
            }
            fmt={props.fmt}
            money={props.money}
          />
        );
      })}
    </>
  );
}


// ==================================================
// DÒNG VÙNG VÀ CHI TIẾT NPP
// ==================================================

function FragmentZoneRow(props: {
  item: Summary;
  expanded: boolean;
  npps: NppSummary[];
  toggle: () => void;
  onClickUsed: () => void;
  fmt: (value: number) => string;
  money: (value: number) => string;
}) {
  const item = props.item;

  const percentWidth = Math.max(
    0,
    Math.min(100, item.percent)
  );

  return (
    <>
      <tr className="zoneRow">
        <td>{item.regionName}</td>

        <td>
          <button
            type="button"
            className="zoneExpandButton"
            onClick={props.toggle}
            aria-expanded={props.expanded}
          >
            <span className="expandIcon">
              {props.expanded ? "▾" : "▸"}
            </span>

            {item.areaName}
          </button>
        </td>

        <td>{item.programName}</td>

        <td className="num">
          {props.fmt(item.allocation)}
        </td>

        <td className="num">
          <button
            type="button"
            className="usedNumberButton"
            onClick={props.onClickUsed}
            disabled={item.used === 0}
            title="Bấm để xem các đơn hàng và NPP tạo ra suất"
          >
            {props.fmt(item.used)}
          </button>
        </td>

        <td className="progressCell">
          <div className="progressInfo">
            <span>{props.fmt(item.percent)}%</span>
          </div>

          <div className="rowProgress">
            <div
              className={
                "rowProgressFill " +
                (
                  item.percent >= 100
                    ? "progressRed"
                    : item.percent >= 80
                    ? "progressYellow"
                    : ""
                )
              }
              style={{
                width: `${percentWidth}%`
              }}
            />
          </div>
        </td>

        <td className="num discountValue">
          {props.money(item.discount)}
        </td>

        <td>
          <span
            className={
              "statusBadge " +
              statusClass(item.status)
            }
          >
            {item.status}
          </span>
        </td>
      </tr>


      {props.expanded && (
        <tr className="nppExpansionRow">
          <td colSpan={8}>
            <div className="nppDetail">
              <div className="nppDetailTitle">
                Chi tiết NPP — {item.areaName}
              </div>

              {props.npps.length > 0 ? (
                <table className="nppTable">
                  <thead>
                    <tr>
                      <th>Tên NPP</th>
                      <th className="num">
                        Suất đã dùng
                      </th>
                      <th className="num">
                        Tổng chiết khấu
                      </th>
                    </tr>
                  </thead>

                  <tbody>
                    {props.npps.map((npp, index) => (
                      <tr
                        key={
                          normalizeKey(npp.npp) +
                          index
                        }
                      >
                        <td>{npp.npp}</td>

                        <td className="num">
                          {props.fmt(npp.used)}
                        </td>

                        <td className="num">
                          {props.money(npp.discount)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <p className="muted">
                  Không có dữ liệu NPP phù hợp.
                </p>
              )}
            </div>
          </td>
        </tr>
      )}
    </>
  );
}
