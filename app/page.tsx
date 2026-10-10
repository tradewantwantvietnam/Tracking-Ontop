"use client";

import {
  useEffect,
  useMemo,
  useState
} from "react";

import * as XLSX from "xlsx";

import {
  parseActual,
  parseAllocation,
  type Actual,
  type Allocation
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

type Report = {
  summaries: Summary[];
  npps: NppSummary[];
  warnings: Warning[];
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


// ==================================================
// HÀM HỖ TRỢ
// ==================================================

function currentMonth(): string {
  const date = new Date();

  return (
    date.getFullYear() +
    "-" +
    String(date.getMonth() + 1).padStart(2, "0")
  );
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

function cleanText(value: unknown): string {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeKey(value: string): string {
  return cleanText(value).toLocaleLowerCase("vi");
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
// TÍNH TRẠNG THÁI
// ==================================================

function getStatus(
  allocation: number,
  used: number
): StatusName {
  const percent = allocation > 0
    ? (used / allocation) * 100
    : (used > 0 ? 100 : 0);

  if (percent >= 100) {
    return "Vượt định biên";
  }

  if (percent >= 80) {
    return "Sắp đầy";
  }

  return "Bình thường";
}

function getPercent(
  allocation: number,
  used: number
): number {
  if (allocation > 0) {
    return (used / allocation) * 100;
  }

  return used > 0 ? 100 : 0;
}


// ==================================================
// TẠO BÁO CÁO
// ==================================================

function buildReport(
  allocations: Allocation[],
  actuals: Actual[]
): Report {

  // ----------------------------------------------
  // A. MAPPING VÙNG + MÃ CT
  // ----------------------------------------------

  const allocationByCode: Map<string, Allocation> =
    new Map();

  function codeKey(
    areaName: string,
    programCode: string
  ): string {
    return (
      normalizeKey(areaName) +
      "|||" +
      normalizeKey(programCode)
    );
  }

  for (const item of allocations) {
    const key = codeKey(
      item.areaName,
      item.programCode
    );

    if (allocationByCode.has(key)) {
      throw new Error(
        "File phân bổ có Tên vùng + Mã CT bị trùng: " +
        item.areaName +
        " / " +
        item.programCode +
        ". Hãy kiểm tra file phân bổ."
      );
    }

    allocationByCode.set(key, item);
  }


  // ----------------------------------------------
  // B. TỔNG PHÂN BỔ
  //
  // Nhóm theo Miền + Vùng + Tên CT.
  // Các mã CT có cùng Tên CT được cộng chung.
  // Thứ tự giữ theo file phân bổ.
  // ----------------------------------------------

  const summaryMap: Map<string, Summary> =
    new Map();

  for (const item of allocations) {
    const name =
      cleanText(item.programName) ||
      "Chưa đặt tên CT";

    const key = detailKey(
      item.regionName,
      item.areaName,
      name
    );

    if (!summaryMap.has(key)) {
      summaryMap.set(key, {
        regionName: item.regionName,
        areaName: item.areaName,
        programName: name,
        allocation: 0,
        used: 0,
        discount: 0,
        percent: 0,
        status: "Bình thường"
      });
    }

    summaryMap.get(key)!.allocation +=
      Number(item.allocation) || 0;
  }


  // ----------------------------------------------
  // C. NHÓM ĐƠN HÀNG
  //
  // Khử trùng theo:
  // Tên CT + Tên vùng + Mã đơn hàng.
  //
  // Không dùng mã CT làm khóa khử trùng.
  // ----------------------------------------------

  type OrderRecord = {
    row: Actual;
    allocation: Allocation;
  };

  type OrderGroup = {
    records: OrderRecord[];
    discount: number;
    hasPositive: boolean;
    ownerKey: string;
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
    const allocation = allocationByCode.get(
      codeKey(row.areaName, row.programCode)
    );


    // --------------------------------------------
    // F1. KHÔNG TÌM THẤY PHÂN BỔ
    // --------------------------------------------

    if (!allocation) {
      const orderCode = cleanText(row.orderCode);

      const warningKey = orderCode
        ? (
            normalizeKey(row.areaName) +
            "|||" +
            normalizeKey(row.programCode) +
            "|||" +
            normalizeKey(orderCode)
          )
        : "MISSING_ORDER_" + rowIndex;

      if (!unmatchedWarnings.has(warningKey)) {
        unmatchedWarnings.set(warningKey, {
          regionName: "",
          areaName: row.areaName,
          programName: "Chưa xác định",
          npp: row.npp,
          orderCode,
          discount: 0,
          reason:
            "Không tìm thấy Tên vùng + Mã CT trong file phân bổ"
        });
      }

      unmatchedWarnings.get(warningKey)!.discount +=
        Number(row.discount) || 0;

      return;
    }


    const programName =
      cleanText(allocation.programName) ||
      "Chưa đặt tên CT";

    const parentKey = detailKey(
      allocation.regionName,
      allocation.areaName,
      programName
    );

    const summary = summaryMap.get(parentKey)!;


    // --------------------------------------------
    // F2. THIẾU MÃ ĐƠN HÀNG
    // --------------------------------------------

    if (!cleanText(row.orderCode)) {
      const discount = Number(row.discount) || 0;

      // Vẫn cộng tiền AK, nhưng không tính suất.
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
    // F3. KHÓA ĐƠN HÀNG
    // Tên CT + Tên vùng + Mã đơn hàng
    // --------------------------------------------

    const orderKey = [
      normalizeKey(programName),
      normalizeKey(allocation.areaName),
      normalizeKey(row.orderCode)
    ].join("|||");

    if (!orderGroups.has(orderKey)) {
      orderGroups.set(orderKey, {
        records: [],
        discount: 0,
        hasPositive: false,
        ownerKey: parentKey
      });
    }

    const group = orderGroups.get(orderKey)!;

    group.records.push({
      row,
      allocation
    });

    group.discount += Number(row.discount) || 0;

    if (Number(row.discount) > 0) {
      group.hasPositive = true;
    }
  });


  // ----------------------------------------------
  // G. TỔNG HỢP ĐƠN HÀNG
  // ----------------------------------------------

  for (const group of orderGroups.values()) {
    const firstRecord = group.records[0];

    const allocation = firstRecord.allocation;

    const programName =
      cleanText(allocation.programName) ||
      "Chưa đặt tên CT";

    const summaryKey = detailKey(
      allocation.regionName,
      allocation.areaName,
      programName
    );

    const summary = summaryMap.get(summaryKey)!;

    // Cộng tiền tổng của đơn hàng.
    summary.discount += group.discount;

    // Chỉ cần có ít nhất một dòng AK > 0:
    // đơn hàng sử dụng đúng 1 suất.
    if (group.hasPositive) {
      summary.used += 1;
    }


    // --------------------------------------------
    // G1. GOM TIỀN THEO NPP
    // --------------------------------------------

    const nppDiscounts: Map<
      string,
      {
        nppName: string;
        discount: number;
        ownerKey: string;
      }
    > = new Map();

    for (const record of group.records) {
      const row = record.row;
      const recordAllocation = record.allocation;

      const nppName =
        cleanText(row.npp) ||
        "(Không có tên NPP)";

      const nppKey = [
        detailKey(
          recordAllocation.regionName,
          recordAllocation.areaName,
          cleanText(recordAllocation.programName) ||
            "Chưa đặt tên CT"
        ),
        normalizeKey(nppName)
      ].join("|||");

      if (!nppDiscounts.has(nppKey)) {
        nppDiscounts.set(nppKey, {
          nppName,
          discount: 0,
          ownerKey: detailKey(
            recordAllocation.regionName,
            recordAllocation.areaName,
            cleanText(recordAllocation.programName) ||
              "Chưa đặt tên CT"
          )
        });
      }

      nppDiscounts.get(nppKey)!.discount +=
        Number(row.discount) || 0;
    }


    // --------------------------------------------
    // G2. GÁN MỘT SUẤT CHO NPP CHỦ ĐƠN HÀNG
    //
    // NPP đầu tiên có AK > 0 nhận suất.
    // Tổng tiền vẫn phân bổ theo từng dòng AK.
    // --------------------------------------------

    const positiveRecord = group.records.find(
      record => Number(record.row.discount) > 0
    );

    const ownerNpp =
      positiveRecord
        ? (
            cleanText(positiveRecord.row.npp) ||
            "(Không có tên NPP)"
          )
        : "";

    const ownerNppKey = normalizeKey(ownerNpp);

    for (const item of nppDiscounts.values()) {
      const usedToAdd =
        group.hasPositive &&
        normalizeKey(item.nppName) === ownerNppKey &&
        item.ownerKey === group.ownerKey
          ? 1
          : 0;

      const nameParts = item.ownerKey.split("|||");

      // Không phân bổ định mức cho NPP.
      // Chỉ lưu số suất đã dùng và tiền chiết khấu.
      addNpp(
        group.records.find(record =>
          detailKey(
            record.allocation.regionName,
            record.allocation.areaName,
            cleanText(record.allocation.programName) ||
              "Chưa đặt tên CT"
          ) === item.ownerKey
        )?.allocation.regionName || "",
        group.records.find(record =>
          detailKey(
            record.allocation.regionName,
            record.allocation.areaName,
            cleanText(record.allocation.programName) ||
              "Chưa đặt tên CT"
          ) === item.ownerKey
        )?.allocation.areaName || "",
        group.records.find(record =>
          detailKey(
            record.allocation.regionName,
            record.allocation.areaName,
            cleanText(record.allocation.programName) ||
              "Chưa đặt tên CT"
          ) === item.ownerKey
        )?.allocation.programName || "",
        item.nppName,
        item.discount,
        usedToAdd
      );
    }
  }


  // ----------------------------------------------
  // H. HOÀN THIỆN TỶ LỆ SỬ DỤNG VÀ TRẠNG THÁI
  // ----------------------------------------------

  const summaries: Summary[] =
    [...summaryMap.values()].map(item => {
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
    });


  // ----------------------------------------------
  // I. TRẢ BÁO CÁO
  // ----------------------------------------------

  return {
    summaries,
    npps: [...nppMap.values()],
    warnings: [
      ...unmatchedWarnings.values(),
      ...missingOrderWarnings
    ]
  };
}


// ==================================================
// COMPONENT DASHBOARD
// ==================================================

export default function Home() {

  const [month, setMonth] = useState(currentMonth());

  const [availableMonths, setAvailableMonths] =
    useState<string[]>([]);

  const [allocation, setAllocation] =
    useState<Allocation[] | null>(null);

  const [actual, setActual] =
    useState<Actual[] | null>(null);

  const [report, setReport] =
    useState<Report | null>(null);

  const [allocationName, setAllocationName] =
    useState("");

  const [actualName, setActualName] =
    useState("");

  const [busy, setBusy] = useState(false);

  const [historyLoading, setHistoryLoading] =
    useState(false);

  const [uploadOpen, setUploadOpen] =
    useState(true);

  const [status, setStatus] = useState("");

  const [statusType, setStatusType] =
    useState<"success" | "error" | "info">("info");

  const [savedAt, setSavedAt] = useState("");

  const [program, setProgram] = useState("ALL");

  const [region, setRegion] = useState("ALL");

  const [area, setArea] = useState("ALL");

  const [search, setSearch] = useState("");

  const [expandedZones, setExpandedZones] =
    useState<string[]>([]);


  // ----------------------------------------------
  // TẢI LỊCH SỬ KHI CHỌN THÁNG
  // ----------------------------------------------

  useEffect(() => {
    let cancelled = false;

    async function loadMonth() {
      setHistoryLoading(true);
      setReport(null);
      setSavedAt("");
      setStatus("");

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
            "Không tải được lịch sử của tháng."
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
          setReport(reportData.report);

          setSavedAt(
            reportData.updatedAt
              ? new Date(
                  reportData.updatedAt
                ).toLocaleString("vi-VN")
              : ""
          );

          setStatus(
            "Đã tải dữ liệu lịch sử tháng " + month + "."
          );

          setStatusType("success");

        } else {
          setReport(null);

          setStatus(
            "Tháng " + month +
            " chưa có dữ liệu đã lưu. " +
            "Bạn có thể upload hai file để tạo dữ liệu tháng này."
          );

          setStatusType("info");
        }

      } catch (error: any) {
        if (cancelled) return;

        setReport(null);

        setStatus(
          error?.message ||
          "Không tải được lịch sử. Hãy kiểm tra cấu hình API."
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
  // DANH SÁCH TÊN CT
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


  // ----------------------------------------------
  // BỘ LỌC MIỀN
  // ----------------------------------------------

  const regions = useMemo(
    () =>
      uniqueInOrder(
        (report?.summaries || [])
          .filter(
            item =>
              program === "ALL" ||
              normalizeKey(item.programName) ===
                normalizeKey(program)
          )
          .map(item => item.regionName)
      ),
    [report, program]
  );


  // ----------------------------------------------
  // BỘ LỌC VÙNG
  // ----------------------------------------------

  const areas = useMemo(
    () =>
      uniqueInOrder(
        (report?.summaries || [])
          .filter(
            item =>
              (
                program === "ALL" ||
                normalizeKey(item.programName) ===
                  normalizeKey(program)
              ) &&
              (
                region === "ALL" ||
                normalizeKey(item.regionName) ===
                  normalizeKey(region)
              )
          )
          .map(item => item.areaName)
      ),
    [report, program, region]
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
  // DỮ LIỆU GỐC THEO BỘ LỌC
  // ----------------------------------------------

  const baseFiltered = useMemo(
    () =>
      (report?.summaries || []).filter(
        item =>
          (
            program === "ALL" ||
            normalizeKey(item.programName) ===
              normalizeKey(program)
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
    [report, program, region, area]
  );


  const baseNpps = useMemo(
    () =>
      (report?.npps || []).filter(
        item =>
          (
            program === "ALL" ||
            normalizeKey(item.programName) ===
              normalizeKey(program)
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
    [report, program, region, area]
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
    ].some(value =>
      normalizeKey(value).includes(query)
    );
  }


  const nppMatchedParents = useMemo(() => {
    const result = new Set<string>();

    if (!query) return result;

    for (const item of baseNpps) {
      if (
        normalizeKey(item.npp).includes(query)
      ) {
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
          nppMatchedParents.has(
            detailKey(
              item.regionName,
              item.areaName,
              item.programName
            )
          )
        );
      }),
    [baseFiltered, query, nppMatchedParents]
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


  const warnings = useMemo(
    () =>
      (report?.warnings || []).filter(item => {
        if (
          region !== "ALL" &&
          item.regionName &&
          normalizeKey(item.regionName) !==
            normalizeKey(region)
        ) {
          return false;
        }

        if (
          area !== "ALL" &&
          normalizeKey(item.areaName) !==
            normalizeKey(area)
        ) {
          return false;
        }

        if (
          program !== "ALL" &&
          item.programName !== "Chưa xác định" &&
          normalizeKey(item.programName) !==
            normalizeKey(program)
        ) {
          return false;
        }

        if (!query) return true;

        return [
          item.regionName,
          item.areaName,
          item.programName,
          item.npp,
          item.orderCode,
          item.reason
        ].some(value =>
          normalizeKey(value).includes(query)
        );
      }),
    [report, region, area, program, query]
  );


  // ----------------------------------------------
  // KPI
  // ----------------------------------------------

  const total = useMemo(
    () =>
      filtered.reduce(
        (sum, item) => ({
          allocation:
            sum.allocation + item.allocation,
          used:
            sum.used + item.used,
          discount:
            sum.discount + item.discount
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
  // NHÓM DÒNG THEO TÊN CT
  //
  // Một dòng tổng chương trình, bên dưới là
  // các vùng được giữ theo thứ tự file phân bổ.
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
  // NPP CHO MỘT VÙNG
  // ----------------------------------------------

  function getNppsForZone(
    item: Summary
  ): NppSummary[] {
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
  }


  // ----------------------------------------------
  // UPLOAD FILE
  // ----------------------------------------------

  async function loadFile(
    type: "allocation" | "actual",
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

      } else {
        const data = await parseActual(file);

        setActual(data);
        setActualName(file.name);

        setStatus(
          "Đã đọc file 5.1 từ các cột E, G, N, W, AK."
        );
      }

      setReport(null);
      setStatusType("success");

    } catch (error: any) {
      setStatus(
        error?.message || "Không đọc được file."
      );

      setStatusType("error");

    } finally {
      setBusy(false);
    }
  }


  // ----------------------------------------------
  // CHỌN THÁNG
  // ----------------------------------------------

  function changeMonth(value: string): void {
    if (!value || value === month) return;

    setMonth(value);

    // Xóa file đã chọn trước đó để tránh vô tình
    // lưu dữ liệu tháng cũ vào tháng mới.
    setAllocation(null);
    setActual(null);

    setAllocationName("");
    setActualName("");

    setExpandedZones([]);
    setStatus("");
    setSavedAt("");
  }


  // ----------------------------------------------
  // LƯU BÁO CÁO
  // ----------------------------------------------

  async function processData(): Promise<void> {
    if (!allocation || !actual) {
      setStatus(
        "Vui lòng upload đủ file phân bổ và file 5.1."
      );

      setStatusType("error");
      return;
    }

    if (!month) {
      setStatus("Vui lòng chọn tháng dữ liệu.");
      setStatusType("error");
      return;
    }

    setBusy(true);
    setStatus("");
    setSavedAt("");

    try {
      const result = buildReport(
        allocation,
        actual
      );

      // Hiển thị báo cáo đã tính toán.
      setReport(result);

      // Lưu báo cáo theo tháng đã chọn.
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
            report: result
          })
        }
      );

      const serverResult = await response.json();

      if (!response.ok || !serverResult.ok) {
        throw new Error(
          serverResult.error ||
          "Không lưu được báo cáo vào Google Sheets."
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
        "Có lỗi trong quá trình đối chiếu hoặc lưu dữ liệu."
      );

      setStatusType("error");

    } finally {
      setBusy(false);
    }
  }


  // ----------------------------------------------
  // XUẤT EXCEL THEO BỘ LỌC
  // ----------------------------------------------

  function exportExcel(): void {
    if (!report) return;

    const workbook = XLSX.utils.book_new();

    const trackingRows: unknown[][] = [
      [
        "Loại dòng",
        "Tên Miền",
        "Tên vùng",
        "Tên CT",
        "Suất phân bổ",
        "Suất đã dùng",
        "% sử dụng",
        "Tổng chiết khấu",
        "Trạng thái"
      ]
    ];

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


    const nppRows: unknown[][] = [
      [
        "Tên Miền",
        "Tên vùng",
        "Tên CT",
        "Tên NPP",
        "Suất đã dùng",
        "Tổng chiết khấu"
      ]
    ];

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
      "Chi tiet NPP"
    );


    const warningRows: unknown[][] = [
      [
        "Tên Miền",
        "Tên vùng",
        "Tên CT",
        "Tên NPP",
        "Mã đơn hàng",
        "Tổng AK",
        "Lý do"
      ]
    ];

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
      "Canh bao"
    );


    XLSX.writeFile(
      workbook,
      `Tracking_phan_bo_${month}.xlsx`
    );
  }


  // ----------------------------------------------
  // ĐẶT LẠI BỘ LỌC
  // ----------------------------------------------

  function resetFilters(): void {
    setProgram("ALL");
    setRegion("ALL");
    setArea("ALL");
    setSearch("");
    setExpandedZones([]);
  }


  // ==================================================
  // RENDER GIAO DIỆN
  // ==================================================

  return (
    <main className="wrap">

      {/* ------------------------------------------
          A. HEADER
      ------------------------------------------ */}

      <header className="header">
        <div>
          <h1 className="pageTitle">
            Theo dõi phân bổ chương trình
          </h1>

          <p className="headerDescription">
            Theo dõi số suất thực tế, tiến độ sử dụng
            và chiết khấu theo chương trình, vùng và NPP.
          </p>

          <div className="headerMeta">
            <span>
              Tháng đang xem: <strong>{month}</strong>
            </span>

            {savedAt && (
              <span>
                Cập nhật dữ liệu: {savedAt}
              </span>
            )}
          </div>
        </div>
      </header>


      {/* ------------------------------------------
          B. UPLOAD FILE — CÓ THỂ THU GỌN
      ------------------------------------------ */}

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

            Tải dữ liệu dành cho quản trị viên
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
                Cập nhật lại tháng đang chọn sẽ thay
                dữ liệu tháng đó, không xóa các tháng khác.
              </p>
            </div>


            <div className="uploadgrid">

              <div className="uploadbox">
                <h3>File phân bổ</h3>

                <p className="muted">
                  6 cột theo mẫu đã thống nhất.
                </p>

                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  disabled={busy}
                  onChange={event => {
                    const file =
                      event.currentTarget.files?.[0];

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
                <h3>File 5.1</h3>

                <p className="muted">
                  Chỉ xử lý các cột E, G, N, W và AK.
                </p>

                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  disabled={busy}
                  onChange={event => {
                    const file =
                      event.currentTarget.files?.[0];

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
                className="btn"
                type="button"
                disabled={
                  busy ||
                  historyLoading ||
                  !allocation ||
                  !actual
                }
                onClick={processData}
              >
                {busy
                  ? "Đang xử lý..."
                  : "Kiểm tra và lưu Google Sheets"}
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
                  Mã CT được dùng để mapping; giao diện
                  hiển thị và nhóm số liệu theo Tên CT.
                </li>

                <li>
                  Một Tên CT + Vùng + Mã đơn hàng chỉ
                  tính tối đa một suất nếu có AK &gt; 0.
                </li>

                <li>
                  Các dòng AK được cộng tiền theo đơn hàng
                  và NPP tương ứng.
                </li>

                <li>
                  Dữ liệu không có phân bổ được đưa vào
                  cảnh báo, không cộng KPI.
                </li>
              </ul>
            </div>

          </div>
        )}
      </section>


      {/* ------------------------------------------
          C. BỘ LỌC VÀ LỊCH SỬ
      ------------------------------------------ */}

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
                <option
                  key={item}
                  value={item}
                >
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
                <option
                  key={item}
                  value={item}
                >
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
              placeholder="Miền, vùng, NPP, tên CT..."
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
            Tháng đã lưu:
          </span>

          {availableMonths.length === 0 && (
            <span className="muted">
              Chưa có lịch sử
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
              Đang tải dữ liệu tháng...
            </span>
          )}
        </div>
      </section>


      {/* ------------------------------------------
          D. DANH MỤC TÊN CT
      ------------------------------------------ */}

      {report && (
        <section className="panel programPanel">
          <div className="sectionHeadingRow">
            <div>
              <h2 className="sectionTitle">
                Danh mục chương trình
              </h2>

              <p className="muted">
                Các mã CT cùng Tên CT được tổng hợp chung.
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
                program === "ALL"
                  ? "programTab active"
                  : "programTab"
              }
              onClick={() => setProgram("ALL")}
            >
              Tất cả
            </button>

            {programs.map(item => (
              <button
                type="button"
                key={item}
                className={
                  normalizeKey(program) === normalizeKey(item)
                    ? "programTab active"
                    : "programTab"
                }
                onClick={() => setProgram(item)}
                title={item}
              >
                {item}
              </button>
            ))}
          </div>
        </section>
      )}


      {/* ------------------------------------------
          E. KPI
      ------------------------------------------ */}

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
              Đơn hàng đủ điều kiện sử dụng suất
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
              Tổng AK theo bộ lọc
            </div>
          </div>

        </section>
      )}


      {/* ------------------------------------------
          F. BẢNG TRACKING CHÍNH
      ------------------------------------------ */}

      {report && (
        <section className="panel trackingPanel">
          <div className="sectionHeadingRow">
            <div>
              <h2 className="sectionTitle">
                Tiến độ sử dụng theo chương trình
              </h2>

              <p className="muted">
                Dòng đậm là tổng chương trình.
                Bấm vào vùng để xem chi tiết NPP.
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
                {programGroups.map(group => {
                  const groupPercent = group.percent;

                  const groupWidth = Math.min(
                    100,
                    Math.max(0, groupPercent)
                  );

                  return (
                    <FragmentProgramGroup
                      key={normalizeKey(group.programName)}
                      group={group}
                      groupWidth={groupWidth}
                      expandedZones={expandedZones}
                      toggleZone={(key: string) => {
                        setExpandedZones(previous =>
                          previous.includes(key)
                            ? previous.filter(
                                item => item !== key
                              )
                            : [...previous, key]
                        );
                      }}
                      getNppsForZone={getNppsForZone}
                      fmt={fmt}
                      money={money}
                    />
                  );
                })}

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


      {/* ------------------------------------------
          G. CẢNH BÁO
      ------------------------------------------ */}

      {report && (
        <section className="panel">
          <div className="sectionHeadingRow">
            <div>
              <h2 className="sectionTitle">
                Cảnh báo dữ liệu
              </h2>

              <p className="muted">
                Các dòng này không được cộng vào KPI
                nếu không có phân bổ.
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


      {/* ------------------------------------------
          TRẠNG THÁI CHƯA CÓ DỮ LIỆU
      ------------------------------------------ */}

      {!report && !historyLoading && (
        <section className="panel emptyState">
          <h2>Chưa có dữ liệu tháng {month}</h2>

          <p className="muted">
            Mở khu vực tải dữ liệu, upload file phân bổ
            và file 5.1, sau đó nhấn kiểm tra và lưu.
          </p>
        </section>
      )}


      <footer className="footer">
        Dashboard theo dõi phân bổ CTKM & CTTL.
        Vui lòng kiểm tra cảnh báo và đối chiếu số liệu
        trước khi sử dụng báo cáo chính thức.
      </footer>

    </main>
  );
}


// ==================================================
// COMPONENT BẢNG: MỖI CHƯƠNG TRÌNH VÀ CÁC VÙNG
// ==================================================

function FragmentProgramGroup(props: {
  group: ProgramGroup;
  groupWidth: number;
  expandedZones: string[];
  toggleZone: (key: string) => void;
  getNppsForZone: (item: Summary) => NppSummary[];
  fmt: (value: number) => string;
  money: (value: number) => string;
}) {
  const group = props.group;

  return (
    <>
      {/* DÒNG TỔNG CHƯƠNG TRÌNH */}

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

        <td className="num usedValue">
          {props.fmt(group.used)}
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
                width: `${props.groupWidth}%`
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


      {/* CHI TIẾT VÙNG */}

      {group.rows.map((item, index) => {
        const key = detailKey(
          item.regionName,
          item.areaName,
          item.programName
        );

        const expanded =
          props.expandedZones.includes(key);

        const npps = props.getNppsForZone(item);

        const percentWidth = Math.min(
          100,
          Math.max(0, item.percent)
        );

        return (
          <FragmentZoneRow
            key={key + index}
            item={item}
            expanded={expanded}
            npps={npps}
            percentWidth={percentWidth}
            toggle={() => props.toggleZone(key)}
            fmt={props.fmt}
            money={props.money}
          />
        );
      })}
    </>
  );
}


// ==================================================
// COMPONENT CHI TIẾT VÙNG VÀ NPP
// ==================================================

function FragmentZoneRow(props: {
  item: Summary;
  expanded: boolean;
  npps: NppSummary[];
  percentWidth: number;
  toggle: () => void;
  fmt: (value: number) => string;
  money: (value: number) => string;
}) {
  const item = props.item;

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

        <td className="num usedValue">
          {props.fmt(item.used)}
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
                width: `${props.percentWidth}%`
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


      {/* BẢNG NPP MỞ RỘNG */}

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
                      <th>NPP</th>
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


// ==================================================
// CLASS CSS CHO TRẠNG THÁI
// ==================================================

function statusClass(status: StatusName): string {
  if (status === "Vượt định biên") {
    return "statusOver";
  }

  if (status === "Sắp đầy") {
    return "statusNear";
  }

  return "statusNormal";
}
