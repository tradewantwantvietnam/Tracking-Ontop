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

function mappingKey(
  areaName: string,
  programCode: string
): string {
  return (
    normalizeKey(areaName) +
    "|||" +
    normalizeKey(programCode)
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

function uniqueInOrder(values: string[]): string[] {
  const result: string[] = [];
  const seen = new Set<string>();

  for (const value of values) {
    const text = cleanText(value);

    if (!text) continue;

    const key = normalizeKey(text);

    if (!seen.has(key)) {
      seen.add(key);
      result.push(text);
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
// TÍNH TOÁN VÀ GỘP DỮ LIỆU
// ==================================================

function buildReport(
  allocations: Allocation[],
  actuals: Actual[]
): Report {

  // ----------------------------------------------
  // A. MAPPING: VÙNG + MÃ CT -> DÒNG PHÂN BỔ
  // ----------------------------------------------

  const allocationByCode: Map<string, Allocation> =
    new Map();

  for (const item of allocations) {
    const key = mappingKey(
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
  // Gộp các MÃ CT có cùng TÊN CT.
  // Giữ thứ tự lần xuất hiện đầu tiên từ file phân bổ.
  // ----------------------------------------------

  const summaryMap: Map<string, Summary> =
    new Map();

  for (const item of allocations) {
    const programName =
      cleanText(item.programName) ||
      "Chưa đặt tên CT";

    const key = detailKey(
      item.regionName,
      item.areaName,
      programName
    );

    if (!summaryMap.has(key)) {
      summaryMap.set(key, {
        regionName: item.regionName,
        areaName: item.areaName,
        programName,
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
  // TÊN CT + VÙNG + MÃ ĐƠN HÀNG.
  //
  // Không dùng Mã CT làm khóa khử trùng.
  // ----------------------------------------------

  type OrderRecord = {
    row: Actual;
    allocation: Allocation;
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
  // D. GHI NHẬN NPP
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
      detailKey(regionName, areaName, programName) +
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
  // F. XỬ LÝ TỪNG DÒNG 5.1
  // ----------------------------------------------

  actuals.forEach((row, rowIndex) => {
    const allocation = allocationByCode.get(
      mappingKey(row.areaName, row.programCode)
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

      // Vẫn cộng tiền, không tự đếm suất.
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
    // F3. NHÓM ĐƠN HÀNG THEO TÊN CT + VÙNG + W
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
      allocation
    });

    // Cộng mọi dòng AK của đơn hàng.
    group.discount += Number(row.discount) || 0;

    // Chỉ cần ít nhất một AK > 0.
    if (Number(row.discount) > 0) {
      group.hasPositive = true;
    }
  });


  // ----------------------------------------------
  // G. TÍNH SUẤT ĐÃ DÙNG VÀ CHI TIẾT ĐƠN HÀNG
  // ----------------------------------------------

  const usedOrders: UsedOrder[] = [];

  for (const group of orderGroups.values()) {
    const summaryKey = detailKey(
      group.regionName,
      group.areaName,
      group.programName
    );

    const summary = summaryMap.get(summaryKey);

    if (!summary) continue;

    // Tiền chiết khấu được cộng tất cả các dòng.
    summary.discount += group.discount;

    // Nhóm không có AK dương không sử dụng suất.
    if (group.hasPositive) {
      summary.used += 1;
    }


    // --------------------------------------------
    // G1. GỘP CHIẾT KHẤU THEO NPP
    // --------------------------------------------

    const discountsByNpp: Map<string, {
      nppName: string;
      discount: number;
    }> = new Map();

    for (const record of group.records) {
      const name =
        cleanText(record.row.npp) ||
        "(Không có tên NPP)";

      const key = normalizeKey(name);

      if (!discountsByNpp.has(key)) {
        discountsByNpp.set(key, {
          nppName: name,
          discount: 0
        });
      }

      discountsByNpp.get(key)!.discount +=
        Number(record.row.discount) || 0;
    }


    // --------------------------------------------
    // G2. GÁN SUẤT CHO NPP
    //
    // Mỗi đơn hàng chỉ có một suất.
    // Suất gán cho NPP đầu tiên trên dòng AK > 0.
    // Tiền vẫn được ghi riêng theo từng NPP.
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

    for (const item of discountsByNpp.values()) {
      const usedToAdd =
        group.hasPositive &&
        normalizeKey(item.nppName) === normalizeKey(slotOwner)
          ? 1
          : 0;

      addNpp(
        group.regionName,
        group.areaName,
        group.programName,
        item.nppName,
        item.discount,
        usedToAdd
      );
    }


    // --------------------------------------------
    // G3. LƯU CHI TIẾT ĐỂ TRA CỨU KHI BẤM "ĐÃ DÙNG"
    // --------------------------------------------

    if (group.hasPositive) {
      usedOrders.push({
        regionName: group.regionName,
        areaName: group.areaName,
        programName: group.programName,
        orderCode:
          cleanText(group.records[0].row.orderCode),
        slotOwner,
        npps: [...discountsByNpp.values()]
          .map(item => item.nppName),
        discount: group.discount
      });
    }
  }


  // ----------------------------------------------
  // H. HOÀN TẤT KPI VÀ TRẠNG THÁI
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
    useState(false);

  const [status, setStatus] = useState("");

  const [statusType, setStatusType] =
    useState<"success" | "error" | "info">("info");

  const [savedAt, setSavedAt] = useState("");

  const [region, setRegion] = useState("ALL");

  const [area, setArea] = useState("ALL");

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

      try {
        const [monthsResponse, reportResponse] =
          await Promise.all([
            fetch("/api/save?action=months", {
              cache: "no-store"
            }),

            fetch(
              "/api/save?action=report&month=" +
              encodeURIComponent(month),
              { cache: "no-store" }
            )
          ]);

        const monthsData = await monthsResponse.json();
        const reportData = await reportResponse.json();

        if (!monthsResponse.ok || !monthsData.ok) {
          throw new Error(
            monthsData.error ||
            "Không tải được danh sách tháng."
          );
        }

        if (!reportResponse.ok || !reportData.ok) {
          throw new Error(
            reportData.error ||
            "Không tải được lịch sử tháng."
          );
        }

        if (cancelled) return;

        setAvailableMonths(
          uniqueInOrder([
            month,
            ...(monthsData.months || [])
          ]).sort((a, b) => b.localeCompare(a))
        );

        if (reportData.found && reportData.report) {
          const storedReport = reportData.report as Report;

          // Các tháng được lưu trước khi bổ sung lịch sử
          // chi tiết đơn hàng có thể chưa có usedOrders.
          if (!Array.isArray(storedReport.usedOrders)) {
            storedReport.usedOrders = [];
          }

          setReport(storedReport);

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
            " chưa có dữ liệu. Hãy mở khu vực tải dữ liệu để upload."
          );

          setStatusType("info");
        }

      } catch (error: any) {
        if (cancelled) return;

        setReport(null);

        setStatus(
          error?.message ||
          "Không tải được lịch sử."
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
  // BỘ LỌC MIỀN VÀ VÙNG
  // ----------------------------------------------

  const regions = useMemo(
    () =>
      uniqueInOrder(
        (report?.summaries || []).map(
          item => item.regionName
        )
      ),
    [report]
  );

  const areas = useMemo(
    () =>
      uniqueInOrder(
        (report?.summaries || [])
          .filter(
            item =>
              region === "ALL" ||
              normalizeKey(item.regionName) === normalizeKey(region)
          )
          .map(item => item.areaName)
      ),
    [report, region]
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
  // LỌC BẢNG CHÍNH THEO MIỀN / VÙNG / TỪ KHÓA
  // ----------------------------------------------

  const baseFiltered = useMemo(
    () =>
      (report?.summaries || []).filter(
        item =>
          (
            region === "ALL" ||
            normalizeKey(item.regionName) === normalizeKey(region)
          ) &&
          (
            area === "ALL" ||
            normalizeKey(item.areaName) === normalizeKey(area)
          )
      ),
    [report, region, area]
  );

  const baseNpps = useMemo(
    () =>
      (report?.npps || []).filter(
        item =>
          (
            region === "ALL" ||
            normalizeKey(item.regionName) === normalizeKey(region)
          ) &&
          (
            area === "ALL" ||
            normalizeKey(item.areaName) === normalizeKey(area)
          )
      ),
    [report, region, area]
  );

  const query = normalizeKey(search);

  function matchesSummarySearch(item: Summary): boolean {
    if (!query) return true;

    return [
      item.regionName,
      item.areaName,
      item.programName
    ].some(value =>
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
          matchesSummarySearch(item) ||
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
          [
            item.regionName,
            item.areaName,
            item.programName
          ].some(value =>
            normalizeKey(value).includes(query)
          )
        );
      }),
    [baseNpps, query]
  );


  // ----------------------------------------------
  // KPI
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
    const groups: Map<string, ProgramGroup> =
      new Map();

    for (const item of filtered) {
      const key = normalizeKey(item.programName);

      if (!groups.has(key)) {
        groups.set(key, {
          programName: item.programName,
          rows: [],
          allocation: 0,
          used: 0,
          discount: 0,
          percent: 0,
          status: "Bình thường"
        });
      }

      const group = groups.get(key)!;

      group.rows.push(item);
      group.allocation += item.allocation;
      group.used += item.used;
      group.discount += item.discount;
    }

    return [...groups.values()].map(group => {
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
  // CÁC ĐƠN HÀNG DÙNG CHO CỬA SỔ TRA CỨU
  // ----------------------------------------------

  const activeUsedOrders = useMemo(() => {
    if (!report || !selectedUsedScope) return [];

    return (report.usedOrders || []).filter(item => {
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
        const searchFields = [
          item.regionName,
          item.areaName,
          item.programName,
          item.orderCode,
          item.slotOwner,
          ...(item.npps || [])
        ];

        if (
          !searchFields.some(value =>
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
    [report, region, area, query]
  );


  // ----------------------------------------------
  // MỞ CHI TIẾT SUẤT ĐÃ DÙNG
  // ----------------------------------------------

  function openUsedDetails(scope: UsedScope) {
    setSelectedUsedScope(scope);
  }

  function closeUsedDetails() {
    setSelectedUsedScope(null);
  }


  // ----------------------------------------------
  // THAY THÁNG
  // ----------------------------------------------

  function changeMonth(value: string) {
    if (!value || value === month) return;

    setAllocation(null);
    setActual(null);

    setAllocationName("");
    setActualName("");

    setExpandedZones([]);
    setSelectedUsedScope(null);

    setStatus("");
    setSavedAt("");

    setMonth(value);
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
          "Đã đọc dữ liệu file 5.1 từ các cột E, G, N, W, AK."
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
  // LƯU DỮ LIỆU THÁNG
  // ----------------------------------------------

  async function processData(): Promise<void> {
    if (!allocation || !actual) {
      setStatus(
        "Vui lòng upload đủ file phân bổ và file 5.1."
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
        actual
      );

      setReport(result);

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
        "."
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
  // XUẤT EXCEL THEO BỘ LỌC
  // ----------------------------------------------

  function exportExcel(): void {
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

    for (const item of programGroups) {
      trackingRows.push([
        "Tổng chương trình",
        "",
        "",
        item.programName,
        item.allocation,
        item.used,
        item.percent,
        item.discount,
        item.status
      ]);

      for (const row of item.rows) {
        trackingRows.push([
          "Chi tiết vùng",
          row.regionName,
          row.areaName,
          row.programName,
          row.allocation,
          row.used,
          row.percent,
          row.discount,
          row.status
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
      "Danh sách NPP trên đơn",
      "Tổng chiết khấu"
    ]];

    for (const item of (report.usedOrders || [])) {
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
  // XÓA BỘ LỌC
  // ----------------------------------------------

  function resetFilters() {
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
            Theo dõi tiến độ sử dụng suất và chiết khấu
            theo chương trình, vùng và NPP.
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


      {/* UPLOAD CÓ THỂ THU GỌN */}

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
                Upload lại sẽ thay dữ liệu tháng đang chọn.
                Các tháng khác được giữ nguyên.
              </p>
            </div>


            <div className="uploadgrid">
              <div className="uploadbox">
                <h3>File phân bổ</h3>

                <p className="muted">
                  6 cột theo bố cục đã thống nhất.
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
                <h3>File 5.1</h3>

                <p className="muted">
                  Chỉ đọc E, G, N, W, AK.
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
                disabled={busy || !allocation || !actual}
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
              <strong>Quy tắc:</strong>
              <ul>
                <li>
                  Giao diện chỉ hiển thị Tên CT.
                  Mã CT được dùng nội bộ để mapping.
                </li>
                <li>
                  Một Tên CT + Vùng + Mã đơn hàng
                  chỉ tính tối đa một suất nếu có AK &gt; 0.
                </li>
                <li>
                  Bấm số suất đã dùng để tra cứu đơn hàng
                  và NPP tạo ra các suất đó.
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
              Bấm số đã dùng trong bảng để xem chi tiết
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
                  width: `${Math.max(
                    0,
                    Math.min(100, totalPercent)
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


      {/* BẢNG TRACKING: TỔNG CHƯƠNG TRÌNH + VÙNG */}

      {report && (
        <section className="panel trackingPanel">
          <div className="sectionHeadingRow">
            <div>
              <h2 className="sectionTitle">
                Tiến độ sử dụng theo chương trình
              </h2>

              <p className="muted">
                Không hiển thị mã CT. Các mã CT có cùng
                Tên CT được gộp lại.
                Bấm số đã dùng để xem các đơn hàng tương ứng.
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

                  return (
                    <FragmentProgramGroup
                      key={normalizeKey(group.programName)}
                      group={group}
                      expandedZones={expandedZones}
                      toggleZone={(key: string) => {
                        setExpandedZones(previous =>
                          previous.includes(key)
                            ? previous.filter(item => item !== key)
                            : [...previous, key]
                        );
                      }}
                      getNppsForZone={(item: Summary) => {
                        const key = detailKey(
                          item.regionName,
                          item.areaName,
                          item.programName
                        );

                        return npps.filter(npp =>
                          detailKey(
                            npp.regionName,
                            npp.areaName,
                            npp.programName
                          ) === key
                        );
                      }}
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
                      Tổng cộng theo bộ lọc
                    </td>

                    <td className="num">
                      {fmt(total.allocation)}
                    </td>

                    <td className="num">
                      <button
                        type="button"
                        className="usedNumberButton"
                        onClick={() =>
                          setSelectedUsedScope(null)
                        }
                        disabled
                        title="Tổng cộng được thể hiện ở KPI phía trên"
                      >
                        {fmt(total.used)}
                      </button>
                    </td>

                    <td>
                      <strong>{fmt(totalPercent)}%</strong>
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
                Các trường hợp không có phân bổ hoặc thiếu
                mã đơn hàng được hiển thị riêng.
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
              closeUsedDetails();
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
                  id="usedModalTitle"
                  className="sectionTitle"
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
                <span>Số đơn hàng được tính suất</span>
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
              Mỗi dòng bên dưới tương ứng một đơn hàng
              được tính một suất. “NPP nhận suất” là NPP
              được gán suất nếu đơn hàng có nhiều NPP.
              Tổng chiết khấu được cộng tất cả dòng AK
              thuộc đơn hàng.
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
                      <td>{(item.npps || []).join(", ")}</td>
                      <td className="num">
                        {money(item.discount)}
                      </td>
                    </tr>
                  ))}

                  {activeUsedOrders.length === 0 && (
                    <tr>
                      <td colSpan={6} className="noData">
                        Không tìm thấy chi tiết đơn hàng.
                        Nếu đây là dữ liệu lịch sử được lưu
                        trước khi bổ sung chức năng tra cứu,
                        hãy upload lại tháng đó để lưu chi tiết.
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


      {!report && !historyLoading && (
        <section className="panel emptyState">
          <h2>
            Chưa có dữ liệu tháng {month}
          </h2>

          <p className="muted">
            Mở khu vực Tải dữ liệu, upload hai file và
            nhấn Kiểm tra và lưu dữ liệu.
          </p>
        </section>
      )}


      <footer className="footer">
        Dashboard theo dõi phân bổ CTKM / CTTL.
        Vui lòng kiểm tra cảnh báo và đối chiếu dữ liệu
        với file hệ thống trước khi sử dụng báo cáo chính thức.
      </footer>

    </main>
  );
}


// ==================================================
// DÒNG TỔNG CHƯƠNG TRÌNH VÀ CHI TIẾT VÙNG
// ==================================================

function FragmentProgramGroup(props: {
  group: ProgramGroup;
  expandedZones: string[];
  toggleZone: (key: string) => void;
  getNppsForZone: (item: Summary) => NppSummary[];
  fmt: (value: number) => string;
  money: (value: number) => string;
  onClickProgramUsed: () => void;
  onClickZoneUsed: (item: Summary) => void;
}) {
  const group = props.group;

  const groupPercent = Math.max(
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
            title="Bấm để tra cứu các đơn hàng đã dùng suất"
            disabled={group.used === 0}
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
              style={{ width: `${groupPercent}%` }}
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

        const percentWidth = Math.max(
          0,
          Math.min(100, item.percent)
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
            onClickUsed={() =>
              props.onClickZoneUsed(item)
            }
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
  percentWidth: number;
  toggle: () => void;
  fmt: (value: number) => string;
  money: (value: number) => string;
  onClickUsed: () => void;
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

        <td className="num">
          <button
            type="button"
            className="usedNumberButton"
            onClick={props.onClickUsed}
            title="Bấm để xem đơn hàng và NPP tạo ra các suất đã dùng"
            disabled={item.used === 0}
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
              style={{ width: `${props.percentWidth}%` }}
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
                      <th className="num">Suất đã dùng</th>
                      <th className="num">Tổng chiết khấu</th>
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
