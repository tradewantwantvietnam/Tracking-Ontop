"use client";

import { useEffect, useMemo, useState } from "react";

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


// ==================================================
// KHAI BÁO KIỂU DỮ LIỆU
// ==================================================

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


// ==================================================
// CÁC HÀM HỖ TRỢ
// ==================================================

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

function normalizeText(value: unknown): string {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeKey(value: string): string {
  return normalizeText(value).toLocaleLowerCase("vi");
}

// Khóa đối chiếu phân bổ: Tên vùng + Mã CT.
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


// ==================================================
// TÍNH TOÁN BÁO CÁO
// ==================================================

function buildReport(
  allocations: Allocation[],
  actuals: Actual[]
): Report {

  // ------------------------------------------------
  // A. LẬP DANH SÁCH PHÂN BỔ
  // ------------------------------------------------

  const allocationMap: Map<string, Allocation> = new Map();

  for (const item of allocations) {
    const key = dataKey(
      item.areaName,
      item.programCode
    );

    if (allocationMap.has(key)) {
      throw new Error(
        "File phân bổ có Tên vùng + Mã CT bị trùng: " +
        item.areaName +
        " / " +
        item.programCode +
        ". Hãy kiểm tra file phân bổ."
      );
    }

    allocationMap.set(key, item);
  }


  // ------------------------------------------------
  // B. KHỞI TẠO TỔNG HỢP VÙNG + MÃ CT
  // ------------------------------------------------

  // Dùng khai báo kiểu Map riêng biệt để tránh
  // lỗi phân tích cú pháp tại new Map<...>().
  const summaryMap: Map<string, Summary> = new Map();

  for (const item of allocations) {
    const key = dataKey(
      item.areaName,
      item.programCode
    );

    summaryMap.set(key, {
      regionName: item.regionName,
      areaName: item.areaName,
      programCode: item.programCode,
      programName: item.programName,
      structure: item.structure,
      allocation: item.allocation,
      used: 0,
      discount: 0,
      remaining: item.allocation,
      percent: 0
    });
  }


  // ------------------------------------------------
  // C. NHÓM ĐƠN HÀNG
  //
  // Mã CT + Tên vùng + Mã đơn hàng
  // ------------------------------------------------

  type OrderGroup = {
    allocation: Allocation;
    records: Actual[];
    discount: number;
    hasPositive: boolean;
  };

  const orderGroups: Map<string, OrderGroup> = new Map();


  // ------------------------------------------------
  // D. TỔNG HỢP THEO NPP
  // ------------------------------------------------

  const nppMap: Map<string, NppSummary> = new Map();

  function addNpp(
    allocation: Allocation,
    nppName: string,
    discountToAdd: number,
    usedToAdd: number
  ): void {
    const displayName =
      normalizeText(nppName) ||
      "(Không có tên NPP)";

    const key =
      dataKey(
        allocation.areaName,
        allocation.programCode
      ) +
      "|||" +
      normalizeKey(displayName);

    if (!nppMap.has(key)) {
      nppMap.set(key, {
        regionName: allocation.regionName,
        areaName: allocation.areaName,
        programCode: allocation.programCode,
        npp: displayName,
        used: 0,
        discount: 0
      });
    }

    const item = nppMap.get(key)!;

    item.discount += discountToAdd;
    item.used += usedToAdd;
  }


  // ------------------------------------------------
  // E. DANH SÁCH CẢNH BÁO
  // ------------------------------------------------

  const unmatchedWarnings: Map<string, Warning> = new Map();

  const missingOrderWarnings: Warning[] = [];


  // ------------------------------------------------
  // F. ĐỌC DỮ LIỆU FILE 5.1
  // ------------------------------------------------

  actuals.forEach((row, rowIndex) => {
    const key = dataKey(
      row.areaName,
      row.programCode
    );

    const allocation = allocationMap.get(key);


    // ----------------------------------------------
    // F1. KHÔNG CÓ PHÂN BỔ
    // ----------------------------------------------

    if (!allocation) {
      const orderCode = normalizeText(row.orderCode);

      const warningKey = orderCode
        ? key + "|||" + normalizeKey(orderCode)
        : "MISSING_ORDER_ROW|||" + rowIndex;

      if (!unmatchedWarnings.has(warningKey)) {
        unmatchedWarnings.set(warningKey, {
          regionName: "",
          areaName: row.areaName,
          programCode: row.programCode,
          npp: row.npp,
          orderCode: orderCode,
          discount: 0,
          reason:
            "Không có cặp Tên vùng + Mã CT trong file phân bổ"
        });
      }

      unmatchedWarnings.get(warningKey)!.discount +=
        row.discount;

      // Đang ở trong forEach nên dùng return,
      // không dùng continue.
      return;
    }


    // ----------------------------------------------
    // F2. THIẾU MÃ ĐƠN HÀNG
    // ----------------------------------------------

    if (!normalizeText(row.orderCode)) {
      const summary = summaryMap.get(key)!;

      // Vẫn cộng tiền AK nếu có phân bổ.
      summary.discount += row.discount;

      // NPP được ghi nhận tiền nhưng chưa tính suất.
      addNpp(
        allocation,
        row.npp,
        row.discount,
        0
      );

      missingOrderWarnings.push({
        regionName: allocation.regionName,
        areaName: allocation.areaName,
        programCode: allocation.programCode,
        npp: row.npp,
        orderCode: "",
        discount: row.discount,
        reason:
          "Thiếu mã đơn hàng. Tiền AK vẫn được cộng; " +
          "suất không tính để tránh đếm trùng sai."
      });

      return;
    }


    // ----------------------------------------------
    // F3. TẠO NHÓM ĐƠN HÀNG
    // ----------------------------------------------

    const orderKey =
      key +
      "|||" +
      normalizeKey(row.orderCode);

    if (!orderGroups.has(orderKey)) {
      orderGroups.set(orderKey, {
        allocation: allocation,
        records: [],
        discount: 0,
        hasPositive: false
      });
    }

    const group = orderGroups.get(orderKey)!;

    group.records.push(row);

    // Cộng toàn bộ AK của các dòng trong đơn hàng.
    group.discount += row.discount;

    // Chỉ cần một dòng AK > 0 là tính một suất.
    if (row.discount > 0) {
      group.hasPositive = true;
    }
  });


  // ------------------------------------------------
  // G. TÍNH SUẤT VÀ TIỀN CHIẾT KHẤU
  // ------------------------------------------------

  for (const group of orderGroups.values()) {
    const allocation = group.allocation;

    const key = dataKey(
      allocation.areaName,
      allocation.programCode
    );

    const summary = summaryMap.get(key)!;

    // Tổng tiền AK.
    summary.discount += group.discount;

    // Khử trùng đơn hàng: một nhóm chỉ tính một suất.
    if (group.hasPositive) {
      summary.used += 1;
    }


    // ----------------------------------------------
    // G1. CỘNG TIỀN THEO TỪNG NPP
    // ----------------------------------------------

    type NppDiscount = {
      displayName: string;
      discount: number;
    };

    const nppDiscounts: Map<string, NppDiscount> = new Map();

    for (const row of group.records) {
      const displayName =
        normalizeText(row.npp) ||
        "(Không có tên NPP)";

      const nppKey = normalizeKey(displayName);

      if (!nppDiscounts.has(nppKey)) {
        nppDiscounts.set(nppKey, {
          displayName: displayName,
          discount: 0
        });
      }

      nppDiscounts.get(nppKey)!.discount +=
        row.discount;
    }


    // ----------------------------------------------
    // G2. GÁN MỘT SUẤT CHO NPP
    //
    // Nếu một đơn hàng có nhiều NPP:
    // - Tổng vùng chỉ tính một suất.
    // - Suất được gán cho NPP đầu tiên có dòng AK > 0.
    // - Tiền AK vẫn được cộng riêng theo NPP.
    // ----------------------------------------------

    const firstPositiveRow = group.records.find(
      item => item.discount > 0
    );

    const slotOwner = normalizeKey(
      firstPositiveRow?.npp ||
      "(Không có tên NPP)"
    );

    for (const [nppKey, item] of nppDiscounts) {
      const usedToAdd =
        group.hasPositive && nppKey === slotOwner
          ? 1
          : 0;

      addNpp(
        allocation,
        item.displayName,
        item.discount,
        usedToAdd
      );
    }
  }


  // ------------------------------------------------
  // H. TÍNH SỐ SUẤT CÒN LẠI VÀ %
  // ------------------------------------------------

  const summaries: Summary[] = [...summaryMap.values()].map(
    item => {
      const remaining =
        item.allocation - item.used;

      const percent =
        item.allocation > 0
          ? (item.used / item.allocation) * 100
          : 0;

      return {
        ...item,
        remaining,
        percent
      };
    }
  );


  // ------------------------------------------------
  // I. TRẢ KẾT QUẢ
  // ------------------------------------------------

  return {
    summaries: summaries,
    npps: [...nppMap.values()],
    warnings: [
      ...unmatchedWarnings.values(),
      ...missingOrderWarnings
    ]
  };
}


// ==================================================
// GIAO DIỆN DASHBOARD
// ==================================================

export default function Home() {
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

  const [status, setStatus] = useState("");

  const [statusType, setStatusType] =
    useState<"success" | "error" | "info">("info");

  const [savedAt, setSavedAt] = useState("");

  const [program, setProgram] = useState("ALL");
  const [region, setRegion] = useState("ALL");
  const [area, setArea] = useState("ALL");


  // ----------------------------------------------
  // DANH SÁCH BỘ LỌC
  // ----------------------------------------------

  const programs = useMemo(
    () =>
      uniqueValues(
        (report?.summaries || []).map(
          item => item.programCode
        )
      ),
    [report]
  );

  const regions = useMemo(
    () =>
      uniqueValues(
        (report?.summaries || [])
          .filter(
            item =>
              program === "ALL" ||
              item.programCode === program
          )
          .map(item => item.regionName)
      ),
    [report, program]
  );

  const areas = useMemo(
    () =>
      uniqueValues(
        (report?.summaries || [])
          .filter(
            item =>
              (
                program === "ALL" ||
                item.programCode === program
              ) &&
              (
                region === "ALL" ||
                item.regionName === region
              )
          )
          .map(item => item.areaName)
      ),
    [report, program, region]
  );


  useEffect(() => {
    if (
      region !== "ALL" &&
      !regions.includes(region)
    ) {
      setRegion("ALL");
    }
  }, [regions, region]);


  useEffect(() => {
    if (
      area !== "ALL" &&
      !areas.includes(area)
    ) {
      setArea("ALL");
    }
  }, [areas, area]);


  // ----------------------------------------------
  // DỮ LIỆU THEO BỘ LỌC
  // ----------------------------------------------

  const filtered = useMemo(
    () =>
      (report?.summaries || []).filter(
        item =>
          (
            program === "ALL" ||
            item.programCode === program
          ) &&
          (
            region === "ALL" ||
            item.regionName === region
          ) &&
          (
            area === "ALL" ||
            item.areaName === area
          )
      ),
    [report, program, region, area]
  );


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

  const percent =
    total.allocation > 0
      ? (total.used / total.allocation) * 100
      : 0;


  // ----------------------------------------------
  // CHI TIẾT NPP
  // ----------------------------------------------

  const npps = useMemo(
    () =>
      (report?.npps || [])
        .filter(
          item =>
            (
              program === "ALL" ||
              item.programCode === program
            ) &&
            (
              region === "ALL" ||
              item.regionName === region
            ) &&
            (
              area === "ALL" ||
              item.areaName === area
            )
        )
        .sort((a, b) => b.used - a.used),
    [report, program, region, area]
  );


  // ----------------------------------------------
  // CẢNH BÁO
  // ----------------------------------------------

  const warnings = useMemo(
    () =>
      (report?.warnings || []).filter(
        item =>
          (
            program === "ALL" ||
            item.programCode === program
          ) &&
          (
            area === "ALL" ||
            item.areaName === area
          ) &&
          (
            region === "ALL" ||
            !item.regionName ||
            item.regionName === region
          )
      ),
    [report, program, region, area]
  );


  // ----------------------------------------------
  // BIỂU ĐỒ CHƯƠNG TRÌNH
  // ----------------------------------------------

  const byProgram = useMemo(() => {
    const map: Map<
      string,
      {
        name: string;
        allocated: number;
        used: number;
      }
    > = new Map();

    for (const item of filtered) {
      if (!map.has(item.programCode)) {
        map.set(item.programCode, {
          name: item.programCode,
          allocated: 0,
          used: 0
        });
      }

      const row = map.get(item.programCode)!;

      row.allocated += item.allocation;
      row.used += item.used;
    }

    return [...map.values()];
  }, [filtered]);


  // ----------------------------------------------
  // BIỂU ĐỒ VÙNG
  // ----------------------------------------------

  const byArea = useMemo(() => {
    const map: Map<
      string,
      {
        name: string;
        allocated: number;
        used: number;
      }
    > = new Map();

    for (const item of filtered) {
      if (!map.has(item.areaName)) {
        map.set(item.areaName, {
          name: item.areaName,
          allocated: 0,
          used: 0
        });
      }

      const row = map.get(item.areaName)!;

      row.allocated += item.allocation;
      row.used += item.used;
    }

    return [...map.values()]
      .sort((a, b) => b.used - a.used)
      .slice(0, 20);
  }, [filtered]);


  // ----------------------------------------------
  // UPLOAD FILE
  // ----------------------------------------------

  async function loadFile(
    type: "allocation" | "actual",
    file?: File
  ): Promise<void> {
    if (!file) {
      return;
    }

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
  // XỬ LÝ VÀ LƯU BÁO CÁO
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
      // Tính toán dữ liệu.
      const result = buildReport(
        allocation,
        actual
      );

      setReport(result);

      // Gửi dữ liệu đến API Vercel.
      // API sử dụng token phía máy chủ.
      try {
        const response = await fetch(
          "/api/save",
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              allocation: allocation,
              actual: actual,
              report: result
            })
          }
        );

        const serverResult = await response.json();

        if (
          !response.ok ||
          !serverResult.ok
        ) {
          throw new Error(
            serverResult.error ||
            "Không lưu được vào Google Sheets."
          );
        }

        setSavedAt(
          new Date().toLocaleString("vi-VN")
        );

        setStatus(
          "Đã đối chiếu và lưu kết quả vào Google Sheets."
        );

        setStatusType("success");

      } catch (saveError: any) {
        setStatus(
          "Dashboard đã tính toán nhưng chưa lưu được " +
          "Google Sheets: " +
          (
            saveError?.message ||
            "Không kết nối được API."
          )
        );

        setStatusType("error");
      }

    } catch (error: any) {
      setReport(null);

      setStatus(
        error?.message ||
        "Có lỗi khi đối chiếu dữ liệu."
      );

      setStatusType("error");

    } finally {
      setBusy(false);
    }
  }


  // ----------------------------------------------
  // RESET BỘ LỌC
  // ----------------------------------------------

  function resetFilters(): void {
    setProgram("ALL");
    setRegion("ALL");
    setArea("ALL");
  }


  // ==================================================
  // HIỂN THỊ GIAO DIỆN
  // ==================================================

  return (
    <main className="wrap">

      <header className="header">
        <div>
          <h1 style={{ margin: "0 0 6px", fontSize: 28 }}>
            Theo dõi phân bổ chương trình
          </h1>

          <div className="muted">
            Đối chiếu phân bổ · Suất đã dùng ·
            Chiết khấu · Phân tích NPP
          </div>
        </div>

        <button
          className="btn secondary"
          onClick={resetFilters}
          disabled={!report}
        >
          Xóa bộ lọc
        </button>
      </header>


      {/* ========================================== */}
      {/* 1. UPLOAD FILE */}
      {/* ========================================== */}

      <section className="panel">
        <h2 style={{ marginTop: 0, fontSize: 18 }}>
          1. Tải dữ liệu
        </h2>

        <div className="uploadgrid">

          <div className="uploadbox">
            <h3>File phân bổ</h3>

            <p className="muted">
              6 cột: Tên Miền, Tên vùng, Mã CT,
              Tên CT, Cơ cấu CT, Suất phân bổ.
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
              <p className="success muted">
                Đã chọn: {allocationName}
              </p>
            )}
          </div>


          <div className="uploadbox">
            <h3>File 5.1</h3>

            <p className="muted">
              Chỉ đọc E (Tên NPP), G (Tên vùng),
              N (Mã CT), W (Mã đơn hàng), AK (Chiết khấu).
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
              <p className="success muted">
                Đã chọn: {actualName}
              </p>
            )}
          </div>

        </div>


        <div className="actions" style={{ marginTop: 14 }}>
          <button
            className="btn"
            disabled={busy || !allocation || !actual}
            onClick={processData}
          >
            {busy
              ? "Đang xử lý..."
              : "Đối chiếu và cập nhật Dashboard"}
          </button>
        </div>


        {status && (
          <div className={`status ${statusType}`}>
            {status}
          </div>
        )}

        {savedAt && (
          <div className="muted" style={{ marginTop: 6 }}>
            Lần lưu gần nhất: {savedAt}
          </div>
        )}

        <div className="notice">
          <strong>Quy tắc xử lý:</strong>
          <br />
          Mã CT lấy phần đầu trước dấu phẩy.
          <br />
          Mỗi Mã CT + Vùng + Mã đơn hàng chỉ tính
          tối đa một suất nếu có ít nhất một dòng AK lớn hơn 0.
          <br />
          Tổng chiết khấu cộng tất cả dòng AK của đơn hàng.
          <br />
          Dữ liệu không có phân bổ được đưa vào cảnh báo,
          không cộng vào KPI phân bổ.
        </div>
      </section>


      {!report && (
        <section className="panel">
          <h2 style={{ marginTop: 0, fontSize: 18 }}>
            Dashboard chưa có dữ liệu
          </h2>

          <p className="muted">
            Vui lòng tải lên hai file và nhấn
            “Đối chiếu và cập nhật Dashboard”.
          </p>
        </section>
      )}


      {report && (
        <>

          {/* ======================================== */}
          {/* 2. BỘ LỌC */}
          {/* ======================================== */}

          <section className="panel">
            <h2 style={{ margin: "0 0 6px", fontSize: 18 }}>
              2. Bộ lọc Dashboard
            </h2>

            <div className="filters">

              <label>
                Chương trình
                <select
                  value={program}
                  onChange={event => setProgram(event.target.value)}
                >
                  <option value="ALL">Tất cả chương trình</option>
                  {programs.map(item => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </select>
              </label>


              <label>
                Miền
                <select
                  value={region}
                  onChange={event => setRegion(event.target.value)}
                >
                  <option value="ALL">Tất cả miền</option>
                  {regions.map(item => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </select>
              </label>


              <label>
                Vùng
                <select
                  value={area}
                  onChange={event => setArea(event.target.value)}
                >
                  <option value="ALL">Tất cả vùng</option>
                  {areas.map(item => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </select>
              </label>

            </div>

            <div className="muted">
              KPI, biểu đồ và bảng chi tiết thay đổi theo bộ lọc.
            </div>
          </section>


          {/* ======================================== */}
          {/* 3. KPI */}
          {/* ======================================== */}

          <div className="grid">

            <div className="metric">
              <div className="label">Tổng phân bổ</div>
              <div className="value">
                {fmt(total.allocation)} suất
              </div>
            </div>

            <div className="metric">
              <div className="label">Suất đã dùng</div>
              <div className="value">
                {fmt(total.used)} suất
              </div>
            </div>

            <div className="metric">
              <div className="label">Còn lại / vượt phân bổ</div>
              <div
                className="value"
                style={{
                  color:
                    total.allocation - total.used < 0
                      ? "#b42318"
                      : "inherit"
                }}
              >
                {fmt(total.allocation - total.used)} suất
              </div>
            </div>

            <div className="metric">
              <div className="label">Tỷ lệ sử dụng</div>
              <div className="value">
                {fmt(percent)}%
              </div>

              <div className="progress" style={{ marginTop: 10 }}>
                <div
                  style={{
                    width: `${Math.max(0, Math.min(100, percent))}%`
                  }}
                />
              </div>
            </div>

            <div className="metric">
              <div className="label">Tổng tiền chiết khấu</div>
              <div className="value" style={{ fontSize: 21 }}>
                {money(total.discount)}
              </div>
            </div>

          </div>


          {/* ======================================== */}
          {/* 4. BIỂU ĐỒ */}
          {/* ======================================== */}

          <div className="chartgrid" style={{ marginTop: 18 }}>

            <section className="panel chart">
              <h3>Phân bổ và đã dùng theo chương trình</h3>

              <ResponsiveContainer width="100%" height="88%">
                <BarChart data={byProgram}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                  <YAxis />
                  <Tooltip />
                  <Legend />

                  <Bar
                    dataKey="allocated"
                    name="Phân bổ"
                    fill="#A8B6D8"
                  />

                  <Bar
                    dataKey="used"
                    name="Đã dùng"
                    fill="#2B67F6"
                  />
                </BarChart>
              </ResponsiveContainer>
            </section>


            <section className="panel chart">
              <h3>Suất đã dùng và phân bổ theo vùng</h3>

              <ResponsiveContainer width="100%" height="88%">
                <BarChart
                  data={byArea}
                  layout="vertical"
                  margin={{ left: 20 }}
                >
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis type="number" />

                  <YAxis
                    type="category"
                    dataKey="name"
                    width={100}
                    tick={{ fontSize: 10 }}
                  />

                  <Tooltip />
                  <Legend />

                  <Bar
                    dataKey="used"
                    name="Đã dùng"
                    fill="#16A34A"
                  />

                  <Bar
                    dataKey="allocated"
                    name="Phân bổ"
                    fill="#A8B6D8"
                  />
                </BarChart>
              </ResponsiveContainer>
            </section>

          </div>


          {/* ======================================== */}
          {/* 5. BẢNG MIỀN / VÙNG / CT */}
          {/* ======================================== */}

          <section className="panel">
            <h2 style={{ fontSize: 18, marginTop: 0 }}>
              3. Chi tiết theo miền / vùng / CT
            </h2>

            <div className="tablewrap">
              <table>
                <thead>
                  <tr>
                    <th>Miền</th>
                    <th>Vùng</th>
                    <th>Mã CT</th>
                    <th>Tên CT</th>
                    <th className="num">Phân bổ</th>
                    <th className="num">Đã dùng</th>
                    <th className="num">Còn lại</th>
                    <th className="num">% dùng</th>
                    <th className="num">Tổng chiết khấu</th>
                  </tr>
                </thead>

                <tbody>
                  {filtered.map((item, index) => (
                    <tr
                      key={`${item.areaName}-${item.programCode}-${index}`}
                    >
                      <td>{item.regionName}</td>
                      <td>{item.areaName}</td>
                      <td>{item.programCode}</td>
                      <td>{item.programName}</td>
                      <td className="num">{fmt(item.allocation)}</td>
                      <td className="num">{fmt(item.used)}</td>
                      <td className="num">{fmt(item.remaining)}</td>
                      <td className="num">{fmt(item.percent)}%</td>
                      <td className="num">{money(item.discount)}</td>
                    </tr>
                  ))}

                  {filtered.length === 0 && (
                    <tr>
                      <td colSpan={9}>
                        Không có dữ liệu phù hợp với bộ lọc.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>


          {/* ======================================== */}
          {/* 6. BẢNG NPP */}
          {/* ======================================== */}

          <section className="panel">
            <h2 style={{ fontSize: 18, marginTop: 0 }}>
              4. Chi tiết sử dụng theo NPP
            </h2>

            <p className="muted">
              NPP không có định mức riêng. Số suất và tiền
              chiết khấu thể hiện phần đóng góp của NPP
              trong vùng và chương trình đang chọn.
            </p>

            <div className="tablewrap">
              <table>
                <thead>
                  <tr>
                    <th>Miền</th>
                    <th>Vùng</th>
                    <th>Mã CT</th>
                    <th>NPP</th>
                    <th className="num">Suất đã dùng</th>
                    <th className="num">Tổng chiết khấu</th>
                    <th className="num">Tỷ trọng suất trong vùng</th>
                  </tr>
                </thead>

                <tbody>
                  {npps.map((item, index) => {
                    const parent = filtered.find(
                      row =>
                        row.areaName === item.areaName &&
                        row.programCode === item.programCode
                    );

                    const contribution =
                      parent && parent.used > 0
                        ? (item.used / parent.used) * 100
                        : 0;

                    return (
                      <tr
                        key={`${item.areaName}-${item.programCode}-${item.npp}-${index}`}
                      >
                        <td>{item.regionName}</td>
                        <td>{item.areaName}</td>
                        <td>{item.programCode}</td>
                        <td>{item.npp}</td>
                        <td className="num">{fmt(item.used)}</td>
                        <td className="num">{money(item.discount)}</td>
                        <td className="num">{fmt(contribution)}%</td>
                      </tr>
                    );
                  })}

                  {npps.length === 0 && (
                    <tr>
                      <td colSpan={7}>
                        Không có dữ liệu NPP phù hợp với bộ lọc.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>


          {/* ======================================== */}
          {/* 7. CẢNH BÁO */}
          {/* ======================================== */}

          <section className="panel">
            <div
              className="actions"
              style={{ justifyContent: "space-between" }}
            >
              <h2 style={{ fontSize: 18, margin: 0 }}>
                5. Cảnh báo dữ liệu
              </h2>

              <span className="pill">
                {warnings.length} cảnh báo
              </span>
            </div>

            <p className="muted">
              Dữ liệu không có phân bổ không cộng vào KPI.
              Dữ liệu thiếu mã đơn hàng vẫn cộng tiền AK
              nếu có phân bổ nhưng không tính suất.
            </p>

            <div className="tablewrap">
              <table>
                <thead>
                  <tr>
                    <th>Miền</th>
                    <th>Vùng</th>
                    <th>Mã CT</th>
                    <th>Mã đơn hàng</th>
                    <th>NPP</th>
                    <th className="num">Tổng AK</th>
                    <th>Lý do</th>
                  </tr>
                </thead>

                <tbody>
                  {warnings.map((item, index) => (
                    <tr
                      key={`${item.areaName}-${item.programCode}-${item.orderCode}-${index}`}
                    >
                      <td>{item.regionName || "(Chưa xác định)"}</td>
                      <td>{item.areaName}</td>
                      <td>{item.programCode}</td>
                      <td>{item.orderCode || "(Trống)"}</td>
                      <td>{item.npp || "(Trống)"}</td>
                      <td className="num">{money(item.discount)}</td>
                      <td>{item.reason}</td>
                    </tr>
                  ))}

                  {warnings.length === 0 && (
                    <tr>
                      <td colSpan={7}>
                        Không có dữ liệu cảnh báo.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>

        </>
      )}


      <footer
        className="muted"
        style={{
          padding: "8px 2px 24px",
          lineHeight: 1.6
        }}
      >
        Dashboard theo dõi phân bổ chương trình.
        Hãy kiểm tra cảnh báo và đối chiếu số liệu
        với file hệ thống trước khi dùng báo cáo chính thức.
      </footer>

    </main>
  );
}
