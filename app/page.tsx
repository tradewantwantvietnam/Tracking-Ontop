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


const money = (value: number) => {
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0
  }).format(value);
};


const fmt = (value: number) => {
  return new Intl.NumberFormat("vi-VN", {
    maximumFractionDigits: 2
  }).format(value);
};


function normalizeKey(value: string) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("vi");
}


function dataKey(areaName: string, programCode: string) {
  return (
    normalizeKey(areaName) +
    "|||" +
    normalizeKey(programCode)
  );
}


function uniqueValues(values: string[]) {
  return [...new Set(values.filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "vi"));
}


/**
 * TÍNH TOÁN BÁO CÁO
 *
 * 1. Mapping: Tên vùng + Mã CT.
 * 2. Mã đơn hàng duy nhất: Tên vùng + Mã CT + Mã đơn hàng.
 * 3. Một đơn hàng có ít nhất một AK > 0 thì tính 1 suất.
 * 4. Tổng chiết khấu cộng tất cả AK của các dòng trong đơn hàng.
 * 5. Không tìm thấy phân bổ: chuyển vào cảnh báo.
 * 6. NPP chỉ là cấp phân tích, không có định mức riêng.
 */
function buildReport(
  allocations: Allocation[],
  actuals: Actual[]
): Report {
  const allocationMap = new Map<string, Allocation>();

  // -----------------------------------------
  // A. KIỂM TRA VÀ ĐỌC FILE PHÂN BỔ
  // -----------------------------------------

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


  // -----------------------------------------
  // B. KHỞI TẠO TỔNG HỢP THEO VÙNG + MÃ CT
  // -----------------------------------------

  const summaryMap = new Map<string, Summary>();

  for (const allocation of allocations) {
    const key = dataKey(
      allocation.areaName,
      allocation.programCode
    );

    summaryMap.set(key, {
      regionName: allocation.regionName,
      areaName: allocation.areaName,
      programCode: allocation.programCode,
      programName: allocation.programName,
      structure: allocation.structure,
      allocation: allocation.allocation,
      used: 0,
      discount: 0,
      remaining: allocation.allocation,
      percent: 0
    });
  }


  // -----------------------------------------
  // C. NHÓM ĐƠN HÀNG CÓ MÃ W
  // -----------------------------------------

  const orderGroups = new Map<
    string,
    {
      allocation: Allocation;
      records: Actual[];
      discount: number;
      hasPositive: boolean;
    }
  >();


  // -----------------------------------------
  // D. KHỞI TẠO TỔNG HỢP THEO NPP
  // -----------------------------------------

  const nppMap = new Map<string, NppSummary>();

  function addNpp(
    allocation: Allocation,
    nppName: string,
    discountToAdd: number,
    usedToAdd: number
  ) {
    const name =
      String(nppName ?? "").trim() ||
      "(Không có tên NPP)";

    const key =
      dataKey(
        allocation.areaName,
        allocation.programCode
      ) +
      "|||" +
      normalizeKey(name);

    if (!nppMap.has(key)) {
      nppMap.set(key, {
        regionName: allocation.regionName,
        areaName: allocation.areaName,
        programCode: allocation.programCode,
        npp: name,
        used: 0,
        discount: 0
      });
    }

    const row = nppMap.get(key)!;

    row.discount += discountToAdd;
    row.used += usedToAdd;
  }


  // -----------------------------------------
  // E. NHÓM DỮ LIỆU KHÔNG CÓ PHÂN BỔ
  // -----------------------------------------

  const unmatchedWarnings = new Map<string, Warning>();
  const missingOrderWarnings: Warning[] = [];


  // -----------------------------------------
  // F. ĐỌC TỪNG DÒNG FILE 5.1
  // -----------------------------------------

  actuals.forEach((row, rowIndex) => {
    const key = dataKey(
      row.areaName,
      row.programCode
    );

    const allocation = allocationMap.get(key);

    // Không tồn tại cặp Vùng + CT trong file phân bổ.
    if (!allocation) {
      const orderCode = row.orderCode.trim();

      const warningKey = orderCode
        ? key + "|||" + normalizeKey(orderCode)
        : "MISSING_ORDER_ROW|||" + rowIndex;

      if (!unmatchedWarnings.has(warningKey)) {
        unmatchedWarnings.set(warningKey, {
          regionName: "",
          areaName: row.areaName,
          programCode: row.programCode,
          npp: row.npp,
          orderCode,
          discount: 0,
          reason:
            "Không có cặp Tên vùng + Mã CT trong file phân bổ"
        });
      }

      unmatchedWarnings.get(warningKey)!.discount +=
        row.discount;

      continue;
    }


    // Đã có phân bổ, nhưng mã đơn hàng bị trống.
    // Vẫn cộng tiền AK; không tự gộp các mã W trống
    // thành một đơn hàng vì sẽ gây sai số suất.
    if (!row.orderCode.trim()) {
      const summary = summaryMap.get(key)!;

      summary.discount += row.discount;

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


    // Khóa duy nhất:
    // Tên vùng + Mã CT + Mã đơn hàng.
    const orderKey =
      key +
      "|||" +
      normalizeKey(row.orderCode);

    if (!orderGroups.has(orderKey)) {
      orderGroups.set(orderKey, {
        allocation,
        records: [],
        discount: 0,
        hasPositive: false
      });
    }

    const group = orderGroups.get(orderKey)!;

    group.records.push(row);

    // Tổng AK của tất cả dòng trong đơn hàng.
    group.discount += row.discount;

    // Chỉ cần một dòng AK > 0 thì đơn hàng
    // đủ điều kiện tính một suất.
    if (row.discount > 0) {
      group.hasPositive = true;
    }
  });


  // -----------------------------------------
  // G. TÍNH SUẤT VÀ TIỀN THEO ĐƠN HÀNG
  // -----------------------------------------

  for (const group of orderGroups.values()) {
    const allocation = group.allocation;

    const key = dataKey(
      allocation.areaName,
      allocation.programCode
    );

    const summary = summaryMap.get(key)!;

    // Tổng tiền chiết khấu.
    // Mỗi dòng trong group đã được cộng đúng một lần.
    summary.discount += group.discount;

    // Một mã CT + vùng + đơn hàng chỉ có tối đa 1 suất.
    if (group.hasPositive) {
      summary.used += 1;
    }


    // Phân bổ tiền theo từng NPP.
    const nppDiscounts = new Map<string, number>();

    for (const row of group.records) {
      const nppName =
        row.npp.trim() ||
        "(Không có tên NPP)";

      nppDiscounts.set(
        nppName,
        (nppDiscounts.get(nppName) || 0) +
        row.discount
      );
    }


    // Nếu cùng một đơn hàng có nhiều NPP,
    // chỉ giao suất cho một NPP để không làm
    // tổng suất NPP vượt tổng suất cấp vùng.
    //
    // NPP nhận suất là NPP đầu tiên xuất hiện
    // trên dòng AK > 0 của đơn hàng đó.
    const firstPositiveRow = group.records.find(
      row => row.discount > 0
    );

    const slotOwner =
      firstPositiveRow?.npp.trim() ||
      "(Không có tên NPP)";


    for (const [nppName, discount] of nppDiscounts) {
      const usedToAdd =
        group.hasPositive &&
        nppName === slotOwner
          ? 1
          : 0;

      addNpp(
        allocation,
        nppName,
        discount,
        usedToAdd
      );
    }
  }


  // -----------------------------------------
  // H. HOÀN TẤT KPI VÀ PHẦN TRĂM
  // -----------------------------------------

  const summaries = [...summaryMap.values()].map(
    row => {
      const remaining =
        row.allocation - row.used;

      const percent =
        row.allocation > 0
          ? (row.used / row.allocation) * 100
          : 0;

      return {
        ...row,
        remaining,
        percent
      };
    }
  );


  // -----------------------------------------
  // I. DANH SÁCH CẢNH BÁO
  // -----------------------------------------

  const warnings = [
    ...unmatchedWarnings.values(),
    ...missingOrderWarnings
  ];

  return {
    summaries,
    npps: [...nppMap.values()],
    warnings
  };
}


/**
 * GIAO DIỆN DASHBOARD
 */
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


  // Bộ lọc Dashboard.
  const [program, setProgram] = useState("ALL");
  const [region, setRegion] = useState("ALL");
  const [area, setArea] = useState("ALL");


  // -----------------------------------------
  // DANH SÁCH LỰA CHỌN TRONG BỘ LỌC
  // -----------------------------------------

  const programs = useMemo(
    () =>
      uniqueValues(
        (report?.summaries || []).map(
          row => row.programCode
        )
      ),
    [report]
  );


  const regions = useMemo(
    () =>
      uniqueValues(
        (report?.summaries || [])
          .filter(
            row =>
              program === "ALL" ||
              row.programCode === program
          )
          .map(row => row.regionName)
      ),
    [report, program]
  );


  const areas = useMemo(
    () =>
      uniqueValues(
        (report?.summaries || [])
          .filter(
            row =>
              (program === "ALL" ||
                row.programCode === program) &&
              (region === "ALL" ||
                row.regionName === region)
          )
          .map(row => row.areaName)
      ),
    [report, program, region]
  );


  // Đưa bộ lọc về ALL nếu lựa chọn hiện tại
  // không còn nằm trong danh sách hợp lệ.
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


  // -----------------------------------------
  // DỮ LIỆU ĐÃ LỌC
  // -----------------------------------------

  const filtered = useMemo(
    () =>
      (report?.summaries || []).filter(
        row =>
          (program === "ALL" ||
            row.programCode === program) &&
          (region === "ALL" ||
            row.regionName === region) &&
          (area === "ALL" ||
            row.areaName === area)
      ),
    [report, program, region, area]
  );


  // -----------------------------------------
  // TÍNH KPI TỪ DỮ LIỆU ĐÃ LỌC
  // -----------------------------------------

  const total = useMemo(() => {
    return filtered.reduce(
      (sum, row) => ({
        allocation:
          sum.allocation + row.allocation,

        used:
          sum.used + row.used,

        discount:
          sum.discount + row.discount
      }),
      {
        allocation: 0,
        used: 0,
        discount: 0
      }
    );
  }, [filtered]);


  const percent =
    total.allocation > 0
      ? (total.used / total.allocation) * 100
      : 0;


  // -----------------------------------------
  // DỮ LIỆU NPP THEO BỘ LỌC
  // -----------------------------------------

  const npps = useMemo(
    () =>
      (report?.npps || [])
        .filter(
          row =>
            (program === "ALL" ||
              row.programCode === program) &&
            (region === "ALL" ||
              row.regionName === region) &&
            (area === "ALL" ||
              row.areaName === area)
        )
        .sort((a, b) => b.used - a.used),
    [report, program, region, area]
  );


  // -----------------------------------------
  // DỮ LIỆU CẢNH BÁO
  // -----------------------------------------

  const warnings = useMemo(
    () =>
      (report?.warnings || []).filter(
        row =>
          (program === "ALL" ||
            row.programCode === program) &&
          (region === "ALL" ||
            row.regionName === region) &&
          (area === "ALL" ||
            row.areaName === area)
      ),
    [report, program, region, area]
  );


  // -----------------------------------------
  // BIỂU ĐỒ THEO CHƯƠNG TRÌNH
  // -----------------------------------------

  const byProgram = useMemo(() => {
    const map = new Map<
      string,
      {
        name: string;
        allocated: number;
        used: number;
      }
    >();

    for (const row of filtered) {
      if (!map.has(row.programCode)) {
        map.set(row.programCode, {
          name: row.programCode,
          allocated: 0,
          used: 0
        });
      }

      const item = map.get(row.programCode)!;

      item.allocated += row.allocation;
      item.used += row.used;
    }

    return [...map.values()];
  }, [filtered]);


  // -----------------------------------------
  // BIỂU ĐỒ THEO VÙNG
  // -----------------------------------------

  const byArea = useMemo(() => {
    const map = new Map<
      string,
      {
        name: string;
        allocated: number;
        used: number;
      }
    >();

    for (const row of filtered) {
      if (!map.has(row.areaName)) {
        map.set(row.areaName, {
          name: row.areaName,
          allocated: 0,
          used: 0
        });
      }

      const item = map.get(row.areaName)!;

      item.allocated += row.allocation;
      item.used += row.used;
    }

    return [...map.values()]
      .sort((a, b) => b.used - a.used)
      .slice(0, 20);
  }, [filtered]);


  // -----------------------------------------
  // UPLOAD FILE
  // -----------------------------------------

  async function loadFile(
    type: "allocation" | "actual",
    file?: File
  ) {
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


  // -----------------------------------------
  // XỬ LÝ VÀ LƯU BÁO CÁO
  // -----------------------------------------

  async function processData() {
    if (!allocation || !actual) {
      setStatus(
        "Vui lòng upload đủ file phân bổ và file 5.1."
      );

      setStatusType("error");

      return;
    }

    setBusy(true);
    setStatus("");

    try {
      // Tính toán ngay trên trình duyệt.
      const result = buildReport(
        allocation,
        actual
      );

      setReport(result);

      // Gửi kết quả sang API nội bộ của Vercel.
      // API này mới gọi Apps Script ở phía máy chủ.
      try {
        const response = await fetch("/api/save", {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            allocation,
            actual,
            report: result
          })
        });

        const resultFromServer =
          await response.json();

        if (
          !response.ok ||
          !resultFromServer.ok
        ) {
          throw new Error(
            resultFromServer.error ||
            "Không lưu được vào Google Sheets."
          );
        }

        setSavedAt(
          new Date().toLocaleString("vi-VN")
        );

        setStatus(
          "Đã đối chiếu dữ liệu và lưu kết quả vào Google Sheets."
        );

        setStatusType("success");

      } catch (saveError: any) {
        // Dashboard vẫn hiển thị kết quả tính toán,
        // đồng thời thông báo nếu chưa lưu được.
        setStatus(
          "Dashboard đã tính toán, nhưng chưa lưu được " +
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


  // -----------------------------------------
  // XÓA BỘ LỌC
  // -----------------------------------------

  function resetFilters() {
    setProgram("ALL");
    setRegion("ALL");
    setArea("ALL");
  }


  // -----------------------------------------
  // GIAO DIỆN
  // -----------------------------------------

  return (
    <main className="wrap">

      <header className="header">
        <div>
          <h1
            style={{
              margin: "0 0 6px",
              fontSize: 28
            }}
          >
            Theo dõi phân bổ chương trình
          </h1>

          <div className="muted">
            Đối chiếu phân bổ · Suất đã dùng ·
            Tiền chiết khấu · Phân tích NPP
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


      {/* =================================== */}
      {/* 1. UPLOAD */}
      {/* =================================== */}

      <section className="panel">
        <h2
          style={{
            marginTop: 0,
            fontSize: 18
          }}
        >
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
              onChange={event =>
                loadFile(
                  "allocation",
                  event.target.files?.[0]
                )
              }
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
              Chỉ lấy E (NPP), G (Vùng), N (Mã CT),
              W (Mã đơn hàng), AK (Chiết khấu).
            </p>

            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              disabled={busy}
              onChange={event =>
                loadFile(
                  "actual",
                  event.target.files?.[0]
                )
              }
            />

            {actualName && (
              <p className="success muted">
                Đã chọn: {actualName}
              </p>
            )}
          </div>

        </div>


        <div
          className="actions"
          style={{ marginTop: 14 }}
        >
          <button
            className="btn"
            disabled={
              busy ||
              !allocation ||
              !actual
            }
            onClick={processData}
          >
            {busy
              ? "Đang xử lý..."
              : "Đối chiếu và cập nhật Dashboard"}
          </button>
        </div>


        {status && (
          <div
            className={`status ${statusType}`}
          >
            {status}
          </div>
        )}

        {savedAt && (
          <div
            className="muted"
            style={{ marginTop: 6 }}
          >
            Lần lưu gần nhất: {savedAt}
          </div>
        )}

        <div className="notice">
          <strong>Quy tắc xử lý:</strong>

          <br />

          Mã CT lấy phần đầu trước dấu phẩy.

          <br />

          Mỗi Mã CT + Vùng + Mã đơn hàng chỉ tính
          tối đa một suất nếu có ít nhất một dòng
          AK lớn hơn 0.

          <br />

          Tổng tiền chiết khấu cộng tất cả dòng AK
          của các đơn hàng tương ứng.

          <br />

          Dữ liệu không có phân bổ được đưa vào
          cảnh báo và không cộng vào KPI phân bổ.
        </div>
      </section>


      {!report && (
        <section className="panel">
          <h2
            style={{
              marginTop: 0,
              fontSize: 18
            }}
          >
            Dashboard chưa có dữ liệu
          </h2>

          <p className="muted">
            Vui lòng upload hai file và nhấn
            “Đối chiếu và cập nhật Dashboard”.
          </p>
        </section>
      )}


      {report && (
        <>

          {/* ================================= */}
          {/* 2. BỘ LỌC */}
          {/* ================================= */}

          <section className="panel">
            <h2
              style={{
                margin: "0 0 6px",
                fontSize: 18
              }}
            >
              2. Bộ lọc Dashboard
            </h2>

            <div className="filters">

              <label>
                Chương trình

                <select
                  value={program}
                  onChange={event =>
                    setProgram(event.target.value)
                  }
                >
                  <option value="ALL">
                    Tất cả chương trình
                  </option>

                  {programs.map(value => (
                    <option
                      key={value}
                      value={value}
                    >
                      {value}
                    </option>
                  ))}
                </select>
              </label>


              <label>
                Miền

                <select
                  value={region}
                  onChange={event =>
                    setRegion(event.target.value)
                  }
                >
                  <option value="ALL">
                    Tất cả miền
                  </option>

                  {regions.map(value => (
                    <option
                      key={value}
                      value={value}
                    >
                      {value}
                    </option>
                  ))}
                </select>
              </label>


              <label>
                Vùng

                <select
                  value={area}
                  onChange={event =>
                    setArea(event.target.value)
                  }
                >
                  <option value="ALL">
                    Tất cả vùng
                  </option>

                  {areas.map(value => (
                    <option
                      key={value}
                      value={value}
                    >
                      {value}
                    </option>
                  ))}
                </select>
              </label>

            </div>

            <div className="muted">
              Các KPI, biểu đồ và bảng bên dưới
              thay đổi theo bộ lọc.
            </div>
          </section>


          {/* ================================= */}
          {/* 3. KPI */}
          {/* ================================= */}

          <div className="grid">

            <div className="metric">
              <div className="label">
                Tổng phân bổ
              </div>

              <div className="value">
                {fmt(total.allocation)} suất
              </div>
            </div>


            <div className="metric">
              <div className="label">
                Suất đã dùng
              </div>

              <div className="value">
                {fmt(total.used)} suất
              </div>
            </div>


            <div className="metric">
              <div className="label">
                Còn lại / vượt phân bổ
              </div>

              <div
                className="value"
                style={{
                  color:
                    total.allocation - total.used < 0
                      ? "#b42318"
                      : "inherit"
                }}
              >
                {fmt(
                  total.allocation - total.used
                )} suất
              </div>
            </div>


            <div className="metric">
              <div className="label">
                Tỷ lệ sử dụng
              </div>

              <div className="value">
                {fmt(percent)}%
              </div>

              <div
                className="progress"
                style={{ marginTop: 10 }}
              >
                <div
                  style={{
                    width:
                      `${Math.max(
                        0,
                        Math.min(100, percent)
                      )}%`
                  }}
                />
              </div>
            </div>


            <div className="metric">
              <div className="label">
                Tổng tiền chiết khấu
              </div>

              <div
                className="value"
                style={{ fontSize: 21 }}
              >
                {money(total.discount)}
              </div>
            </div>

          </div>


          {/* ================================= */}
          {/* 4. BIỂU ĐỒ */}
          {/* ================================= */}

          <div
            className="chartgrid"
            style={{ marginTop: 18 }}
          >

            <section className="panel chart">
              <h3>
                Phân bổ và đã dùng theo chương trình
              </h3>

              <ResponsiveContainer
                width="100%"
                height="88%"
              >
                <BarChart data={byProgram}>
                  <CartesianGrid
                    strokeDasharray="3 3"
                  />

                  <XAxis
                    dataKey="name"
                    tick={{ fontSize: 10 }}
                  />

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
              <h3>
                Suất đã dùng theo vùng
              </h3>

              <ResponsiveContainer
                width="100%"
                height="88%"
              >
                <BarChart
                  data={byArea}
                  layout="vertical"
                  margin={{ left: 20 }}
                >
                  <CartesianGrid
                    strokeDasharray="3 3"
                  />

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


          {/* ================================= */}
          {/* 5. BẢNG CHI TIẾT VÙNG */}
          {/* ================================= */}

          <section className="panel">
            <h2
              style={{
                fontSize: 18,
                marginTop: 0
              }}
            >
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
                    <th className="num">
                      Phân bổ
                    </th>
                    <th className="num">
                      Đã dùng
                    </th>
                    <th className="num">
                      Còn lại
                    </th>
                    <th className="num">
                      % dùng
                    </th>
                    <th className="num">
                      Tổng chiết khấu
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {filtered.map((row, index) => (
                    <tr
                      key={
                        row.areaName +
                        row.programCode +
                        index
                      }
                    >
                      <td>{row.regionName}</td>
                      <td>{row.areaName}</td>
                      <td>{row.programCode}</td>
                      <td>{row.programName}</td>

                      <td className="num">
                        {fmt(row.allocation)}
                      </td>

                      <td className="num">
                        {fmt(row.used)}
                      </td>

                      <td className="num">
                        {fmt(row.remaining)}
                      </td>

                      <td className="num">
                        {fmt(row.percent)}%
                      </td>

                      <td className="num">
                        {money(row.discount)}
                      </td>
                    </tr>
                  ))}

                  {filtered.length === 0 && (
                    <tr>
                      <td colSpan={9}>
                        Không có dữ liệu phù hợp
                        với bộ lọc.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>


          {/* ================================= */}
          {/* 6. BẢNG CHI TIẾT NPP */}
          {/* ================================= */}

          <section className="panel">
            <h2
              style={{
                fontSize: 18,
                marginTop: 0
              }}
            >
              4. Chi tiết sử dụng theo NPP
            </h2>

            <p className="muted">
              NPP không có suất phân bổ riêng.
              Số suất và tiền chiết khấu phản ánh
              phần đóng góp của NPP vào tổng sử dụng
              của vùng và chương trình.
            </p>

            <div className="tablewrap">
              <table>
                <thead>
                  <tr>
                    <th>Miền</th>
                    <th>Vùng</th>
                    <th>Mã CT</th>
                    <th>NPP</th>
                    <th className="num">
                      Suất đã dùng
                    </th>
                    <th className="num">
                      Tổng chiết khấu
                    </th>
                    <th className="num">
                      Tỷ trọng suất trong vùng
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {npps.map((row, index) => {
                    const parent = filtered.find(
                      item =>
                        item.areaName === row.areaName &&
                        item.programCode === row.programCode
                    );

                    const contribution =
                      parent && parent.used > 0
                        ? row.used / parent.used * 100
                        : 0;

                    return (
                      <tr
                        key={
                          row.areaName +
                          row.programCode +
                          row.npp +
                          index
                        }
                      >
                        <td>{row.regionName}</td>
                        <td>{row.areaName}</td>
                        <td>{row.programCode}</td>
                        <td>{row.npp}</td>

                        <td className="num">
                          {fmt(row.used)}
                        </td>

                        <td className="num">
                          {money(row.discount)}
                        </td>

                        <td className="num">
                          {fmt(contribution)}%
                        </td>
                      </tr>
                    );
                  })}

                  {npps.length === 0 && (
                    <tr>
                      <td colSpan={7}>
                        Không có dữ liệu NPP phù hợp
                        với bộ lọc.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>


          {/* ================================= */}
          {/* 7. CẢNH BÁO */}
          {/* ================================= */}

          <section className="panel">
            <div
              className="actions"
              style={{
                justifyContent: "space-between"
              }}
            >
              <h2
                style={{
                  fontSize: 18,
                  margin: 0
                }}
              >
                5. Cảnh báo dữ liệu
              </h2>

              <span className="pill">
                {warnings.length} cảnh báo
              </span>
            </div>

            <p className="muted">
              Cảnh báo gồm dữ liệu không có phân bổ
              và dữ liệu thiếu mã đơn hàng. Các dòng
              không có phân bổ không được cộng vào
              KPI. Dòng thiếu mã đơn hàng không được
              tính suất để tránh đếm trùng sai.
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
                    <th className="num">
                      Tổng AK
                    </th>
                    <th>Lý do</th>
                  </tr>
                </thead>

                <tbody>
                  {warnings.map((row, index) => (
                    <tr
                      key={
                        row.areaName +
                        row.programCode +
                        row.orderCode +
                        index
                      }
                    >
                      <td>{row.regionName}</td>
                      <td>{row.areaName}</td>
                      <td>{row.programCode}</td>
                      <td>
                        {row.orderCode || "(Trống)"}
                      </td>
                      <td>
                        {row.npp || "(Trống)"}
                      </td>

                      <td className="num">
                        {money(row.discount)}
                      </td>

                      <td>{row.reason}</td>
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
        Dashboard phân bổ chương trình.
        Vui lòng kiểm tra mục Cảnh báo dữ liệu
        và đối chiếu kết quả trước khi dùng làm
        báo cáo chính thức.
      </footer>

    </main>
  );
}
