import * as XLSX from "xlsx";


// ==================================================
// KIỂU DỮ LIỆU FILE PHÂN BỔ — 5 CỘT
// ==================================================

export type Allocation = {
  regionName: string;
  areaName: string;
  programName: string;
  structure: string;
  allocation: number;
};


// ==================================================
// KIỂU DỮ LIỆU FILE TÊN CT — 2 CỘT
// ==================================================

export type ProgramMapping = {
  programCode: string;
  programName: string;
};


// ==================================================
// KIỂU DỮ LIỆU FILE 5.1 — CHỈ 5 TRƯỜNG
// ==================================================

export type Actual = {
  npp: string;
  areaName: string;
  programCode: string;
  orderCode: string;
  discount: number;
};


// ==================================================
// HÀM HỖ TRỢ
// ==================================================

function normalizeText(value: unknown): string {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ");
}


/**
 * Đọc sheet đầu tiên.
 * Dòng đầu tiên được xem là tiêu đề.
 * raw:false giúp ưu tiên giá trị hiển thị trong Excel.
 */
async function readFirstSheet(
  file: File
): Promise<unknown[][]> {
  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: false
  });

  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error(
      `File "${file.name}" không có sheet dữ liệu.`
    );
  }

  const sheet = workbook.Sheets[sheetName];

  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: ""
  }) as unknown[][];
}


/**
 * Chuyển số, tiền tệ và số dạng text về number.
 *
 * Hỗ trợ ví dụ:
 * 150000
 * "150.000"
 * "1.500.000"
 * "150,000"
 */
function parseNumber(value: unknown): number {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  let text = normalizeText(value);

  if (!text) {
    return 0;
  }

  text = text
    .replace(/[₫đ\s]/gi, "")
    .replace(/[^\d,.-]/g, "");

  if (!text) {
    return 0;
  }

  if (text.includes(",") && text.includes(".")) {
    if (text.lastIndexOf(",") > text.lastIndexOf(".")) {
      // 1.234,56
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      // 1,234.56
      text = text.replace(/,/g, "");
    }
  } else if (text.includes(",")) {
    const parts = text.split(",");

    if (
      parts.length === 2 &&
      parts[1].length <= 2
    ) {
      text =
        parts[0].replace(/\./g, "") +
        "." +
        parts[1];
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (text.includes(".")) {
    const parts = text.split(".");

    if (
      parts.length === 2 &&
      parts[1].length === 3
    ) {
      // 150.000
      text = text.replace(/\./g, "");
    } else if (parts.length > 2) {
      text = text.replace(/\./g, "");
    }
  }

  const result = Number(text);

  return Number.isFinite(result) ? result : 0;
}


// ==================================================
// FILE 1: PHÂN BỔ — 5 CỘT
//
// A: Tên Miền
// B: Tên vùng
// C: Tên CT
// D: Cơ cấu CT
// E: Suất phân bổ
// ==================================================

export async function parseAllocation(
  file: File
): Promise<Allocation[]> {
  const rows = await readFirstSheet(file);

  if (rows.length < 2) {
    throw new Error(
      "File phân bổ cần có tiêu đề và dữ liệu."
    );
  }

  const result: Allocation[] = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] as unknown[];

    const regionName = normalizeText(row[0]);
    const areaName = normalizeText(row[1]);
    const programName = normalizeText(row[2]);
    const structure = normalizeText(row[3]);
    const allocation = parseNumber(row[4]);

    // Bỏ qua dòng trống.
    if (
      !regionName &&
      !areaName &&
      !programName &&
      !structure &&
      !normalizeText(row[4])
    ) {
      continue;
    }

    if (!regionName || !areaName || !programName) {
      throw new Error(
        `File phân bổ, dòng ${i + 1}: ` +
        "thiếu Tên Miền, Tên vùng hoặc Tên CT."
      );
    }

    if (allocation < 0) {
      throw new Error(
        `File phân bổ, dòng ${i + 1}: ` +
        "Suất phân bổ không được âm."
      );
    }

    result.push({
      regionName,
      areaName,
      programName,
      structure,
      allocation
    });
  }

  if (result.length === 0) {
    throw new Error(
      "Không tìm thấy dữ liệu phân bổ hợp lệ."
    );
  }

  return result;
}


// ==================================================
// FILE 2: TÊN CT — 2 CỘT
//
// A: Mã CT
// B: Tên CT
//
// Cho phép nhiều mã CT cùng thuộc một Tên CT.
// Mỗi mã CT chỉ được gán một Tên CT.
// ==================================================

export async function parseProgramMapping(
  file: File
): Promise<ProgramMapping[]> {
  const rows = await readFirstSheet(file);

  if (rows.length < 2) {
    throw new Error(
      "File Tên CT cần có tiêu đề và dữ liệu."
    );
  }

  const result: ProgramMapping[] = [];

  const mappingByCode = new Map<string, string>();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] as unknown[];

    const programCode = normalizeText(row[0]);
    const programName = normalizeText(row[1]);

    if (!programCode && !programName) {
      continue;
    }

    if (!programCode || !programName) {
      throw new Error(
        `File Tên CT, dòng ${i + 1}: ` +
        "phải có đủ Mã CT và Tên CT."
      );
    }

    // Mã trong file quy ước phải là một mã đơn,
    // không để nhiều mã chung trong một ô.
    if (programCode.includes(",")) {
      throw new Error(
        `File Tên CT, dòng ${i + 1}: ` +
        "mỗi dòng chỉ được chứa một Mã CT."
      );
    }

    const codeKey = programCode.toLocaleLowerCase("vi");
    const nameKey = programName.toLocaleLowerCase("vi");

    if (mappingByCode.has(codeKey)) {
      const existingName = mappingByCode.get(codeKey)!;

      if (
        existingName.toLocaleLowerCase("vi") !== nameKey
      ) {
        throw new Error(
          `Mã CT "${programCode}" đang được gán cho ` +
          "nhiều Tên CT khác nhau. Hãy kiểm tra file Tên CT."
        );
      }

      // Nếu trùng mã và cùng tên thì bỏ qua dòng lặp.
      continue;
    }

    mappingByCode.set(codeKey, programName);

    result.push({
      programCode,
      programName
    });
  }

  if (result.length === 0) {
    throw new Error(
      "Không tìm thấy quy ước Mã CT / Tên CT hợp lệ."
    );
  }

  return result;
}


// ==================================================
// FILE 3: 5.1
//
// E  = Tên NPP
// G  = Tên vùng
// N  = Mã CT
// W  = Mã đơn hàng
// AK = Tiền chiết khấu
//
// Chỉ đọc 5 trường trên.
// ==================================================

export async function parseActual(
  file: File
): Promise<Actual[]> {
  const rows = await readFirstSheet(file);

  if (rows.length < 2) {
    throw new Error(
      "File 5.1 cần có tiêu đề và dữ liệu."
    );
  }

  const result: Actual[] = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] as unknown[];

    const npp = normalizeText(row[4]);
    const areaName = normalizeText(row[6]);

    const fullCode = normalizeText(row[13]);

    // Lấy mã đầu tiên trước dấu phẩy.
    const programCode = fullCode
      .split(",")[0]
      .trim();

    const orderCode = normalizeText(row[22]);
    const discount = parseNumber(row[36]);

    // Bỏ qua dòng không có dữ liệu ở các cột cần dùng.
    if (
      !npp &&
      !areaName &&
      !fullCode &&
      !orderCode &&
      !normalizeText(row[36])
    ) {
      continue;
    }

    // Không thể đối chiếu khi thiếu vùng hoặc mã CT.
    if (!areaName || !programCode) {
      continue;
    }

    result.push({
      npp,
      areaName,
      programCode,
      orderCode,
      discount
    });
  }

  if (result.length === 0) {
    throw new Error(
      "Không tìm thấy dữ liệu hợp lệ ở các cột E, G, N, W, AK."
    );
  }

  return result;
}
