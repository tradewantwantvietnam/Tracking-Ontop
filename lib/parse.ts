import * as XLSX from "xlsx";

export type Allocation = {
  regionName: string;
  areaName: string;
  programCode: string;
  programName: string;
  structure: string;
  allocation: number;
};

export type Actual = {
  npp: string;
  areaName: string;
  programCode: string;
  orderCode: string;
  discount: number;
};

function normalizeText(value: unknown): string {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ");
}


/**
 * Chuyển tiền chiết khấu và suất phân bổ về dạng số.
 *
 * Hỗ trợ dữ liệu số và dữ liệu text thông dụng,
 * ví dụ 150000, "150,000", "150.000".
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
      // Ví dụ: 1.234,56
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      // Ví dụ: 1,234.56
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
      // Ví dụ: 150.000 là 150 nghìn.
      text = text.replace(/\./g, "");
    } else if (parts.length > 2) {
      text = text.replace(/\./g, "");
    }
  }

  const result = Number(text);

  return Number.isFinite(result) ? result : 0;
}


/**
 * Đọc sheet đầu tiên của file Excel hoặc CSV.
 * Dòng đầu tiên được xem là tiêu đề.
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
    throw new Error("File không có sheet dữ liệu.");
  }

  const sheet = workbook.Sheets[sheetName];

  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: ""
  }) as unknown[][];
}


/**
 * FILE PHÂN BỔ
 *
 * A: Tên Miền
 * B: Tên vùng
 * C: Mã CT
 * D: Tên CT
 * E: Cơ cấu CT
 * F: Suất phân bổ
 */
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

    const programCode = normalizeText(row[2])
      .split(",")[0]
      .trim();

    const programName = normalizeText(row[3]);
    const structure = normalizeText(row[4]);
    const allocation = parseNumber(row[5]);

    // Bỏ qua dòng trống.
    if (
      !regionName &&
      !areaName &&
      !programCode
    ) {
      continue;
    }

    if (!areaName || !programCode) {
      throw new Error(
        `File phân bổ: dòng ${i + 1} thiếu Tên vùng hoặc Mã CT.`
      );
    }

    if (allocation < 0) {
      throw new Error(
        `File phân bổ: dòng ${i + 1} có suất phân bổ âm.`
      );
    }

    result.push({
      regionName,
      areaName,
      programCode,
      programName,
      structure,
      allocation
    });
  }

  if (result.length === 0) {
    throw new Error(
      "Không tìm thấy dòng phân bổ hợp lệ."
    );
  }

  return result;
}


/**
 * FILE 5.1
 *
 * E  = cột 5  = Tên NPP
 * G  = cột 7  = Tên vùng
 * N  = cột 14 = Mã CT
 * W  = cột 23 = Mã đơn hàng
 * AK = cột 37 = Tiền chiết khấu
 *
 * Các cột còn lại không được đưa vào dữ liệu xử lý.
 */
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

    // Chỉ lấy mã CT đầu tiên trước dấu phẩy.
    const programCode = normalizeText(row[13])
      .split(",")[0]
      .trim();

    const orderCode = normalizeText(row[22]);
    const discount = parseNumber(row[36]);

    // Bỏ qua dòng hoàn toàn trống trong các cột cần đọc.
    if (
      !npp &&
      !areaName &&
      !programCode &&
      !orderCode &&
      row[36] === ""
    ) {
      continue;
    }

    // Thiếu vùng hoặc mã CT thì bỏ khỏi tính toán.
    // Những trường hợp không khớp phân bổ sẽ được cảnh báo
    // nếu vẫn có đủ Tên vùng và Mã CT để đối chiếu.
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
      "Không tìm thấy dữ liệu hợp lệ tại các cột E, G, N, W, AK."
    );
  }

  return result;
}
