import { NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";


// ==================================================
// KIỂM TRA CẤU HÌNH
// ==================================================

function getConfig() {
  const scriptUrl = process.env.APPS_SCRIPT_URL;
  const token = process.env.APPS_SCRIPT_TOKEN;

  if (!scriptUrl) {
    throw new Error(
      "Chưa cấu hình APPS_SCRIPT_URL trên Vercel."
    );
  }

  if (!token) {
    throw new Error(
      "Chưa cấu hình APPS_SCRIPT_TOKEN trên Vercel."
    );
  }

  return {
    scriptUrl,
    token
  };
}


// ==================================================
// GỌI GOOGLE APPS SCRIPT
// ==================================================

async function callAppsScript(
  params: Record<string, string>,
  body?: Record<string, unknown>
) {
  const config = getConfig();

  const url = new URL(config.scriptUrl);

  Object.entries(params).forEach(
    ([key, value]) => {
      url.searchParams.set(key, value);
    }
  );

  if (body) {
    const response = await fetch(url.toString(), {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify({
        ...body,
        token: config.token
      }),
      cache: "no-store",
      redirect: "follow"
    });

    const text = await response.text();

    let result: any;

    try {
      result = JSON.parse(text);
    } catch {
      throw new Error(
        "Apps Script không trả về JSON hợp lệ. " +
        "Hãy kiểm tra URL /exec và quyền truy cập."
      );
    }

    if (!response.ok || !result.ok) {
      throw new Error(
        result.error ||
        "Google Apps Script xử lý không thành công."
      );
    }

    return result;
  }

  url.searchParams.set("token", config.token);

  const response = await fetch(url.toString(), {
    method: "GET",
    cache: "no-store",
    redirect: "follow"
  });

  const text = await response.text();

  let result: any;

  try {
    result = JSON.parse(text);
  } catch {
    throw new Error(
      "Không đọc được phản hồi lịch sử từ Apps Script."
    );
  }

  if (!response.ok || !result.ok) {
    throw new Error(
      result.error ||
      "Không tải được dữ liệu Google Sheets."
    );
  }

  return result;
}


// ==================================================
// GET: LẤY DANH SÁCH THÁNG HOẶC LỊCH SỬ MỘT THÁNG
// ==================================================

export async function GET(request: Request) {
  try {
    const requestUrl = new URL(request.url);

    const action =
      requestUrl.searchParams.get("action") || "months";

    if (action === "months") {
      const result = await callAppsScript({
        action: "months"
      });

      return NextResponse.json(result);
    }

    if (action === "report") {
      const month =
        requestUrl.searchParams.get("month") || "";

      if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
        return NextResponse.json(
          {
            ok: false,
            error: "Tháng không hợp lệ."
          },
          { status: 400 }
        );
      }

      const result = await callAppsScript({
        action: "report",
        month
      });

      return NextResponse.json(result);
    }

    return NextResponse.json(
      {
        ok: false,
        error: "Action không được hỗ trợ."
      },
      { status: 400 }
    );

  } catch (error: any) {
    return NextResponse.json(
      {
        ok: false,
        error:
          error?.message ||
          "Không thể kết nối Apps Script."
      },
      { status: 500 }
    );
  }
}


// ==================================================
// POST: LƯU BÁO CÁO CỦA THÁNG ĐANG CHỌN
// ==================================================

export async function POST(request: Request) {
  try {
    const payload = await request.json();

    const month = String(payload.month || "");

    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
      return NextResponse.json(
        {
          ok: false,
          error:
            "Vui lòng chọn tháng hợp lệ trước khi lưu."
        },
        { status: 400 }
      );
    }

    if (
      !Array.isArray(payload.allocation) ||
      !payload.report
    ) {
      return NextResponse.json(
        {
          ok: false,
          error: "Thiếu dữ liệu phân bổ hoặc báo cáo."
        },
        { status: 400 }
      );
    }

    // Không gửi toàn bộ file 5.1 gốc.
    // Chỉ gửi dữ liệu phân bổ và báo cáo đã tổng hợp.
    // Nhờ vậy giảm dung lượng truyền dữ liệu.
    const result = await callAppsScript(
      {},
      {
        action: "saveReport",
        month,
        allocation: payload.allocation,
        report: payload.report
      }
    );

    return NextResponse.json({
      ok: true,
      month: result.month,
      updatedAt: result.updatedAt,
      months: result.months,
      counts: result.counts
    });

  } catch (error: any) {
    return NextResponse.json(
      {
        ok: false,
        error:
          error?.message ||
          "Không lưu được báo cáo."
      },
      { status: 500 }
    );
  }
}
