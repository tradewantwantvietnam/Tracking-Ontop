import { NextResponse } from "next/server";

export const runtime = "nodejs";


export async function POST(request: Request) {
  try {
    const scriptUrl = process.env.APPS_SCRIPT_URL;
    const token = process.env.APPS_SCRIPT_TOKEN;

    if (!scriptUrl) {
      return NextResponse.json(
        {
          ok: false,
          error:
            "Chưa cấu hình APPS_SCRIPT_URL trên Vercel."
        },
        { status: 503 }
      );
    }

    if (!token) {
      return NextResponse.json(
        {
          ok: false,
          error:
            "Chưa cấu hình APPS_SCRIPT_TOKEN trên Vercel."
        },
        { status: 503 }
      );
    }

    const payload = await request.json();

    if (
      !Array.isArray(payload.allocation) ||
      !Array.isArray(payload.actual) ||
      !payload.report
    ) {
      return NextResponse.json(
        {
          ok: false,
          error: "Dữ liệu gửi lên không đầy đủ."
        },
        { status: 400 }
      );
    }

    // Token được thêm tại phía máy chủ.
    // Không nhận token trực tiếp từ trình duyệt.
    const requestBody = {
      action: "saveReport",
      allocation: payload.allocation,
      actual: payload.actual,
      report: payload.report,
      token
    };

    const response = await fetch(scriptUrl, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(requestBody),
      cache: "no-store",
      redirect: "follow"
    });

    const responseText = await response.text();

    let result: any;

    try {
      result = JSON.parse(responseText);
    } catch {
      throw new Error(
        "Apps Script không trả về JSON hợp lệ. " +
        "Hãy kiểm tra URL /exec và quyền truy cập."
      );
    }

    if (!result.ok) {
      throw new Error(
        result.error || "Apps Script xử lý không thành công."
      );
    }

    return NextResponse.json({
      ok: true,
      savedAt: result.savedAt,
      counts: result.counts
    });

  } catch (error: any) {
    return NextResponse.json(
      {
        ok: false,
        error:
          error?.message ||
          "Không thể kết nối Google Apps Script."
      },
      { status: 500 }
    );
  }
}
