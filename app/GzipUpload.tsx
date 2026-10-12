"use client";

import { useEffect } from "react";


// Giới hạn của Vercel cho một request là khoảng 4,5 MB.
// Đặt ngưỡng an toàn thấp hơn một chút.
const MAX_BODY_CHARS = 4_300_000;


// ==================================================
// NÉN CHUỖI BẰNG GZIP RỒI ĐỔI SANG BASE64
// ==================================================

async function gzipToBase64(
  text: string
): Promise<string> {
  const stream = new Blob([text])
    .stream()
    .pipeThrough(new CompressionStream("gzip"));

  const buffer = await new Response(stream).arrayBuffer();
  const bytes = new Uint8Array(buffer);

  let binary = "";
  const step = 0x8000;

  for (let index = 0; index < bytes.length; index += step) {
    binary += String.fromCharCode.apply(
      null,
      Array.from(bytes.subarray(index, index + step))
    );
  }

  return btoa(binary);
}


// ==================================================
// TỰ ĐỘNG NÉN DỮ LIỆU KHI LƯU BÁO CÁO
//
// Khi trang gửi POST đến /api/save kèm báo cáo,
// dữ liệu được nén gzip trước khi gửi đi.
// Các request khác không bị ảnh hưởng.
// ==================================================

export default function GzipUpload() {
  useEffect(() => {
    const originalFetch = window.fetch.bind(window);

    window.fetch = async (
      input: RequestInfo | URL,
      init?: RequestInit
    ): Promise<Response> => {
      try {
        const rawUrl =
          typeof input === "string"
            ? input
            : input instanceof URL
            ? input.href
            : input.url;

        const url = new URL(
          rawUrl,
          window.location.origin
        );

        const isSaveRequest =
          url.pathname === "/api/save" &&
          String(init?.method || "GET").toUpperCase() ===
            "POST" &&
          typeof init?.body === "string" &&
          typeof CompressionStream !== "undefined";

        if (isSaveRequest) {
          const payload = JSON.parse(init!.body as string);

          // Chỉ nén request lưu báo cáo đầy đủ.
          if (payload && payload.report) {
            const data = await gzipToBase64(
              JSON.stringify({
                allocation: payload.allocation,
                programMap: payload.programMap,
                report: payload.report
              })
            );

            const body = JSON.stringify({
              month: payload.month,
              encoding: "gzip-base64",
              data
            });

            if (body.length > MAX_BODY_CHARS) {
              throw new Error(
                "Dữ liệu sau khi nén vẫn quá lớn (" +
                (body.length / 1024 / 1024).toFixed(1) +
                " MB, giới hạn khoảng 4,3 MB). " +
                "Cần chuyển sang gửi trực tiếp lên Apps Script."
              );
            }

            return originalFetch(input, {
              ...init,
              body
            });
          }
        }

      } catch (error) {
        // Lỗi dung lượng phải hiện cho người dùng.
        if (
          error instanceof Error &&
          error.message.startsWith("Dữ liệu sau khi nén")
        ) {
          throw error;
        }
      }

      return originalFetch(input, init);
    };

    return () => {
      window.fetch = originalFetch;
    };
  }, []);

  return null;
}
