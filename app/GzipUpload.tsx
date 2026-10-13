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
// XỬ LÝ TRƯỚC KHI GỬI LƯU BÁO CÁO
//
// 1. Bỏ toàn bộ cảnh báo (không gửi, không lưu).
// 2. Nén gzip để request nhỏ hơn giới hạn của Vercel.
//
// Các request khác không bị ảnh hưởng.
// ==================================================

export default function GzipUpload() {
  useEffect(() => {
    const originalFetch = window.fetch.bind(window);

    window.fetch = async (
      input: RequestInfo | URL,
      init?: RequestInit
    ): Promise<Response> => {
      let nextInit: RequestInit | undefined = init;

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
          typeof init?.body === "string";

        if (isSaveRequest) {
          const payload = JSON.parse(init!.body as string);

          // Chỉ xử lý request lưu báo cáo đầy đủ.
          if (payload && payload.report) {

            // Bỏ hết cảnh báo.
            payload.report.warnings = [];

            let body: string;

            if (typeof CompressionStream !== "undefined") {
              const data = await gzipToBase64(
                JSON.stringify({
                  allocation: payload.allocation,
                  programMap: payload.programMap,
                  report: payload.report
                })
              );

              body = JSON.stringify({
                month: payload.month,
                encoding: "gzip-base64",
                data
              });

            } else {
              // Trình duyệt không hỗ trợ nén:
              // vẫn gửi bình thường nhưng đã bỏ cảnh báo.
              body = JSON.stringify(payload);
            }

            nextInit = {
              ...init,
              body
            };
          }
        }

      } catch {
        // Nếu xử lý lỗi, gửi nguyên yêu cầu ban đầu.
        nextInit = init;
      }

      // Báo rõ nếu dữ liệu vẫn vượt giới hạn của Vercel.
      if (
        nextInit !== init &&
        typeof nextInit?.body === "string" &&
        nextInit.body.length > MAX_BODY_CHARS
      ) {
        throw new Error(
          "Dữ liệu sau khi nén vẫn quá lớn (" +
          (nextInit.body.length / 1024 / 1024).toFixed(1) +
          " MB, giới hạn khoảng 4,3 MB). " +
          "Cần chuyển sang gửi trực tiếp lên Apps Script."
        );
      }

      return originalFetch(input, nextInit);
    };

    return () => {
      window.fetch = originalFetch;
    };
  }, []);

  return null;
}
