import { ImageResponse } from "next/og";

export const size = {
  width: 32,
  height: 32
};

export const contentType = "image/png";

export default function Icon(): ImageResponse {
  return new ImageResponse(
    (
      <div
        style={{
          width: "100%",
          height: "100%",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: "#ffffff"
        }}
      >
        <div
          style={{
            position: "relative",
            width: 24,
            height: 24,
            display: "flex",
            alignItems: "center",
            justifyContent: "center"
          }}
        >
          <div
            style={{
              position: "absolute",
              inset: 0,
              borderRadius: 8,
              background: "linear-gradient(135deg, #3b82f6, #6366f1)",
              opacity: 0.25,
              transform: "rotate(12deg)"
            }}
          />
          <div
            style={{
              width: 18,
              height: 18,
              borderRadius: 6,
              background: "linear-gradient(135deg, #2563eb, #4f46e5)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center"
            }}
          >
            <div
              style={{
                width: 4,
                height: 4,
                borderRadius: 9999,
                background: "#ffffff"
              }}
            />
          </div>
        </div>
      </div>
    ),
    size
  );
}
