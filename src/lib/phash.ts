import { serializeSvg } from "@/lib/export/svg";

/**
 * Perceptual hashing (dHash) so a re-encoded / pasted / screenshotted chart
 * image can still be matched back to its chart. 9x8 grayscale => 64-bit hash
 * (16 hex chars); compare with Hamming distance.
 */

const W = 9;
const H = 8;

/** Hamming distance between two equal-length hex strings (bit differences). */
export function hammingHex(a: string, b: string): number {
  if (a.length !== b.length) return Math.max(a.length, b.length) * 4;
  let d = 0;
  for (let i = 0; i < a.length; i++) {
    let x = parseInt(a[i], 16) ^ parseInt(b[i], 16);
    while (x) {
      d += x & 1;
      x >>= 1;
    }
  }
  return d;
}

/** dHash of a row-major grayscale matrix: compare horizontally adjacent pixels. */
export function dhashFromGray(gray: number[], w = W, h = H): string {
  let bits = "";
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w - 1; x++) {
      bits += gray[y * w + x] < gray[y * w + x + 1] ? "1" : "0";
    }
  }
  let hex = "";
  for (let i = 0; i < bits.length; i += 4) {
    hex += parseInt(bits.slice(i, i + 4).padEnd(4, "0"), 2).toString(16);
  }
  return hex;
}

function toGray(source: CanvasImageSource): number[] {
  const c = document.createElement("canvas");
  c.width = W;
  c.height = H;
  const ctx = c.getContext("2d");
  if (!ctx) return new Array(W * H).fill(0);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, W, H);
  ctx.drawImage(source, 0, 0, W, H);
  const data = ctx.getImageData(0, 0, W, H).data;
  const gray: number[] = [];
  for (let i = 0; i < W * H; i++) {
    gray.push(0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2]);
  }
  return gray;
}

function loadImg(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Failed to load image"));
    img.src = url;
  });
}

export function dhashFromImage(img: HTMLImageElement): string {
  return dhashFromGray(toGray(img));
}

/** Hash the currently-rendered chart SVG (used when exporting). */
export async function dhashFromSvg(svg: SVGSVGElement): Promise<string> {
  const blob = new Blob([serializeSvg(svg)], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  try {
    return dhashFromImage(await loadImg(url));
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Hash an uploaded image file (used when reopening). */
export async function dhashFromFile(file: File): Promise<string> {
  const url = URL.createObjectURL(file);
  try {
    return dhashFromImage(await loadImg(url));
  } finally {
    URL.revokeObjectURL(url);
  }
}
