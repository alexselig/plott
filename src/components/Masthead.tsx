import Link from "next/link";

import PlottMark from "@/components/PlottMark";

/**
 * The Plott masthead: morph mark + serif wordmark, with optional right-side
 * content. `gallery` variant is the taller home header; `flow` is the slimmer
 * header used on the new-chart screens (logo links back to the gallery).
 */
export default function Masthead({
  variant = "flow",
  right,
  logoHref = "/",
}: {
  variant?: "gallery" | "flow";
  right?: React.ReactNode;
  logoHref?: string;
}) {
  const gallery = variant === "gallery";
  return (
    <header
      className="flex items-center justify-between border-b-[1.5px] border-ink"
      style={{ padding: gallery ? "22px 40px 18px" : "18px 40px" }}
    >
      <Link href={logoHref} className="flex items-center gap-3.5">
        <PlottMark size={42} />
        <span
          className="plott-serif"
          style={{
            fontSize: gallery ? 44 : 40,
            lineHeight: 0.8,
            letterSpacing: "-.01em",
          }}
        >
          Plott
        </span>
      </Link>
      {right}
    </header>
  );
}
