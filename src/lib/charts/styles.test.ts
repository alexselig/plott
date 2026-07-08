import { describe, expect, it } from "vitest";

import {
  applyPalette,
  applyStyle,
  defaultChartStyle,
  PALETTE_ORDER,
  PALETTES,
  STYLE_ORDER,
  STYLES,
} from "@/lib/charts/styles";

describe("Plott styles + palettes", () => {
  it("has 12 styles and 8 palettes", () => {
    expect(STYLE_ORDER.length).toBe(12);
    expect(PALETTE_ORDER.length).toBe(8);
    STYLE_ORDER.forEach((k) => expect(STYLES[k]).toBeTruthy());
    PALETTE_ORDER.forEach((k) => expect(PALETTES[k]).toBeTruthy());
  });

  it("default style is Newsprint with the Auto palette", () => {
    const s = defaultChartStyle();
    expect(s.styleName).toBe("newsprint");
    expect(s.paletteName).toBe("auto");
    expect(s.bg).toBe("#faf5ea");
    expect(s.palette[0]).toBe("#c8492e");
  });

  it("applyStyle bakes the treatment + records the key", () => {
    const s = applyStyle(defaultChartStyle(), "blueprint");
    expect(s.styleName).toBe("blueprint");
    expect(s.bg).toBe("#122238");
    expect(s.shape).toBe("thin");
    expect(s.gridStyle).toBe("dots");
    expect(s.lineDash).toBe("6 4");
  });

  it("a named palette overrides the style's colors", () => {
    const s = applyPalette(applyStyle(defaultChartStyle(), "newsprint"), "ocean");
    expect(s.paletteName).toBe("ocean");
    expect(s.palette).toEqual(PALETTES.ocean.colors);
  });

  it("switching styles preserves a chosen named palette", () => {
    let s = applyStyle(defaultChartStyle(), "newsprint");
    s = applyPalette(s, "forest");
    s = applyStyle(s, "duotone"); // switch style while Forest is active
    expect(s.styleName).toBe("duotone");
    expect(s.paletteName).toBe("forest");
    expect(s.palette).toEqual(PALETTES.forest.colors);
    // ...but the treatment is Duotone's (dark canvas).
    expect(s.bg).toBe("#1f1c17");
  });

  it("auto restores the current style's signature palette", () => {
    let s = applyStyle(defaultChartStyle(), "sunset");
    s = applyPalette(s, "berry");
    expect(s.palette).toEqual(PALETTES.berry.colors);
    s = applyPalette(s, "auto");
    expect(s.paletteName).toBe("auto");
    expect(s.palette).toEqual(STYLES.sunset.treatment.palette);
  });
});
