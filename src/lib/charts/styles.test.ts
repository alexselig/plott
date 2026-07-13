import { describe, expect, it } from "vitest";

import {
  applyPalette,
  applyTreatment,
  cardBg,
  darken,
  defaultChartStyle,
  lighten,
  PALETTE_ORDER,
  PALETTES,
  TREATMENT_ORDER,
  TREATMENTS,
  treatmentOf,
} from "@/lib/charts/styles";

describe("Chart Treatment System", () => {
  it("has 4 palettes (5 colors each) and 13 treatments", () => {
    expect(PALETTE_ORDER.length).toBe(4);
    expect(TREATMENT_ORDER.length).toBe(13);
    PALETTE_ORDER.forEach((k) => expect(PALETTES[k].colors.length).toBe(5));
    TREATMENT_ORDER.forEach((k) => expect(TREATMENTS[k].chrome).toBeTruthy());
  });

  it("ships the Signal palette with the exact spec hexes", () => {
    expect(PALETTES.signal.colors).toEqual(["#6E56CF", "#3B82C4", "#2F9E6E", "#C97D2E", "#C4456B"]);
  });

  it("lighten/darken blend toward 255/0", () => {
    expect(lighten("#000000", 0.5)).toBe("#808080");
    expect(darken("#ffffff", 0.5)).toBe("#808080");
    expect(lighten("#6E56CF", 0)).toBe("#6e56cf");
  });

  it("default style is Studio Flat with the Signal palette", () => {
    const s = defaultChartStyle();
    expect(s.treatment).toBe("studioFlat");
    expect(s.paletteName).toBe("signal");
    expect(treatmentOf(s)).toBe("studioFlat");
    // Background behind the chart is constant across treatments (matches the app).
    expect(cardBg()).toBe("#fffdf8");
  });

  it("applyTreatment records the key; applyPalette swaps colors", () => {
    let s = defaultChartStyle();
    s = applyTreatment(s, "gradientGlow");
    expect(s.treatment).toBe("gradientGlow");
    // The treatment still flips its dark chrome flag…
    expect(TREATMENTS[treatmentOf(s)].chrome.dark).toBe(true);
    // …but the rendered background stays constant regardless of treatment.
    expect(cardBg()).toBe("#fffdf8");
    s = applyPalette(s, "forest");
    expect(s.paletteName).toBe("forest");
    expect(s.palette).toEqual(PALETTES.forest.colors);
  });
});
