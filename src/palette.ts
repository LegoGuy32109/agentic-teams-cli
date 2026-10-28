/** The Teams text and highlight colors, read from the web client's own formatted message (2026-10-16). */

export type PaletteEntry = { name: string; hex: string };

/** In the same order as the web client's color picker, so one short alias works for both lists. */
export const TEXT_COLORS: PaletteEntry[] = [
  { name: "Red", hex: "#B6424C" },
  { name: "Tuscan Orange", hex: "#CD5937" },
  { name: "Sunglow", hex: "#FDC030" },
  { name: "Pear", hex: "#BDCB4C" },
  { name: "Eucalyptus", hex: "#2B9B62" },
  { name: "Faded Jade", hex: "#37797B" },
  { name: "Fun Blue", hex: "#1E53A3" },
  { name: "Tyrian Purple", hex: "#A5397A" },
];

export const HIGHLIGHTS: PaletteEntry[] = [
  { name: "Kobi Red", hex: "#DF9299" },
  { name: "Rose Bud", hex: "#F4A593" },
  { name: "Goldenrod", hex: "#FDD472" },
  { name: "Primrose", hex: "#E5F18F" },
  { name: "Mantis", hex: "#82CDA8" },
  { name: "Regent St. Blue", hex: "#9DD9DB" },
  { name: "Periwinkle Gray", hex: "#C7D4E8" },
  { name: "Pale Rose", hex: "#EBD3E1" },
];

/** Short names. Each applies to the text color and to the highlight at the same position. */
export const ALIASES = ["red", "orange", "yellow", "lime", "green", "teal", "blue", "purple"];

const kebab = (name: string) => name.toLowerCase().replace(/\./g, "").replace(/\s+/g, "-");

/** Finds a color by short alias or by its Teams name in kebab case, such as `faded-jade` or `goldenrod`. */
export function findColor(list: PaletteEntry[], name: string): string | undefined {
  const alias = ALIASES.indexOf(name);
  if (alias >= 0) return list[alias]!.hex;
  return list.find(entry => kebab(entry.name) === name)?.hex;
}

export const textColor = (name: string) => findColor(TEXT_COLORS, name);
export const highlightColor = (name: string) => findColor(HIGHLIGHTS, name);

/**
 * A badge takes its background from the text colors, not the highlights. The text colors are stronger, so green is
 * a solid green and not a pale highlight, and white or black text stays readable on them.
 */
export const badgeColor = textColor;

/** Relative luminance of a `#RRGGBB` color, from 0 (black) to 1 (white), as defined in WCAG. */
export function luminance(hex: string): number {
  const [r, g, b] = [1, 3, 5].map(at => {
    const channel = parseInt(hex.slice(at, at + 2), 16) / 255;
    return channel <= 0.03928 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4;
  }) as [number, number, number];
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * Black on a light background, white on every other. The limit is above the point where black and white
 * have equal contrast (0.18), so a mid tone such as green keeps white text, as in the Teams badge that was checked live.
 */
export const badgeTextColor = (hex: string) => (luminance(hex) > 0.4 ? "black" : "white");
