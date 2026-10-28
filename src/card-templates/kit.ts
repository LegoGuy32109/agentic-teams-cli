import type { BadgeStyle } from "./validate.ts";

/**
 * The parts that every template builds from. The rules came from checking cards live in Teams (docs/cards.md):
 * - Every TextBlock wraps.
 * - A badge sits in a `width: auto` column that centers it, so it does not stretch or sit on the text baseline.
 * - An expand or collapse control is a text link (a Container with a `selectAction`), never a button.
 */
export type El = Record<string, any>;

export const text = (value: string, extra: El = {}): El => ({ type: "TextBlock", text: value, wrap: true, ...extra });

export const badge = (value: string, style: BadgeStyle, appearance: "Tint" | "Filled" = "Tint", icon?: string): El => ({
  type: "Badge", text: value, style, appearance, ...(icon ? { icon } : {}),
});

export const column = (items: El[], width: "auto" | "stretch" = "auto", extra: El = {}): El => ({
  type: "Column", width, verticalContentAlignment: "Center", ...extra, items,
});

export type HeaderOptions = {
  title: string;
  subtitle?: string;
  /** The corner badge. */
  badge?: { text: string; style: BadgeStyle };
  /** A Fluent icon name, or a PNG data URI for an image. */
  icon?: string;
  image?: string;
};

export function header(options: HeaderOptions): El {
  const columns: El[] = [];
  if (options.image) columns.push(column([{ type: "Image", url: options.image, width: "32px", height: "32px" }]));
  else if (options.icon) columns.push(column([{ type: "Icon", name: options.icon, size: "Medium", color: "Accent" }]));
  columns.push(column([
    text(options.title, { size: "Large", weight: "Bolder" }),
    ...(options.subtitle ? [text(options.subtitle, { isSubtle: true, spacing: "None", size: "Small" })] : []),
  ], "stretch"));
  if (options.badge) columns.push(column([badge(options.badge.text, options.badge.style, "Filled")]));
  return { type: "ColumnSet", columns };
}

/** A bordered box. `style` is the Container style: default, emphasis, good, warning, attention, accent. */
export function box(items: El[], title?: string, style = "emphasis"): El {
  return {
    type: "Container", style, showBorder: true, roundedCorners: true, spacing: "Medium",
    items: [...(title ? [text(title, { weight: "Bolder" })] : []), ...items],
  };
}

export const bullets = (lines: string[], extra: El = {}): El[] => lines.map(line => text(`• ${line}`, { spacing: "Small", ...extra }));

/** Lines in a colored box. One line stands alone, as plain text. Several lines get bullets. */
export const noted = (lines: string[], color: string): El[] =>
  lines.length === 1 ? [text(lines[0]!, { color })] : bullets(lines, { color });

export type RowOptions = {
  title: string;
  subline?: string;
  right?: El;
  /** A Fluent icon at the left: name and color. */
  icon?: { name: string; color: string };
  separator?: boolean;
};

/** The columns of one line: a bold title, an optional subtle line below it, and an optional badge at the right. */
export function rowColumns(options: RowOptions): El {
  const columns: El[] = [];
  if (options.icon) columns.push(column([{ type: "Icon", name: options.icon.name, size: "Standard", color: options.icon.color }]));
  columns.push(column([
    text(options.title, { weight: "Bolder" }),
    ...(options.subline ? [text(options.subline, { isSubtle: true, size: "Small", spacing: "None" })] : []),
  ], "stretch"));
  if (options.right) columns.push(column([options.right]));
  return { type: "ColumnSet", columns };
}

/** One line of a list, in a Container that has a separator line above it. */
export function row(options: RowOptions): El {
  return { type: "Container", separator: options.separator ?? true, spacing: "Small", items: [rowColumns(options)] };
}

/** Rows in a list: only the first has no separator line above it. */
export const rows = (items: Array<Omit<RowOptions, "separator">>): El[] => items.map((item, index) => row({ ...item, separator: index > 0 }));

/** A text link that shows or hides elements. It is a Container with a `selectAction`, not a button. `id` makes it a target too. */
export function toggleLink(label: string, targets: string[], extra: El = {}): El {
  return {
    type: "Container", spacing: "Small",
    selectAction: { type: "Action.ToggleVisibility", targetElements: targets },
    items: [text(label, { color: "Accent", size: "Small", weight: "Bolder" })],
    ...extra,
  };
}

export const openUrl = (title: string, url: string, extra: El = {}): El => ({ type: "Action.OpenUrl", title, url, ...extra });

export function adaptiveCard(body: El[], extra: El = {}): El {
  return { $schema: "http://adaptivecards.io/schemas/adaptive-card.json", type: "AdaptiveCard", version: "1.5", body, ...extra };
}

/** A ProgressBar takes only these colors. */
export function progressColor(style: BadgeStyle): "Accent" | "Good" | "Warning" | "Attention" {
  return style === "Good" || style === "Warning" || style === "Attention" ? style : "Accent";
}

export const plural = (count: number, one: string, many = `${one}s`) => `${count} ${count === 1 ? one : many}`;
