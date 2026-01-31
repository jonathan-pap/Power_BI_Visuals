import powerbi from "powerbi-visuals-api";

export type Orientation = "TD" | "LR";
export type LineStyle = "solid" | "dashed";
export type LineTipStyle = "arrow" | "none" | "square" | "diamond" | "circle";
export type TextAlign = "left" | "center" | "right";
export type FontStyle = "normal" | "bold" | "italic" | "boldItalic";
export type NodeShape = "rounded" | "square" | "pill";
export type ViewMode = "tree" | "table";

export interface LayoutSettings {
  orientation: Orientation;
  levelSpacing: number;
  siblingSpacing: number;
  cardWidth: number;
  cardHeight: number;
}

export interface AppearanceSettings {
  useBackground: boolean;
  backgroundColor: string;
}

export interface LineSettings {
  lineColor: string;
  activeColor: string;
  lineWidth: number;
  lineStyle: LineStyle;
  showArrows: boolean;
  tipStyle: LineTipStyle;
  tipSize: number;
}

export interface NodeSettings {
  fillColor: string;
  strokeColor: string;
  strokeWidth: number;
  cornerRadius: number;
  shape: NodeShape;
  showShadow: boolean;
  titleColor: string;
  valueColor: string;
  titleAlign: TextAlign;
  titleWrap: boolean;
  titleFontSize: number;
  titleFontFamily: string;
  titleFontStyle: FontStyle;
}

export interface LevelSettings {
  enable: boolean;
  levelColors: string[];
}

export interface ControlSettings {
  showControls: boolean;
  showSearch: boolean;
  showHierarchyFilter: boolean;
  showParentFilter: boolean;
  showDropdownFilter: boolean;
  showZoom: boolean;
  showViewToggle: boolean;
  showCollapseExpand: boolean;
  defaultView: ViewMode;
  doubleClickZoomPercent: number;
}

export interface TableSettings {
  showHeader: boolean;
  rowHeight: number;
  zebra: boolean;
}

export interface VisualSettings {
  layout: LayoutSettings;
  appearance: AppearanceSettings;
  lines: LineSettings;
  nodes: NodeSettings;
  levels: LevelSettings;
  controls: ControlSettings;
  table: TableSettings;
}

export const DefaultLayoutSettings: LayoutSettings = {
  orientation: "TD",
  levelSpacing: 70,
  siblingSpacing: 18,
  cardWidth: 120,
  cardHeight: 40
};

export const DefaultVisualSettings: VisualSettings = {
  layout: DefaultLayoutSettings,
  appearance: {
    useBackground: true,
    backgroundColor: "#ffffff"
  },
  lines: {
    lineColor: "#f3b27a",
    activeColor: "#f08b2e",
    lineWidth: 1,
    lineStyle: "solid",
    showArrows: true,
    tipStyle: "arrow",
    tipSize: 6
  },
  nodes: {
    fillColor: "#ffffff",
    strokeColor: "#e5e7eb",
    strokeWidth: 1,
    cornerRadius: 6,
    shape: "rounded",
    showShadow: true,
    titleColor: "#111827",
    valueColor: "#6b7280",
    titleAlign: "center",
    titleWrap: true,
    titleFontSize: 11,
    titleFontFamily: "Segoe UI",
    titleFontStyle: "bold"
  },
  levels: {
    enable: false,
    levelColors: [
      "#ffffff",
      "#f8fafc",
      "#f1f5f9",
      "#e2e8f0",
      "#e0f2fe",
      "#ede9fe"
    ]
  },
  controls: {
    showControls: true,
    showSearch: true,
    showHierarchyFilter: true,
    showParentFilter: true,
    showDropdownFilter: true,
    showZoom: true,
    showViewToggle: true,
    showCollapseExpand: true,
    defaultView: "tree",
    doubleClickZoomPercent: 130
  },
  table: {
    showHeader: true,
    rowHeight: 28,
    zebra: true
  }
};

export function getVisualSettings(dataView?: powerbi.DataView): VisualSettings {
  const objects = (dataView?.metadata?.objects ?? {}) as any;

  const layout = objects.layout ?? {};
  const appearance = objects.appearance ?? {};
  const lines = objects.lines ?? {};
  const nodes = objects.nodes ?? {};
  const levels = objects.levels ?? {};
  const controls = objects.controls ?? {};
  const table = objects.table ?? {};

  const levelColors = [
    toColor(levels.level1Color, DefaultVisualSettings.levels.levelColors[0]),
    toColor(levels.level2Color, DefaultVisualSettings.levels.levelColors[1]),
    toColor(levels.level3Color, DefaultVisualSettings.levels.levelColors[2]),
    toColor(levels.level4Color, DefaultVisualSettings.levels.levelColors[3]),
    toColor(levels.level5Color, DefaultVisualSettings.levels.levelColors[4]),
    toColor(levels.level6Color, DefaultVisualSettings.levels.levelColors[5])
  ];

  return {
    layout: {
      orientation: (layout.orientation as Orientation) ?? DefaultLayoutSettings.orientation,
      levelSpacing: toNumber(layout.levelSpacing, DefaultLayoutSettings.levelSpacing),
      siblingSpacing: toNumber(layout.siblingSpacing, DefaultLayoutSettings.siblingSpacing),
      cardWidth: toNumber(layout.cardWidth, DefaultLayoutSettings.cardWidth),
      cardHeight: toNumber(layout.cardHeight, DefaultLayoutSettings.cardHeight)
    },
    appearance: {
      useBackground: toBoolean(appearance.useBackground, DefaultVisualSettings.appearance.useBackground),
      backgroundColor: toColor(appearance.backgroundColor, DefaultVisualSettings.appearance.backgroundColor)
    },
    lines: {
      lineColor: toColor(lines.lineColor, DefaultVisualSettings.lines.lineColor),
      activeColor: toColor(lines.activeColor, DefaultVisualSettings.lines.activeColor),
      lineWidth: toNumber(lines.lineWidth, DefaultVisualSettings.lines.lineWidth),
      lineStyle: toEnum(lines.lineStyle, ["solid", "dashed"], DefaultVisualSettings.lines.lineStyle),
      showArrows: toBoolean(lines.showArrows, DefaultVisualSettings.lines.showArrows),
      tipStyle: toEnum(
        lines.tipStyle,
        ["arrow", "none", "square", "diamond", "circle"],
        DefaultVisualSettings.lines.tipStyle
      ),
      tipSize: toNumber(lines.tipSize, DefaultVisualSettings.lines.tipSize)
    },
    nodes: {
      fillColor: toColor(nodes.fillColor, DefaultVisualSettings.nodes.fillColor),
      strokeColor: toColor(nodes.strokeColor, DefaultVisualSettings.nodes.strokeColor),
      strokeWidth: toNumber(nodes.strokeWidth, DefaultVisualSettings.nodes.strokeWidth),
      cornerRadius: toNumber(nodes.cornerRadius, DefaultVisualSettings.nodes.cornerRadius),
      shape: toEnum(nodes.shape, ["rounded", "square", "pill"], DefaultVisualSettings.nodes.shape),
      showShadow: toBoolean(nodes.showShadow, DefaultVisualSettings.nodes.showShadow),
      titleColor: toColor(nodes.titleColor, DefaultVisualSettings.nodes.titleColor),
      valueColor: toColor(nodes.valueColor, DefaultVisualSettings.nodes.valueColor),
      titleAlign: toEnum(nodes.titleAlign, ["left", "center", "right"], DefaultVisualSettings.nodes.titleAlign),
      titleWrap: toBoolean(nodes.titleWrap, DefaultVisualSettings.nodes.titleWrap),
      titleFontSize: toNumber(nodes.titleFontSize, DefaultVisualSettings.nodes.titleFontSize),
      titleFontFamily: toText(nodes.titleFontFamily, DefaultVisualSettings.nodes.titleFontFamily),
      titleFontStyle: toEnum(
        nodes.titleFontStyle,
        ["normal", "bold", "italic", "boldItalic"],
        DefaultVisualSettings.nodes.titleFontStyle
      )
    },
    levels: {
      enable: toBoolean(levels.enableLevelColors, DefaultVisualSettings.levels.enable),
      levelColors
    },
    controls: {
      showControls: toBoolean(controls.showControls, DefaultVisualSettings.controls.showControls),
      showSearch: toBoolean(controls.showSearch, DefaultVisualSettings.controls.showSearch),
      showHierarchyFilter: toBoolean(
        controls.showHierarchyFilter,
        DefaultVisualSettings.controls.showHierarchyFilter
      ),
      showParentFilter: toBoolean(controls.showParentFilter, DefaultVisualSettings.controls.showParentFilter),
      showDropdownFilter: toBoolean(
        controls.showDropdownFilter,
        DefaultVisualSettings.controls.showDropdownFilter
      ),
      showZoom: toBoolean(controls.showZoom, DefaultVisualSettings.controls.showZoom),
      showViewToggle: toBoolean(controls.showViewToggle, DefaultVisualSettings.controls.showViewToggle),
      showCollapseExpand: toBoolean(
        controls.showCollapseExpand,
        DefaultVisualSettings.controls.showCollapseExpand
      ),
      defaultView: toEnum(controls.defaultView, ["tree", "table"], DefaultVisualSettings.controls.defaultView),
      doubleClickZoomPercent: toNumber(
        controls.doubleClickZoomPercent,
        DefaultVisualSettings.controls.doubleClickZoomPercent
      )
    },
    table: {
      showHeader: toBoolean(table.showHeader, DefaultVisualSettings.table.showHeader),
      rowHeight: toNumber(table.rowHeight, DefaultVisualSettings.table.rowHeight),
      zebra: toBoolean(table.zebra, DefaultVisualSettings.table.zebra)
    }
  };
}

function toNumber(v: any, fallback: number): number {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function toBoolean(v: any, fallback: boolean): boolean {
  if (typeof v === "boolean") return v;
  return fallback;
}

function toEnum<T extends string>(v: any, allowed: T[], fallback: T): T {
  if (typeof v === "string" && allowed.includes(v as T)) return v as T;
  return fallback;
}

function toColor(v: any, fallback: string): string {
  if (!v) return fallback;
  if (typeof v === "string") return v;
  const solid = v.solid;
  if (solid && typeof solid.color === "string") return solid.color;
  return fallback;
}

function toText(v: any, fallback: string): string {
  if (typeof v === "string" && v.trim().length > 0) return v;
  return fallback;
}
