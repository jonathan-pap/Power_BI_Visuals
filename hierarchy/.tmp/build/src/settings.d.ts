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
export declare const DefaultLayoutSettings: LayoutSettings;
export declare const DefaultVisualSettings: VisualSettings;
export declare function getVisualSettings(dataView?: powerbi.DataView): VisualSettings;
