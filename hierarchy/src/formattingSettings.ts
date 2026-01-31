import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import Model = formattingSettings.Model;
import SimpleCard = formattingSettings.SimpleCard;
import ToggleSwitch = formattingSettings.ToggleSwitch;
import ColorPicker = formattingSettings.ColorPicker;
import NumUpDown = formattingSettings.NumUpDown;
import ItemDropdown = formattingSettings.ItemDropdown;
import TextInput = formattingSettings.TextInput;

const orientationOptions: powerbi.IEnumMember[] = [
  { value: "TD", displayName: "Top-Down" },
  { value: "LR", displayName: "Left-Right" }
];

const lineStyleOptions: powerbi.IEnumMember[] = [
  { value: "solid", displayName: "Solid" },
  { value: "dashed", displayName: "Dashed" }
];

const lineTipOptions: powerbi.IEnumMember[] = [
  { value: "arrow", displayName: "Arrow" },
  { value: "none", displayName: "None" },
  { value: "square", displayName: "Square" },
  { value: "diamond", displayName: "Diamond" },
  { value: "circle", displayName: "Circle" }
];

const shapeOptions: powerbi.IEnumMember[] = [
  { value: "rounded", displayName: "Rounded" },
  { value: "square", displayName: "Square" },
  { value: "pill", displayName: "Pill" }
];

const viewOptions: powerbi.IEnumMember[] = [
  { value: "tree", displayName: "Tree" },
  { value: "table", displayName: "Table" }
];

const textAlignOptions: powerbi.IEnumMember[] = [
  { value: "left", displayName: "Left" },
  { value: "center", displayName: "Center" },
  { value: "right", displayName: "Right" }
];

const fontStyleOptions: powerbi.IEnumMember[] = [
  { value: "normal", displayName: "Normal" },
  { value: "bold", displayName: "Bold" },
  { value: "italic", displayName: "Italic" },
  { value: "boldItalic", displayName: "Bold Italic" }
];

class LayoutCardSettings extends SimpleCard {
  name = "layout";
  displayName = "Layout";
  slices = [
    new ItemDropdown({
      name: "orientation",
      displayName: "Orientation",
      items: orientationOptions,
      value: orientationOptions[0]
    }),
    new NumUpDown({ name: "levelSpacing", displayName: "Level spacing", value: 70 }),
    new NumUpDown({ name: "siblingSpacing", displayName: "Sibling spacing", value: 18 }),
    new NumUpDown({ name: "cardWidth", displayName: "Card width", value: 120 }),
    new NumUpDown({ name: "cardHeight", displayName: "Card height", value: 40 })
  ];
}

class AppearanceCardSettings extends SimpleCard {
  name = "appearance";
  displayName = "Appearance";
  slices = [
    new ToggleSwitch({ name: "useBackground", displayName: "Use background", value: true }),
    new ColorPicker({
      name: "backgroundColor",
      displayName: "Background color",
      value: { value: "#ffffff" }
    })
  ];
}

class LinesCardSettings extends SimpleCard {
  name = "lines";
  displayName = "Lines";
  slices = [
    new ColorPicker({ name: "lineColor", displayName: "Line color", value: { value: "#f3b27a" } }),
    new ColorPicker({ name: "activeColor", displayName: "Active line color", value: { value: "#f08b2e" } }),
    new NumUpDown({ name: "lineWidth", displayName: "Line width", value: 1 }),
    new ItemDropdown({
      name: "lineStyle",
      displayName: "Line style",
      items: lineStyleOptions,
      value: lineStyleOptions[0]
    }),
    new ToggleSwitch({ name: "showArrows", displayName: "Show pointers", value: true }),
    new ItemDropdown({
      name: "tipStyle",
      displayName: "Line tip",
      items: lineTipOptions,
      value: lineTipOptions[0]
    }),
    new NumUpDown({ name: "tipSize", displayName: "Tip size", value: 6 })
  ];
}

class NodesCardSettings extends SimpleCard {
  name = "nodes";
  displayName = "Nodes";
  slices = [
    new ColorPicker({ name: "fillColor", displayName: "Fill color", value: { value: "#ffffff" } }),
    new ColorPicker({ name: "strokeColor", displayName: "Border color", value: { value: "#e5e7eb" } }),
    new NumUpDown({ name: "strokeWidth", displayName: "Border width", value: 1 }),
    new NumUpDown({ name: "cornerRadius", displayName: "Corner radius", value: 6 }),
    new ItemDropdown({
      name: "shape",
      displayName: "Shape",
      items: shapeOptions,
      value: shapeOptions[0]
    }),
    new ToggleSwitch({ name: "showShadow", displayName: "Shadow", value: true }),
    new ColorPicker({ name: "titleColor", displayName: "Title color", value: { value: "#111827" } }),
    new ColorPicker({ name: "valueColor", displayName: "Value color", value: { value: "#6b7280" } }),
    new ItemDropdown({
      name: "titleAlign",
      displayName: "Title align",
      items: textAlignOptions,
      value: textAlignOptions[1]
    }),
    new ToggleSwitch({ name: "titleWrap", displayName: "Title wrap", value: true }),
    new NumUpDown({ name: "titleFontSize", displayName: "Title font size", value: 11 }),
    new TextInput({
      name: "titleFontFamily",
      displayName: "Title font family",
      value: "Segoe UI",
      placeholder: "Segoe UI"
    }),
    new ItemDropdown({
      name: "titleFontStyle",
      displayName: "Title font style",
      items: fontStyleOptions,
      value: fontStyleOptions[1]
    })
  ];
}

class LevelsCardSettings extends SimpleCard {
  name = "levels";
  displayName = "Level colors";
  slices = [
    new ToggleSwitch({ name: "enableLevelColors", displayName: "Enable", value: false }),
    new ColorPicker({ name: "level1Color", displayName: "Level 1", value: { value: "#ffffff" } }),
    new ColorPicker({ name: "level2Color", displayName: "Level 2", value: { value: "#f8fafc" } }),
    new ColorPicker({ name: "level3Color", displayName: "Level 3", value: { value: "#f1f5f9" } }),
    new ColorPicker({ name: "level4Color", displayName: "Level 4", value: { value: "#e2e8f0" } }),
    new ColorPicker({ name: "level5Color", displayName: "Level 5", value: { value: "#e0f2fe" } }),
    new ColorPicker({ name: "level6Color", displayName: "Level 6", value: { value: "#ede9fe" } })
  ];
}

class ControlsCardSettings extends SimpleCard {
  name = "controls";
  displayName = "Controls";
  slices = [
    new ToggleSwitch({ name: "showControls", displayName: "Show toolbar", value: true }),
    new ToggleSwitch({ name: "showSearch", displayName: "Show search", value: true }),
    new ToggleSwitch({ name: "showHierarchyFilter", displayName: "Show hierarchy filter", value: true }),
    new ToggleSwitch({ name: "showParentFilter", displayName: "Show parent filter", value: true }),
    new ToggleSwitch({ name: "showDropdownFilter", displayName: "Show dropdown filter", value: true }),
    new ToggleSwitch({ name: "showZoom", displayName: "Show zoom", value: true }),
    new ToggleSwitch({ name: "showViewToggle", displayName: "Show view toggle", value: true }),
    new ToggleSwitch({ name: "showCollapseExpand", displayName: "Show collapse/expand", value: true }),
    new ItemDropdown({
      name: "defaultView",
      displayName: "Default view",
      items: viewOptions,
      value: viewOptions[0]
    }),
    new NumUpDown({ name: "doubleClickZoomPercent", displayName: "Double-click zoom (%)", value: 130 })
  ];
}

class TableCardSettings extends SimpleCard {
  name = "table";
  displayName = "Table view";
  slices = [
    new ToggleSwitch({ name: "showHeader", displayName: "Show header", value: true }),
    new NumUpDown({ name: "rowHeight", displayName: "Row height", value: 28 }),
    new ToggleSwitch({ name: "zebra", displayName: "Zebra rows", value: true })
  ];
}

export class VisualFormattingSettingsModel extends Model {
  layout = new LayoutCardSettings();
  appearance = new AppearanceCardSettings();
  lines = new LinesCardSettings();
  nodes = new NodesCardSettings();
  levels = new LevelsCardSettings();
  controls = new ControlsCardSettings();
  table = new TableCardSettings();

  cards = [
    this.layout,
    this.appearance,
    this.lines,
    this.nodes,
    this.levels,
    this.controls,
    this.table
  ];
}
