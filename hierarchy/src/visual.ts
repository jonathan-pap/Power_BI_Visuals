"use strict";

import powerbi from "powerbi-visuals-api";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

import DataView = powerbi.DataView;
import DataViewCategorical = powerbi.DataViewCategorical;

import ISelectionId = powerbi.visuals.ISelectionId;
import ISelectionManager = powerbi.extensibility.ISelectionManager;

import { tree, stratify, HierarchyNode } from "d3-hierarchy";
import { getVisualSettings, VisualSettings, ViewMode } from "./settings";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./formattingSettings";

/** Base styling (uses host palette for accent where possible) */
const UI = {
  bg: "#ffffff",
  link: "#f3b27a",
  linkActive: "#f08b2e",

  cardFill: "#ffffff",
  cardStroke: "#e5e7eb",
  cardShadow: "rgba(15, 23, 42, 0.08)",

  title: "#111827",
  subtext: "#6b7280",

  accent: "#f08b2e",
  accentSoft: "#fff4e6",

  toggleFill: "#ffffff",
  toggleStroke: "#d1d5db",
  toggleText: "#f08b2e"
};

const FONT = {
  title: "600 11px Segoe UI",
  value: "500 10px Segoe UI",
  toggle: "700 10px Segoe UI"
};

type NodeRow = {
  id: string;
  parentId: string | null;
  label: string;
  value?: number | string | null;
  sparkline?: number | string | null;
  tooltip?: number | string | null;
  dropdown?: string | null;
  selectionId: ISelectionId;
};

type LayoutNode = {
  id: string;
  label: string;
  value?: number | string | null;
  sparkline?: number | string | null;
  tooltip?: number | string | null;
  x: number;
  y: number;
  selectionId: ISelectionId;
  depth: number;
  parent?: LayoutNode;
  children?: LayoutNode[];
};

type Hit = {
  node: LayoutNode;
  worldX: number;
  worldY: number;
  localX: number;
  localY: number;
};

export class Visual implements IVisual {
  private host: powerbi.extensibility.visual.IVisualHost;
  private selectionManager: ISelectionManager;
  private tooltipService?: powerbi.extensibility.ITooltipService;
  private localizationManager: powerbi.extensibility.ILocalizationManager;
  private formattingSettingsService: FormattingSettingsService;
  private formattingSettingsModel: VisualFormattingSettingsModel;

  private root: HTMLElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;

  private settings: VisualSettings;

  private viewMode: ViewMode = "tree";
  private userSetView = false;

  private toolbar: HTMLDivElement;
  private viewGroup: HTMLDivElement;
  private zoomGroup: HTMLDivElement;
  private collapseGroup: HTMLDivElement;
  private searchGroup: HTMLDivElement;
  private searchInput: HTMLInputElement;
  private filterGroup: HTMLDivElement;
  private hierarchyFilter: HTMLSelectElement;
  private parentFilter: HTMLSelectElement;
  private dropdownFilter: HTMLSelectElement;
  private zoomLabel: HTMLInputElement;
  private treeButton: HTMLButtonElement;
  private tableButton: HTMLButtonElement;
  private collapseAllButton: HTMLButtonElement;
  private expandAllButton: HTMLButtonElement;

  private tableContainer: HTMLDivElement;
  private tableEl: HTMLTableElement;
  private tableHead: HTMLTableSectionElement;
  private tableBody: HTMLTableSectionElement;
  private tableRows: LayoutNode[] = [];

  private landingPage: HTMLDivElement;
  private focusedNodeId: string | null = null;
  private focusedIndex = 0;

  // view transform
  private tx = 20;
  private ty = 20;
  private scale = 1;

  // layout cache
  private nodes: LayoutNode[] = [];
  private links: Array<{ source: LayoutNode; target: LayoutNode }> = [];

  // hit testing rects (world coords)
  private nodeRects: Array<{ node: LayoutNode; x: number; y: number; w: number; h: number }> = [];
  private toggleRects: Array<{ nodeId: string; x: number; y: number; w: number; h: number }> = [];

  // last known viewport (CSS pixels)
  private lastViewportW = 0;
  private lastViewportH = 0;

  // interaction state
  private hoveredId: string | null = null;
  private selectedIds = new Set<string>();
  private searchQuery = "";
  private hierarchyFilterValue: string | null = null;
  private parentFilterValue: string | null = null;
  private dropdownFilterValue: string | null = null;

  // data + collapse state
  private allRows: NodeRow[] = [];
  private fullRows: NodeRow[] = [];
  private childrenMap = new Map<string, string[]>();
  private fullChildrenMap = new Map<string, string[]>();
  private collapsedIds = new Set<string>();
  private valueDisplayName = "Value";
  private sparklineDisplayName = "Sparkline";
  private tooltipDisplayName = "Tooltip";
  private labelDisplayName = "Name";
  private dropdownDisplayName = "Dropdown";
  private hasDropdownField = false;
  private sparklineMin: number | null = null;
  private sparklineMax: number | null = null;

  // accessibility / host behaviour
  private allowInteractions = true;

  constructor(options: VisualConstructorOptions) {
    this.host = options.host;
    this.selectionManager = this.host.createSelectionManager();
    this.tooltipService = this.host.tooltipService;
    this.localizationManager = this.host.createLocalizationManager();
    this.formattingSettingsService = new FormattingSettingsService();
    this.formattingSettingsModel = new VisualFormattingSettingsModel();

    this.root = options.element;
    this.root.style.position = "relative";

    this.canvas = document.createElement("canvas");
    this.canvas.style.width = "100%";
    this.canvas.style.height = "100%";
    this.canvas.style.display = "block";
    (this.canvas.style as any).touchAction = "none";
    this.canvas.tabIndex = 0;
    this.canvas.setAttribute("role", "application");
    this.canvas.setAttribute("aria-label", this.localize("Visual.Name", "Hierarchy Flow"));
    this.root.appendChild(this.canvas);

    const ctx = this.canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas not supported.");
    this.ctx = ctx;

    this.settings = getVisualSettings(undefined);

    this.createToolbar();
    this.createTable();
    this.createLandingPage();
    this.wireInteractions();
  }

  public update(options: VisualUpdateOptions): void {
    const eventService = this.host.eventService;
    eventService?.renderingStarted(options);

    const dv: DataView | undefined = options.dataViews?.[0];
    const viewport = options.viewport;

    this.allowInteractions = (this.host as any).allowInteractions !== false;

    this.lastViewportW = viewport.width;
    this.lastViewportH = viewport.height;

    this.resizeCanvas(viewport.width, viewport.height);

    // formatting model + settings
    if (dv) {
      this.formattingSettingsModel = this.formattingSettingsService.populateFormattingSettingsModel(
        VisualFormattingSettingsModel,
        dv
      );
    } else {
      this.formattingSettingsModel = new VisualFormattingSettingsModel();
    }

    this.settings = getVisualSettings(dv);
    this.updateLandingPageStyle();
    if (!this.userSetView || !this.settings.controls.showViewToggle) {
      this.viewMode = this.settings.controls.defaultView;
    }
    this.applyToolbarSettings();

    try {
      // parse data
      const model = this.parseDataView(dv);
      if (!model || model.length === 0) {
        this.nodes = [];
        this.links = [];
        this.tableRows = [];
        this.allRows = [];
        this.childrenMap.clear();
        this.sparklineMin = null;
        this.sparklineMax = null;
        this.hasDropdownField = false;
        this.dropdownDisplayName = this.localize("Toolbar.FilterDropdown", "Dropdown filter");
        this.updateFilterOptions();
        this.applyToolbarSettings();
        this.clearMessage();
        this.showLandingPage(true);
        eventService?.renderingFinished(options);
        return;
      }

      this.clearMessage();
      this.showLandingPage(false);

      const isFiltered = dv?.metadata?.isDataFilterApplied === true;
      const hasCache = this.fullRows.length > 0;
      const reduced = hasCache && model.length < this.fullRows.length;
      const canUseCache = isFiltered && hasCache && reduced && this.hasAllIds(this.fullRows, model);

      if (!canUseCache) {
        this.fullRows = model;
        this.fullChildrenMap = this.buildChildrenMap(this.fullRows);
      }

      // store dataset for collapsing logic (expanded if filtered)
      this.allRows = canUseCache ? this.expandRowsForFilter(model, this.fullRows) : model;
      this.childrenMap = this.fullChildrenMap.size ? this.fullChildrenMap : this.buildChildrenMap(this.allRows);
      this.setSparklineRange(this.allRows);
      this.updateFilterOptions();
      this.applyToolbarSettings();

      // compute layout based on collapsed state
      const ok = this.computeLayoutFromState();
      if (!ok) {
        eventService?.renderingFailed(options, this.localize(
          "Message.InvalidHierarchy",
          "Invalid hierarchy: duplicates, cycles, or missing parents."
        ));
        return;
      }

      this.ensureFocus();

      if (dv?.metadata?.segment) {
        this.host.fetchMoreData?.();
      }

      // render
      this.renderView();
      eventService?.renderingFinished(options);
    } catch (err) {
      const reason = err instanceof Error ? err.message : undefined;
      eventService?.renderingFailed(options, reason);
    }
  }

  public getFormattingModel(): powerbi.visuals.FormattingModel {
    return this.formattingSettingsService.buildFormattingModel(this.formattingSettingsModel);
  }

  // ---------------------------
  // Canvas helpers
  // ---------------------------
  private resizeCanvas(cssW: number, cssH: number): void {
    const dpr = window.devicePixelRatio || 1;
    this.canvas.width = Math.max(1, Math.floor(cssW * dpr));
    this.canvas.height = Math.max(1, Math.floor(cssH * dpr));
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  // ---------------------------
  // Data parsing
  // ---------------------------
  private parseDataView(dv?: DataView): NodeRow[] | null {
    const categorical = dv?.categorical as DataViewCategorical | undefined;
    const cats = categorical?.categories;
    if (!cats || cats.length < 2) return null;

    const nodeIdCat = cats.find(c => c.source.roles?.["hierarchy"]);
    const parentIdCat = cats.find(c => c.source.roles?.["parent"]);
    const fieldCats = cats.filter(c => c.source.roles?.["fields"]);
    const dropdownCat = cats.find(c => c.source.roles?.["dropdown"]);
    if (!nodeIdCat || !parentIdCat) return null;

    this.labelDisplayName = fieldCats.length > 0
      ? this.localize("Table.Header.Fields", "Fields")
      : (nodeIdCat.source?.displayName ?? this.localize("Table.Header.Fields", "Fields"));
    this.dropdownDisplayName = dropdownCat?.source?.displayName ?? this.localize("Toolbar.FilterDropdown", "Dropdown filter");
    this.hasDropdownField = Boolean(dropdownCat);

    const values = categorical?.values;
    const sparkCol = values?.find(v => v.source.roles?.["sparkline"]);
    const valuesCol = values?.find(v => v.source.roles?.["values"]);
    const tooltipCol = values?.find(v => v.source.roles?.["tooltip"]);

    this.sparklineDisplayName = sparkCol?.source?.displayName ?? "Sparkline";
    this.valueDisplayName = valuesCol?.source?.displayName ?? "Value";
    this.tooltipDisplayName = tooltipCol?.source?.displayName ?? "Tooltip";

    const rows: NodeRow[] = [];
    const len = nodeIdCat.values.length;

    for (let i = 0; i < len; i++) {
      const id = String(nodeIdCat.values[i] ?? "").trim();
      if (!id) continue;

      const pidRaw = parentIdCat.values[i];
      const parentId =
        pidRaw === null || pidRaw === undefined || String(pidRaw).trim() === ""
          ? null
          : String(pidRaw).trim();

      const label = fieldCats.length > 0
        ? fieldCats
          .map(c => c.values[i])
          .map(v => (v === null || v === undefined ? "" : String(v).trim()))
          .filter(Boolean)
          .join(" / ") || id
        : id;

      const sparkline = sparkCol ? this.getValueWithHighlight(sparkCol, i) : null;
      const value = valuesCol ? this.getValueWithHighlight(valuesCol, i) : null;
      const tooltip = tooltipCol ? (tooltipCol.values[i] as number | string | null) : null;
      const dropdownRaw = dropdownCat ? dropdownCat.values[i] : null;
      const dropdown = dropdownRaw === null || dropdownRaw === undefined ? null : String(dropdownRaw);

      const selectionId = this.host
        .createSelectionIdBuilder()
        .withCategory(nodeIdCat, i)
        .createSelectionId();

      rows.push({ id, parentId, label, value, sparkline, tooltip, dropdown, selectionId });
    }

    return rows;
  }

  private getValueWithHighlight(column: powerbi.DataViewValueColumn, index: number): number | string | null {
    const highlights = (column as any).highlights as (number | string | null)[] | undefined;
    if (highlights && highlights.length > index) {
      const highlighted = highlights[index];
      if (highlighted !== null && highlighted !== undefined) return highlighted;
    }
    return column.values[index] as number | string | null;
  }

  private buildChildrenMap(rows: NodeRow[]): Map<string, string[]> {
    const map = new Map<string, string[]>();
    for (const r of rows) {
      if (!r.parentId) continue;
      const arr = map.get(r.parentId) ?? [];
      arr.push(r.id);
      map.set(r.parentId, arr);
    }
    return map;
  }

  private hasAllIds(baseRows: NodeRow[], subset: NodeRow[]): boolean {
    const ids = new Set(baseRows.map(r => r.id));
    for (const r of subset) {
      if (!ids.has(r.id)) return false;
    }
    return true;
  }

  private expandRowsForFilter(filteredRows: NodeRow[], fullRows: NodeRow[]): NodeRow[] {
    if (filteredRows.length === 0) return [];

    const fullById = new Map(fullRows.map(r => [r.id, r]));
    const filteredById = new Map(filteredRows.map(r => [r.id, r]));

    const rootIds = new Set<string>();
    for (const r of filteredRows) {
      let currentId = r.id;
      while (true) {
        const parentId = fullById.get(currentId)?.parentId;
        if (!parentId || !fullById.has(parentId)) break;
        currentId = parentId;
      }
      rootIds.add(currentId);
    }

    const include = new Set<string>();
    const stack = Array.from(rootIds);
    while (stack.length) {
      const id = stack.pop()!;
      if (include.has(id)) continue;
      include.add(id);
      const children = this.fullChildrenMap.get(id) ?? [];
      for (const c of children) stack.push(c);
    }

    return fullRows
      .filter(r => include.has(r.id))
      .map(r => {
        const filtered = filteredById.get(r.id);
        if (!filtered) {
          return {
            ...r,
            value: null,
            sparkline: null,
            tooltip: null
          };
        }
        return {
          ...r,
          value: filtered.value,
          sparkline: filtered.sparkline,
          tooltip: filtered.tooltip,
          selectionId: filtered.selectionId
        };
      });
  }

  private applySearchFilter(rows: NodeRow[], query: string): NodeRow[] {
    const q = query.trim().toLowerCase();
    if (!q) return rows;

    const fullById = new Map(rows.map(r => [r.id, r]));
    const matchedRoots = new Set<string>();

    for (const r of rows) {
      const text = (r.label ?? "").toLowerCase();
      if (!text.includes(q)) continue;

      let currentId = r.id;
      while (true) {
        const parentId = fullById.get(currentId)?.parentId;
        if (!parentId || !fullById.has(parentId)) break;
        currentId = parentId;
      }
      matchedRoots.add(currentId);
    }

    if (matchedRoots.size === 0) return [];

    const include = new Set<string>();
    const stack = Array.from(matchedRoots);
    while (stack.length) {
      const id = stack.pop()!;
      if (include.has(id)) continue;
      include.add(id);
      const children = this.childrenMap.get(id) ?? [];
      for (const c of children) stack.push(c);
    }

    return rows.filter(r => include.has(r.id));
  }

  private applyDropdownFilters(rows: NodeRow[]): NodeRow[] {
    let result = rows;
    if (this.hierarchyFilterValue) {
      result = this.filterToNodeBranch(result, this.hierarchyFilterValue);
    }
    if (this.parentFilterValue) {
      result = this.filterToParentBranch(result, this.parentFilterValue);
    }
    if (this.dropdownFilterValue) {
      result = this.filterToDropdownBranch(result, this.dropdownFilterValue);
    }
    return result;
  }

  private filterToNodeBranch(rows: NodeRow[], nodeId: string): NodeRow[] {
    const byId = new Map(rows.map(r => [r.id, r]));
    if (!byId.has(nodeId)) return rows;

    const include = new Set<string>();

    // ancestors
    let currentId: string | null = nodeId;
    while (currentId) {
      include.add(currentId);
      const parentId = byId.get(currentId)?.parentId ?? null;
      if (!parentId || !byId.has(parentId)) break;
      currentId = parentId;
    }

    // descendants
    const stack = [nodeId];
    while (stack.length) {
      const id = stack.pop()!;
      const children = this.childrenMap.get(id) ?? [];
      for (const c of children) {
        if (!include.has(c)) {
          include.add(c);
          stack.push(c);
        }
      }
    }

    return rows.filter(r => include.has(r.id));
  }

  private filterToParentBranch(rows: NodeRow[], parentId: string): NodeRow[] {
    const byId = new Map(rows.map(r => [r.id, r]));
    const children = rows.filter(r => r.parentId === parentId).map(r => r.id);
    if (children.length === 0) return rows;

    const include = new Set<string>();

    // include parent + ancestors
    let currentId: string | null = parentId;
    while (currentId) {
      if (byId.has(currentId)) include.add(currentId);
      const nextParent = byId.get(currentId)?.parentId ?? null;
      if (!nextParent || !byId.has(nextParent)) break;
      currentId = nextParent;
    }

    // include descendants of all children
    const stack = [...children];
    for (const c of children) include.add(c);
    while (stack.length) {
      const id = stack.pop()!;
      const kids = this.childrenMap.get(id) ?? [];
      for (const k of kids) {
        if (!include.has(k)) {
          include.add(k);
          stack.push(k);
        }
      }
    }

    return rows.filter(r => include.has(r.id));
  }

  private filterToDropdownBranch(rows: NodeRow[], value: string): NodeRow[] {
    const byId = new Map(rows.map(r => [r.id, r]));
    const matchedRoots = new Set<string>();
    rows
      .filter(r => {
        const dropdownValue = (r.dropdown ?? "").trim();
        return dropdownValue === value || r.label === value || r.id === value;
      })
      .forEach(r => {
        const candidate = r.parentId && byId.has(r.parentId) ? r.parentId : r.id;
        matchedRoots.add(candidate);
      });

    if (matchedRoots.size === 0) return [];

    const include = new Set<string>();
    for (const id of matchedRoots) {
      if (!byId.has(id)) continue;

      // include node + ancestors
      let currentId: string | null = id;
      while (currentId) {
        include.add(currentId);
        const nextParent = byId.get(currentId)?.parentId ?? null;
        if (!nextParent || !byId.has(nextParent)) break;
        currentId = nextParent;
      }

      // include descendants
      const stack = [id];
      while (stack.length) {
        const nodeId = stack.pop()!;
        const kids = this.childrenMap.get(nodeId) ?? [];
        for (const k of kids) {
          if (!include.has(k)) {
            include.add(k);
            stack.push(k);
          }
        }
      }
    }

    return rows.filter(r => include.has(r.id));
  }


  private updateFilterOptions(): void {
    if (!this.hierarchyFilter || !this.parentFilter) return;

    const sourceRows = this.fullRows.length > 0 ? this.fullRows : this.allRows;
    const ids = new Set<string>();
    const parentIds = new Set<string>();
    const dropdownValues = new Set<string>();

    for (const r of sourceRows) {
      ids.add(r.id);
      if (r.parentId) parentIds.add(r.parentId);
      if (r.dropdown) dropdownValues.add(r.dropdown);
    }

    const buildOptions = (select: HTMLSelectElement, values: string[], placeholder: string) => {
      const current = select.value;
      while (select.firstChild) select.removeChild(select.firstChild);

      const allOption = document.createElement("option");
      allOption.value = "";
      allOption.textContent = placeholder;
      select.appendChild(allOption);

      for (const v of values) {
        const opt = document.createElement("option");
        opt.value = v;
        opt.textContent = v;
        select.appendChild(opt);
      }

      const desired = current && values.includes(current) ? current : "";
      select.value = desired;
      return desired || null;
    };

    const sortedIds = Array.from(ids).sort((a, b) => a.localeCompare(b));
    const sortedParentIds = Array.from(parentIds).sort((a, b) => a.localeCompare(b));

    this.hierarchyFilterValue = buildOptions(
      this.hierarchyFilter,
      sortedIds,
      this.localize("Toolbar.FilterHierarchyAll", "All hierarchy")
    );
    this.parentFilterValue = buildOptions(
      this.parentFilter,
      sortedParentIds,
      this.localize("Toolbar.FilterParentAll", "All parents")
    );

    const sortedDropdown = Array.from(dropdownValues).sort((a, b) => a.localeCompare(b));
    const dropdownPlaceholder = this.dropdownDisplayName
      ? `${this.localize("Toolbar.FilterDropdownAll", "All")}: ${this.dropdownDisplayName}`
      : this.localize("Toolbar.FilterDropdownAll", "All dropdown");

    this.dropdownFilterValue = buildOptions(
      this.dropdownFilter,
      sortedDropdown,
      dropdownPlaceholder
    );

    this.hasDropdownField = sortedDropdown.length > 0;
  }

  private setSparklineRange(rows: NodeRow[]): void {
    const values: number[] = [];
    for (const r of rows) {
      if (typeof r.sparkline === "number" && Number.isFinite(r.sparkline)) {
        values.push(r.sparkline);
      }
    }
    if (values.length === 0) {
      this.sparklineMin = null;
      this.sparklineMax = null;
      return;
    }
    this.sparklineMin = Math.min(...values);
    this.sparklineMax = Math.max(...values);
  }

  // ---------------------------
  // Collapse-aware layout
  // ---------------------------
  private computeLayoutFromState(autoFit = true, focusNodeId?: string): boolean {
    const visibleRows = this.computeVisibleRows(this.allRows, this.collapsedIds);
    if (visibleRows.length === 0) {
      this.nodes = [];
      this.links = [];
      this.tableRows = [];
      if (this.searchQuery || this.hierarchyFilterValue || this.parentFilterValue) {
        this.clearAndMessage(this.localize("Message.NoMatches", "No matches."));
      }
      return true;
    }
    const focusScreen = autoFit ? null : this.getNodeScreenPoint(focusNodeId ?? null);
    const ok = this.computeLayout(visibleRows, autoFit);
    if (!ok) return false;
    if (!autoFit && focusNodeId && focusScreen) {
      this.keepNodeScreenPosition(focusNodeId, focusScreen);
    }
    return true;
  }

  private computeVisibleRows(rows: NodeRow[], collapsed: Set<string>): NodeRow[] {
    const dropdownRows = this.applyDropdownFilters(rows);
    const filteredRows = this.applySearchFilter(dropdownRows, this.searchQuery);
    const sourceRows = filteredRows;
    const idSet = new Set<string>();
    for (const r of sourceRows) {
      idSet.add(r.id);
    }

    // roots: missing parent or parent not found
    const roots = sourceRows.filter(r => !r.parentId || !idSet.has(r.parentId));

    const visible = new Set<string>();
    const stack: string[] = [];

    for (const r of roots) stack.push(r.id);

    while (stack.length) {
      const id = stack.pop()!;
      if (visible.has(id)) continue;
      visible.add(id);

      // if collapsed, do not traverse descendants
      if (collapsed.has(id)) continue;

      const children = this.childrenMap.get(id) ?? [];
      for (const c of children) stack.push(c);
    }

    // Preserve original row order for stability
    return sourceRows.filter(r => visible.has(r.id));
  }

  private computeLayout(rows: NodeRow[], autoFit = true): boolean {
    const s = this.settings.layout;

    // Determine roots in the *visible* graph
    const idSet = new Set(rows.map(r => r.id));
    const roots = rows.filter(r => !r.parentId || !idSet.has(r.parentId));

    let working = rows.slice();
    let syntheticRootId: string | null = null;

    if (roots.length !== 1) {
      syntheticRootId = "__root__";
      const synthetic: NodeRow = {
        id: syntheticRootId,
        parentId: null,
        label: "All",
        selectionId: this.host.createSelectionIdBuilder().createSelectionId()
      };
      working = [...working, synthetic];

      for (const r of roots) {
        const idx = working.findIndex(x => x.id === r.id);
        if (idx >= 0) working[idx] = { ...working[idx], parentId: syntheticRootId };
      }
    }

    let root: HierarchyNode<NodeRow>;
    try {
      const strat = stratify<NodeRow>()
        .id(d => d.id)
        .parentId(d => (d.parentId ?? undefined));

      root = strat(working);
    } catch (e) {
      this.nodes = [];
      this.links = [];
      this.tableRows = [];
      const msg = this.localize(
        "Message.InvalidHierarchy",
        "Invalid hierarchy: duplicates, cycles, or missing parents."
      );
      this.host.displayWarningIcon?.(msg, msg);
      this.clearAndMessage(msg);
      return false;
    }

    const layout = tree<NodeRow>().nodeSize([
      s.cardWidth + s.siblingSpacing,
      s.cardHeight + s.levelSpacing
    ]);

    const laidOut = layout(root);

    const flat: LayoutNode[] = [];
    const ordered: LayoutNode[] = [];
    const links: Array<{ source: LayoutNode; target: LayoutNode }> = [];
    const map = new Map<HierarchyNode<NodeRow>, LayoutNode>();
    const depthOffset = syntheticRootId ? 1 : 0;

    laidOut.eachBefore((n) => {
      const d = n.data;
      const ln: LayoutNode = {
        id: d.id,
        label: d.label,
        value: d.value,
        sparkline: d.sparkline,
        tooltip: d.tooltip,
        x: n.x,
        y: n.y,
        selectionId: d.selectionId,
        depth: n.depth - depthOffset
      };
      map.set(n, ln);
      flat.push(ln);
      ordered.push(ln);
    });

    laidOut.each((n) => {
      const src = map.get(n);
      if (!src || !n.children) return;
      for (const c of n.children) {
        const tgt = map.get(c);
        if (!tgt) continue;
        tgt.parent = src;
        (src.children ??= []).push(tgt);
        links.push({ source: src, target: tgt });
      }
    });

    if (syntheticRootId) {
      this.nodes = flat.filter(n => n.id !== syntheticRootId);
      this.links = links.filter(l => l.source.id !== syntheticRootId);
      this.tableRows = ordered.filter(n => n.id !== syntheticRootId);
    } else {
      this.nodes = flat;
      this.links = links;
      this.tableRows = ordered;
    }

    if (s.orientation === "LR") {
      for (const n of this.nodes) {
        const t = n.x;
        n.x = n.y;
        n.y = t;
      }
    }

    if (autoFit) this.fitToViewport();
    return true;
  }

  private getNodeScreenPoint(nodeId: string | null): { x: number; y: number } | null {
    if (!nodeId) return null;
    const node = this.nodes.find(n => n.id === nodeId);
    if (!node) return null;
    return {
      x: node.x * this.scale + this.tx,
      y: node.y * this.scale + this.ty
    };
  }

  private keepNodeScreenPosition(nodeId: string, screenPoint: { x: number; y: number }): void {
    const node = this.nodes.find(n => n.id === nodeId);
    if (!node) return;
    this.tx = screenPoint.x - node.x * this.scale;
    this.ty = screenPoint.y - node.y * this.scale;
    this.updateZoomLabel();
  }

  private fitToViewport(): void {
    if (this.nodes.length === 0 || this.lastViewportW <= 0 || this.lastViewportH <= 0) return;

    const s = this.settings.layout;

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const n of this.nodes) {
      const x0 = n.x - s.cardWidth / 2;
      const y0 = n.y - s.cardHeight / 2;
      const x1 = x0 + s.cardWidth;
      const y1 = y0 + s.cardHeight;
      minX = Math.min(minX, x0);
      minY = Math.min(minY, y0);
      maxX = Math.max(maxX, x1);
      maxY = Math.max(maxY, y1);
    }

    const contentW = maxX - minX;
    const contentH = maxY - minY;

    const pad = 24;
    const availableW = Math.max(1, this.lastViewportW - pad * 2);
    const availableH = Math.max(1, this.lastViewportH - pad * 2);

    const sx = availableW / Math.max(1, contentW);
    const sy = availableH / Math.max(1, contentH);

    this.scale = Math.min(2, Math.max(0.2, Math.min(sx, sy)));

    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;

    this.tx = this.lastViewportW / 2 - cx * this.scale;
    this.ty = this.lastViewportH / 2 - cy * this.scale;
    this.updateZoomLabel();
  }

  private localize(key: string, fallback: string): string {
    const value = this.localizationManager?.getDisplayName(key);
    return value || fallback;
  }

  private createToolbar(): void {
    this.toolbar = document.createElement("div");
    this.toolbar.style.position = "absolute";
    this.toolbar.style.top = "8px";
    this.toolbar.style.right = "8px";
    this.toolbar.style.display = "flex";
    this.toolbar.style.gap = "6px";
    this.toolbar.style.zIndex = "10";
    this.toolbar.style.font = "600 11px Segoe UI";

    const makeGroup = (): HTMLDivElement => {
      const group = document.createElement("div");
      group.style.display = "flex";
      group.style.alignItems = "center";
      group.style.background = "#ffffff";
      group.style.border = "1px solid #d1d5db";
      group.style.borderRadius = "6px";
      group.style.boxShadow = "0 1px 2px rgba(0,0,0,0.08)";
      group.style.overflow = "hidden";
      return group;
    };

    const makeButton = (label: string, title: string): HTMLButtonElement => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = label;
      btn.title = title;
      btn.style.padding = "4px 6px";
      btn.style.border = "none";
      btn.style.background = "transparent";
      btn.style.cursor = "pointer";
      btn.style.color = "#111827";
      btn.style.font = "600 11px Segoe UI";
      btn.addEventListener("pointerdown", (e) => {
        e.stopPropagation();
      });
      return btn;
    };

    // Search group
    this.searchGroup = makeGroup();
    this.searchInput = document.createElement("input");
    this.searchInput.type = "text";
    this.searchInput.placeholder = this.localize("Toolbar.SearchPlaceholder", "Search");
    this.searchInput.setAttribute("aria-label", this.localize("Toolbar.SearchPlaceholder", "Search"));
    this.searchInput.style.border = "none";
    this.searchInput.style.outline = "none";
    this.searchInput.style.padding = "4px 6px";
    this.searchInput.style.font = "600 11px Segoe UI";
    this.searchInput.style.minWidth = "120px";
    this.searchInput.style.background = "transparent";
    this.searchInput.style.color = "#111827";
    this.searchInput.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
    });
    this.searchInput.addEventListener("input", () => {
      const value = this.searchInput.value ?? "";
      this.setSearchQuery(value);
    });

    const clearSearch = makeButton("×", this.localize("Toolbar.SearchClear", "Clear search"));
    clearSearch.style.fontSize = "12px";
    clearSearch.addEventListener("click", (e) => {
      e.stopPropagation();
      this.searchInput.value = "";
      this.setSearchQuery("");
    });

    this.searchGroup.appendChild(this.searchInput);
    this.searchGroup.appendChild(clearSearch);

    // Filter group (Hierarchy / Parent)
    const styleSelect = (select: HTMLSelectElement) => {
      select.style.border = "none";
      select.style.outline = "none";
      select.style.padding = "4px 6px";
      select.style.font = "600 11px Segoe UI";
      select.style.background = "transparent";
      select.style.color = "#111827";
      select.style.maxWidth = "140px";
    };

    this.filterGroup = makeGroup();
    this.hierarchyFilter = document.createElement("select");
    this.hierarchyFilter.setAttribute("aria-label", this.localize("Toolbar.FilterHierarchy", "Hierarchy filter"));
    styleSelect(this.hierarchyFilter);
    this.hierarchyFilter.addEventListener("pointerdown", (e) => e.stopPropagation());
    this.hierarchyFilter.addEventListener("change", () => {
      this.setHierarchyFilter(this.hierarchyFilter.value);
    });

    this.parentFilter = document.createElement("select");
    this.parentFilter.setAttribute("aria-label", this.localize("Toolbar.FilterParent", "Parent filter"));
    styleSelect(this.parentFilter);
    this.parentFilter.addEventListener("pointerdown", (e) => e.stopPropagation());
    this.parentFilter.addEventListener("change", () => {
      this.setParentFilter(this.parentFilter.value);
    });

    this.dropdownFilter = document.createElement("select");
    this.dropdownFilter.setAttribute("aria-label", this.localize("Toolbar.FilterDropdown", "Dropdown filter"));
    styleSelect(this.dropdownFilter);
    this.dropdownFilter.addEventListener("pointerdown", (e) => e.stopPropagation());
    this.dropdownFilter.addEventListener("change", () => {
      this.setDropdownFilter(this.dropdownFilter.value);
    });

    this.filterGroup.appendChild(this.hierarchyFilter);
    this.filterGroup.appendChild(this.parentFilter);
    this.filterGroup.appendChild(this.dropdownFilter);

    // View toggle group
    this.viewGroup = makeGroup();
    this.treeButton = makeButton(this.localize("Toolbar.Tree", "Tree"), "Tree view");
    this.tableButton = makeButton(this.localize("Toolbar.Table", "Table"), "Table view");
    this.viewGroup.appendChild(this.treeButton);
    this.viewGroup.appendChild(this.tableButton);

    this.treeButton.addEventListener("click", (e) => {
      e.stopPropagation();
      this.setViewMode("tree", true);
    });
    this.tableButton.addEventListener("click", (e) => {
      e.stopPropagation();
      this.setViewMode("table", true);
    });

    // Collapse/expand group
    this.collapseGroup = makeGroup();
    this.collapseAllButton = makeButton(this.localize("Toolbar.Collapse", "Collapse"), "Collapse all");
    this.expandAllButton = makeButton(this.localize("Toolbar.Expand", "Expand"), "Expand all");
    this.collapseGroup.appendChild(this.collapseAllButton);
    this.collapseGroup.appendChild(this.expandAllButton);

    this.collapseAllButton.addEventListener("click", (e) => {
      e.stopPropagation();
      this.collapseAll();
    });
    this.expandAllButton.addEventListener("click", (e) => {
      e.stopPropagation();
      this.expandAll();
    });

    // Zoom group
    this.zoomGroup = makeGroup();
    const zoomOut = makeButton("-", this.localize("Toolbar.ZoomOut", "Zoom out"));
    const zoomIn = makeButton("+", this.localize("Toolbar.ZoomIn", "Zoom in"));
    this.zoomLabel = document.createElement("input");
    this.zoomLabel.type = "text";
    this.zoomLabel.value = "100%";
    this.zoomLabel.inputMode = "numeric";
    this.zoomLabel.setAttribute("aria-label", this.localize("Toolbar.ZoomInput", "Zoom percent"));
    this.zoomLabel.style.padding = "2px 4px";
    this.zoomLabel.style.minWidth = "44px";
    this.zoomLabel.style.textAlign = "center";
    this.zoomLabel.style.border = "none";
    this.zoomLabel.style.outline = "none";
    this.zoomLabel.style.background = "transparent";
    this.zoomLabel.style.font = "600 11px Segoe UI";
    this.zoomLabel.style.color = "#111827";
    this.zoomLabel.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
    });

    this.zoomGroup.appendChild(zoomOut);
    this.zoomGroup.appendChild(this.zoomLabel);
    this.zoomGroup.appendChild(zoomIn);

    const commitZoom = () => {
      if (this.viewMode !== "tree") {
        this.updateZoomLabel();
        return;
      }

      const raw = this.zoomLabel.value.trim();
      const num = Number.parseFloat(raw.replace("%", ""));
      if (!Number.isFinite(num)) {
        this.updateZoomLabel();
        return;
      }

      const clamped = Math.min(400, Math.max(20, num));
      const next = clamped / 100;
      const prev = this.scale;
      if (next === prev) {
        this.updateZoomLabel();
        return;
      }

      const cx = this.lastViewportW / 2;
      const cy = this.lastViewportH / 2;
      this.tx = cx - (cx - this.tx) * (next / prev);
      this.ty = cy - (cy - this.ty) * (next / prev);
      this.scale = next;
      this.updateZoomLabel();
      this.renderTree(this.lastViewportW, this.lastViewportH);
    };

    this.zoomLabel.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      commitZoom();
      this.canvas.focus();
    });
    this.zoomLabel.addEventListener("blur", () => commitZoom());

    zoomOut.addEventListener("click", (e) => {
      e.stopPropagation();
      this.zoomBy(0.9);
    });
    zoomIn.addEventListener("click", (e) => {
      e.stopPropagation();
      this.zoomBy(1.1);
    });

    this.toolbar.appendChild(this.searchGroup);
    this.toolbar.appendChild(this.filterGroup);
    this.toolbar.appendChild(this.viewGroup);
    this.toolbar.appendChild(this.collapseGroup);
    this.toolbar.appendChild(this.zoomGroup);
    this.root.appendChild(this.toolbar);
  }

  private createTable(): void {
    this.tableContainer = document.createElement("div");
    this.tableContainer.style.position = "absolute";
    this.tableContainer.style.inset = "0";
    this.tableContainer.style.display = "none";
    this.tableContainer.style.overflow = "auto";
    this.tableContainer.style.font = "12px Segoe UI";
    this.tableContainer.tabIndex = 0;
    this.tableContainer.setAttribute("role", "grid");
    this.tableContainer.setAttribute("aria-label", this.localize("Visual.Name", "Hierarchy Flow"));

    this.tableEl = document.createElement("table");
    this.tableHead = document.createElement("thead");
    this.tableBody = document.createElement("tbody");

    this.tableEl.appendChild(this.tableHead);
    this.tableEl.appendChild(this.tableBody);
    this.tableContainer.appendChild(this.tableEl);
    this.root.appendChild(this.tableContainer);
  }

  private createLandingPage(): void {
    this.landingPage = document.createElement("div");
    this.landingPage.style.position = "absolute";
    this.landingPage.style.inset = "0";
    this.landingPage.style.display = "none";
    this.landingPage.style.alignItems = "center";
    this.landingPage.style.justifyContent = "center";
    this.landingPage.style.textAlign = "center";
    this.landingPage.style.padding = "24px";
    this.landingPage.style.font = "12px Segoe UI";
    this.landingPage.style.color = UI.subtext;
    this.landingPage.style.pointerEvents = "none";

    const title = document.createElement("div");
    title.style.font = "600 14px Segoe UI";
    title.style.marginBottom = "6px";
    title.textContent = this.localize("Landing.Title", "Build your hierarchy");

    const body = document.createElement("div");
    body.textContent = this.localize(
      "Landing.Body",
      "Add Hierarchy Field, Parent Field, and Name to start. Optional: Sparkline, Values, Tooltip."
    );

    this.landingPage.appendChild(title);
    this.landingPage.appendChild(body);
    this.root.appendChild(this.landingPage);
  }

  private showLandingPage(show: boolean): void {
    if (!this.landingPage) return;
    this.landingPage.style.display = show ? "flex" : "none";
    if (show) {
      this.canvas.style.display = "none";
      if (this.tableContainer) this.tableContainer.style.display = "none";
    }
  }

  private updateLandingPageStyle(): void {
    if (!this.landingPage) return;
    const palette = (this.host as any).colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette | undefined;
    const isHighContrast = palette?.isHighContrast;
    const background = isHighContrast
      ? palette?.background?.value
      : (this.settings?.appearance?.useBackground ? this.settings.appearance.backgroundColor : "transparent");
    const foreground = isHighContrast ? palette?.foreground?.value : UI.subtext;

    this.landingPage.style.background = background || "transparent";
    this.landingPage.style.color = foreground || UI.subtext;
  }

  private ensureFocus(): void {
    const list = this.viewMode === "table" ? this.tableRows : this.nodes;
    if (!list.length) {
      this.focusedNodeId = null;
      this.focusedIndex = 0;
      return;
    }
    let idx = list.findIndex(n => n.id === this.focusedNodeId);
    if (idx < 0) idx = 0;
    this.focusedIndex = idx;
    this.focusedNodeId = list[idx].id;
  }

  private applyToolbarSettings(): void {
    const controls = this.settings.controls;
    const showDropdown = controls.showDropdownFilter && this.hasDropdownField;
    this.toolbar.style.display = controls.showControls ? "flex" : "none";
    this.searchGroup.style.display = controls.showSearch ? "flex" : "none";
    this.filterGroup.style.display = (controls.showHierarchyFilter || controls.showParentFilter || showDropdown) ? "flex" : "none";
    this.hierarchyFilter.style.display = controls.showHierarchyFilter ? "inline-flex" : "none";
    this.parentFilter.style.display = controls.showParentFilter ? "inline-flex" : "none";
    this.dropdownFilter.style.display = showDropdown ? "inline-flex" : "none";
    this.viewGroup.style.display = controls.showViewToggle ? "flex" : "none";
    this.zoomGroup.style.display = controls.showZoom ? "flex" : "none";
    this.collapseGroup.style.display = controls.showCollapseExpand ? "flex" : "none";
    this.toolbar.style.pointerEvents = this.allowInteractions ? "auto" : "none";
    this.toolbar.style.opacity = this.allowInteractions ? "1" : "0.6";

    if (!controls.showSearch && this.searchQuery) {
      this.searchQuery = "";
      if (this.searchInput) this.searchInput.value = "";
      this.clearMessage();
      this.computeLayoutFromState(true);
      this.renderView();
    }
    if (!controls.showHierarchyFilter && this.hierarchyFilterValue) {
      this.hierarchyFilterValue = null;
      if (this.hierarchyFilter) this.hierarchyFilter.value = "";
      this.clearMessage();
      this.computeLayoutFromState(true);
      this.renderView();
    }
    if (!controls.showParentFilter && this.parentFilterValue) {
      this.parentFilterValue = null;
      if (this.parentFilter) this.parentFilter.value = "";
      this.clearMessage();
      this.computeLayoutFromState(true);
      this.renderView();
    }
    if ((!controls.showDropdownFilter || !this.hasDropdownField) && this.dropdownFilterValue) {
      this.dropdownFilterValue = null;
      if (this.dropdownFilter) this.dropdownFilter.value = "";
      this.clearMessage();
      this.computeLayoutFromState(true);
      this.renderView();
    }
  }

  private syncToolbarState(): void {
    if (!this.toolbar) return;
    this.updateZoomLabel();

    const activeBg = "#e5e7eb";
    const inactiveBg = "transparent";

    this.treeButton.style.background = this.viewMode === "tree" ? activeBg : inactiveBg;
    this.tableButton.style.background = this.viewMode === "table" ? activeBg : inactiveBg;

    if (this.viewMode === "table") {
      this.zoomGroup.style.display = "none";
    } else {
      this.zoomGroup.style.display = this.settings.controls.showZoom ? "flex" : "none";
    }
  }

  private setViewMode(mode: ViewMode, fromUser = false): void {
    if (this.viewMode === mode) return;
    this.viewMode = mode;
    if (fromUser) this.userSetView = true;
    this.hoveredId = null;
    this.ensureFocus();
    this.hideTooltip();
    this.renderView();
  }

  private zoomBy(factor: number): void {
    if (this.viewMode !== "tree") return;
    const prev = this.scale;
    const next = Math.min(4, Math.max(0.2, this.scale * factor));

    const cx = this.lastViewportW / 2;
    const cy = this.lastViewportH / 2;

    this.tx = cx - (cx - this.tx) * (next / prev);
    this.ty = cy - (cy - this.ty) * (next / prev);
    this.scale = next;
    this.updateZoomLabel();
    this.renderTree(this.lastViewportW, this.lastViewportH);
  }

  private updateZoomLabel(): void {
    if (!this.zoomLabel) return;
    this.zoomLabel.value = `${Math.round(this.scale * 100)}%`;
  }

  private collapseAll(): void {
    this.collapsedIds.clear();
    for (const [parentId] of this.childrenMap) this.collapsedIds.add(parentId);
    this.computeLayoutFromState();
    this.renderView();
  }

  private expandAll(): void {
    this.collapsedIds.clear();
    this.computeLayoutFromState();
    this.renderView();
  }

  // ---------------------------
  // Rendering
  // ---------------------------
  private renderView(): void {
    this.syncToolbarState();

    if (this.viewMode === "table") {
      this.canvas.style.display = "none";
      this.tableContainer.style.display = "block";
      this.renderTable();
      return;
    }

    this.tableContainer.style.display = "none";
    this.canvas.style.display = "block";
    this.renderTree(this.lastViewportW, this.lastViewportH);
  }

  private setSearchQuery(value: string): void {
    const trimmed = value.trim();
    if (this.searchQuery === trimmed) return;
    this.searchQuery = trimmed;
    this.clearMessage();
    this.hideTooltip();
    this.computeLayoutFromState(true);
    this.ensureFocus();
    this.renderView();
  }

  private setHierarchyFilter(value: string): void {
    const next = value || null;
    if (this.hierarchyFilterValue === next) return;
    this.hierarchyFilterValue = next;
    this.clearMessage();
    this.hideTooltip();
    this.computeLayoutFromState(true);
    this.ensureFocus();
    this.renderView();
  }

  private setParentFilter(value: string): void {
    const next = value || null;
    if (this.parentFilterValue === next) return;
    this.parentFilterValue = next;
    this.clearMessage();
    this.hideTooltip();
    this.computeLayoutFromState(true);
    this.ensureFocus();
    this.renderView();
  }

  private setDropdownFilter(value: string): void {
    const next = value || null;
    if (this.dropdownFilterValue === next) return;
    this.dropdownFilterValue = next;
    this.clearMessage();
    this.hideTooltip();
    this.computeLayoutFromState(true);
    this.ensureFocus();
    this.renderView();
  }

  private clearElement(el: HTMLElement): void {
    while (el.firstChild) {
      el.removeChild(el.firstChild);
    }
  }

  private renderTable(): void {
    const tableSettings = this.settings.table;
    const appearance = this.settings.appearance;
    const nodes = this.settings.nodes;
    const palette = (this.host as any).colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette | undefined;
    const accent = palette?.getColor?.("HierarchyFlowAccent")?.value ?? UI.accent;
    const isHighContrast = palette?.isHighContrast === true;
    const hcForeground = palette?.foreground?.value;
    const hcBackground = palette?.background?.value;
    const tableBackground = isHighContrast
      ? (hcBackground || UI.bg)
      : (appearance.useBackground ? appearance.backgroundColor : "transparent");
    const tableText = isHighContrast ? (hcForeground || UI.title) : (nodes.titleColor || UI.title);
    const headerBg = isHighContrast ? (hcBackground || "#ffffff") : (appearance.useBackground ? appearance.backgroundColor : "#ffffff");
    const headerBorder = isHighContrast ? (hcForeground || "#e5e7eb") : "#e5e7eb";
    const rowBorder = isHighContrast ? (hcForeground || "#e5e7eb") : "#f3f4f6";
    const valueText = isHighContrast ? (hcForeground || nodes.valueColor) : (nodes.valueColor || UI.subtext);
    const titleSpec = this.getTitleFontSpec();

    this.tableContainer.style.background = tableBackground;
    this.tableContainer.style.color = tableText;

    this.tableEl.style.width = "100%";
    this.tableEl.style.borderCollapse = "collapse";

    // header
    this.clearElement(this.tableHead);
    if (tableSettings.showHeader) {
      const headerRow = document.createElement("tr");
      const headers = [
        this.labelDisplayName || this.localize("Table.Header.Fields", "Fields"),
        this.valueDisplayName || this.localize("Table.Header.Value", "Value"),
        this.sparklineDisplayName || this.localize("Table.Header.Sparkline", "Sparkline")
      ];

      for (const text of headers) {
        const th = document.createElement("th");
        th.textContent = text;
        th.style.textAlign = "left";
        th.style.font = "600 12px Segoe UI";
        th.style.padding = "6px 8px";
        th.style.borderBottom = `1px solid ${headerBorder}`;
        th.style.position = "sticky";
        th.style.top = "0";
        th.style.background = headerBg;
        headerRow.appendChild(th);
      }
      this.tableHead.appendChild(headerRow);
    }

    // body
    this.clearElement(this.tableBody);
    const rowHeight = Math.max(20, tableSettings.rowHeight);

    for (let i = 0; i < this.tableRows.length; i++) {
      const row = this.tableRows[i];
      const tr = document.createElement("tr");
      tr.style.cursor = this.allowInteractions ? "pointer" : "default";

      const isSelected = this.selectedIds.has(row.id);
      if (isSelected) {
        if (isHighContrast) {
          tr.style.background = hcForeground || UI.linkActive;
          tr.style.color = hcBackground || UI.bg;
        } else {
          tr.style.background = UI.accentSoft;
        }
      } else if (!isHighContrast && tableSettings.zebra && i % 2 === 1) {
        tr.style.background = "#f9fafb";
      }

      if (row.id === this.focusedNodeId) {
        tr.style.outline = `2px dashed ${isHighContrast ? (hcForeground || UI.linkActive) : accent}`;
        tr.style.outlineOffset = "-2px";
      }

      const nameCell = document.createElement("td");
      nameCell.style.padding = "0 8px";
      nameCell.style.height = `${rowHeight}px`;
      nameCell.style.borderBottom = `1px solid ${rowBorder}`;
      nameCell.style.font = titleSpec.font;
      nameCell.style.textAlign = nodes.titleAlign;
      nameCell.style.whiteSpace = nodes.titleWrap ? "normal" : "nowrap";
      nameCell.style.overflow = "hidden";
      nameCell.style.textOverflow = nodes.titleWrap ? "clip" : "ellipsis";

      const indent = 8 + Math.max(0, row.depth) * 14;
      nameCell.style.paddingLeft = `${indent}px`;

      const hasChildren = (this.childrenMap.get(row.id)?.length ?? 0) > 0;
      if (hasChildren) {
        const toggle = document.createElement("span");
        toggle.textContent = this.collapsedIds.has(row.id) ? "+" : "–";
        toggle.style.display = "inline-block";
        toggle.style.width = "14px";
        toggle.style.height = "14px";
        toggle.style.lineHeight = "14px";
        toggle.style.textAlign = "center";
        toggle.style.marginRight = "6px";
        toggle.style.border = `1px solid ${isHighContrast ? (hcForeground || "#d1d5db") : "#d1d5db"}`;
        toggle.style.borderRadius = "3px";
        toggle.style.cursor = "pointer";
        toggle.addEventListener("click", (e) => {
          e.stopPropagation();
          this.toggleCollapse(row.id);
        });
        nameCell.appendChild(toggle);
      } else {
        const spacer = document.createElement("span");
        spacer.style.display = "inline-block";
        spacer.style.width = "20px";
        nameCell.appendChild(spacer);
      }

      const nameText = document.createElement("span");
      nameText.textContent = row.label;
      nameCell.appendChild(nameText);

      const valueCell = document.createElement("td");
      valueCell.style.padding = "0 8px";
      valueCell.style.height = `${rowHeight}px`;
      valueCell.style.borderBottom = `1px solid ${rowBorder}`;
      valueCell.style.color = valueText;
      valueCell.textContent = this.formatValue(row.value);

      const sparkCell = document.createElement("td");
      sparkCell.style.padding = "0 8px";
      sparkCell.style.height = `${rowHeight}px`;
      sparkCell.style.borderBottom = `1px solid ${rowBorder}`;
      sparkCell.style.color = valueText;
      const sparkValue = this.formatValue(row.sparkline);
      if (typeof row.sparkline === "number" && this.sparklineMin !== null) {
        const range = (this.sparklineMax ?? this.sparklineMin) - this.sparklineMin;
        const t = range === 0 ? 1 : (row.sparkline - this.sparklineMin) / range;
        const bar = document.createElement("div");
        bar.style.height = "6px";
        bar.style.width = `${Math.max(6, Math.round(60 * Math.max(0, Math.min(1, t))))}px`;
        bar.style.background = isHighContrast ? (hcForeground || accent) : accent;
        bar.style.borderRadius = "3px";
        sparkCell.appendChild(bar);
      } else {
        sparkCell.textContent = sparkValue;
      }


      tr.appendChild(nameCell);
      tr.appendChild(valueCell);
      tr.appendChild(sparkCell);

      tr.addEventListener("click", async (e) => {
        if (!this.allowInteractions) return;
        const isMulti = (e as MouseEvent).ctrlKey || (e as MouseEvent).metaKey;

        this.focusedNodeId = row.id;
        this.focusedIndex = i;
        this.tableContainer.focus();

        if (!isMulti) {
          this.selectedIds.clear();
          this.selectedIds.add(row.id);
        } else {
          if (this.selectedIds.has(row.id)) this.selectedIds.delete(row.id);
          else this.selectedIds.add(row.id);
        }

        await this.selectionManager.select(row.selectionId, isMulti);
        this.renderView();
      });

      this.tableBody.appendChild(tr);
    }
  }
  private renderTree(width: number, height: number): void {
    const ctx = this.ctx;
    const s = this.settings.layout;
    const appearance = this.settings.appearance;
    const lines = this.settings.lines;
    const nodes = this.settings.nodes;
    const levels = this.settings.levels;

    const palette = (this.host as any).colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette | undefined;
    const accent = palette?.getColor?.("HierarchyFlowAccent")?.value ?? UI.accent;
    const isHighContrast = palette?.isHighContrast === true;
    const hcForeground = palette?.foreground?.value;
    const hcBackground = palette?.background?.value;
    const lineColor = isHighContrast ? (hcForeground || lines.lineColor) : lines.lineColor;
    const activeLineColor = isHighContrast ? (hcForeground || lines.activeColor) : lines.activeColor;
    const nodeFillColor = isHighContrast ? (hcBackground || nodes.fillColor) : nodes.fillColor;
    const nodeStrokeColor = isHighContrast ? (hcForeground || nodes.strokeColor) : nodes.strokeColor;
    const titleColor = isHighContrast ? (hcForeground || nodes.titleColor) : nodes.titleColor;
    const valueColor = isHighContrast ? (hcForeground || nodes.valueColor) : nodes.valueColor;
    const tipStyle = lines.showArrows ? lines.tipStyle : "none";
    const tipSize = Math.max(2, lines.tipSize);

    // background
    ctx.clearRect(0, 0, width, height);
    if (isHighContrast) {
      ctx.fillStyle = hcBackground || UI.bg;
      ctx.fillRect(0, 0, width, height);
    } else if (appearance.useBackground) {
      ctx.fillStyle = appearance.backgroundColor || UI.bg;
      ctx.fillRect(0, 0, width, height);
    }

    ctx.save();
    ctx.translate(this.tx, this.ty);
    ctx.scale(this.scale, this.scale);

    // LINKS behind nodes
    const linkWidth = Math.max(0, lines.lineWidth);
    if (linkWidth > 0) {
      ctx.lineWidth = linkWidth / this.scale;
      ctx.setLineDash(lines.lineStyle === "dashed" ? [4 / this.scale, 3 / this.scale] : []);
      for (const l of this.links) {
        const active = this.hoveredId && (l.source.id === this.hoveredId || l.target.id === this.hoveredId);
      ctx.strokeStyle = active ? activeLineColor : lineColor;

        const x1 = l.source.x;
        const y1 = l.source.y;
        const x2 = l.target.x;
        const y2 = l.target.y;

        if (s.orientation === "TD") {
          const startY = y1 + s.cardHeight / 2;
          const endY = y2 - s.cardHeight / 2;
          const midY = (startY + endY) / 2;

          ctx.beginPath();
          ctx.moveTo(x1, startY);
          ctx.lineTo(x1, midY);
          ctx.lineTo(x2, midY);
          ctx.lineTo(x2, endY);
          ctx.stroke();

          if (tipStyle !== "none") this.drawLineTip(ctx, x2, endY, "down", tipStyle, tipSize);
        } else {
          const startX = x1 + s.cardWidth / 2;
          const endX = x2 - s.cardWidth / 2;
          const midX = (startX + endX) / 2;

          ctx.beginPath();
          ctx.moveTo(startX, y1);
          ctx.lineTo(midX, y1);
          ctx.lineTo(midX, y2);
          ctx.lineTo(endX, y2);
          ctx.stroke();

          if (tipStyle !== "none") this.drawLineTip(ctx, endX, y2, "right", tipStyle, tipSize);
        }
      }
      ctx.setLineDash([]);
    }

    // NODES
    this.nodeRects = [];
    this.toggleRects = [];

    for (const n of this.nodes) {
      const x = n.x - s.cardWidth / 2;
      const y = n.y - s.cardHeight / 2;
      const w = s.cardWidth;
      const h = s.cardHeight;

      const isHovered = this.hoveredId === n.id;
      const isSelected = this.selectedIds.has(n.id);
      const hasChildren = (this.childrenMap.get(n.id)?.length ?? 0) > 0;
      const toggleSize = 14;
      const textLeftPad = 6;
      const textRightPad = hasChildren ? (toggleSize + 10) : 6;
      const textWidth = Math.max(0, w - textLeftPad - textRightPad);

      // shadow
      ctx.save();
      if (nodes.showShadow && !isHighContrast) {
        ctx.shadowColor = UI.cardShadow;
        ctx.shadowBlur = (isHovered || isSelected) ? 10 : 6;
        ctx.shadowOffsetY = 2;
      } else {
        ctx.shadowColor = "transparent";
        ctx.shadowBlur = 0;
        ctx.shadowOffsetY = 0;
      }

      // card
      const baseFill = levels.enable
        ? (levels.levelColors[n.depth % levels.levelColors.length] || nodeFillColor)
        : nodeFillColor;
      ctx.fillStyle = isSelected ? UI.accentSoft : baseFill;
      ctx.strokeStyle = isSelected
        ? accent
        : (isHovered ? activeLineColor : nodeStrokeColor);
      ctx.lineWidth = Math.max(0, nodes.strokeWidth) / this.scale;

      const radius =
        nodes.shape === "pill"
          ? Math.min(w / 2, h / 2)
          : nodes.shape === "square"
            ? 0
            : Math.max(0, nodes.cornerRadius);

      this.roundRect(ctx, x, y, w, h, radius);
      ctx.fill();
      ctx.stroke();
      ctx.restore();

      if (this.focusedNodeId === n.id) {
        ctx.save();
        ctx.strokeStyle = isHighContrast ? (hcForeground || activeLineColor) : activeLineColor;
        ctx.lineWidth = 1.5 / this.scale;
        ctx.setLineDash([3 / this.scale, 2 / this.scale]);
        this.roundRect(ctx, x, y, w, h, radius);
        ctx.stroke();
        ctx.restore();
      }

      // title (wrap 2 lines, centered)
      const titleSpec = this.getTitleFontSpec();
      const align = nodes.titleAlign === "left" ? "left" : nodes.titleAlign === "right" ? "right" : "center";
      const textX =
        align === "left"
          ? x + textLeftPad
          : align === "right"
            ? x + textLeftPad + textWidth
            : x + textLeftPad + textWidth / 2;

      ctx.textBaseline = "top";
      ctx.textAlign = align;
      ctx.fillStyle = titleColor || UI.title;
      ctx.font = titleSpec.font;

      if (nodes.titleWrap) {
        this.drawWrappedText(ctx, n.label, textX, y + 6, textWidth, titleSpec.lineHeight, 2, align);
      } else {
        this.drawSingleLineText(ctx, n.label, textX, y + 6, textWidth);
      }

      // value line (optional)
      const valueText = this.formatValue(n.value);
      if (valueText) {
        ctx.fillStyle = valueColor || UI.subtext;
        ctx.font = FONT.value;
        ctx.textBaseline = "bottom";
        ctx.textAlign = "center";
        ctx.fillText(valueText, x + w / 2, y + h - 6);
      }

      // sparkline indicator (optional)
      if (typeof n.sparkline === "number" && Number.isFinite(n.sparkline) && this.sparklineMin !== null) {
        const range = (this.sparklineMax ?? this.sparklineMin) - this.sparklineMin;
        const t = range === 0 ? 1 : (n.sparkline - this.sparklineMin) / range;
        const lineW = (w - 16) * Math.max(0, Math.min(1, t));
        const sparkY = valueText ? y + h - 14 : y + h - 8;

        ctx.strokeStyle = isHighContrast ? (hcForeground || accent) : accent;
        ctx.lineWidth = 2 / this.scale;
        ctx.beginPath();
        ctx.moveTo(x + 8, sparkY);
        ctx.lineTo(x + 8 + lineW, sparkY);
        ctx.stroke();
      }

      // collapse toggle (+/-) if node has children in the full dataset
      if (hasChildren) {
        const tW = toggleSize;
        const tH = toggleSize;
        const tX = x + w - tW - 6;
        const tY = y + 6;

        ctx.fillStyle = isHighContrast ? (hcBackground || UI.toggleFill) : UI.toggleFill;
        ctx.strokeStyle = isHighContrast ? (hcForeground || UI.toggleStroke) : UI.toggleStroke;
        ctx.lineWidth = 1 / this.scale;
        this.roundRect(ctx, tX, tY, tW, tH, 3);
        ctx.fill();
        ctx.stroke();

        const isCollapsed = this.collapsedIds.has(n.id);
        ctx.fillStyle = isHighContrast ? (hcForeground || UI.toggleText) : UI.toggleText;
        ctx.font = FONT.toggle;
        ctx.textBaseline = "middle";
        ctx.textAlign = "center";
        ctx.fillText(isCollapsed ? "+" : "–", tX + tW / 2, tY + tH / 2 + 0.25);

        ctx.textAlign = "left";
        ctx.textBaseline = "top";

        this.toggleRects.push({ nodeId: n.id, x: tX, y: tY, w: tW, h: tH });
      }

      this.nodeRects.push({ node: n, x, y, w, h });
    }

    ctx.restore();
  }

  private roundRect(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number): void {
    const rr = Math.min(r, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(x + rr, y);
    ctx.arcTo(x + w, y, x + w, y + h, rr);
    ctx.arcTo(x + w, y + h, x, y + h, rr);
    ctx.arcTo(x, y + h, x, y, rr);
    ctx.arcTo(x, y, x + w, y, rr);
    ctx.closePath();
  }

  private drawLineTip(
    ctx: CanvasRenderingContext2D,
    x: number,
    y: number,
    dir: "down" | "right",
    style: "arrow" | "square" | "diamond" | "circle",
    sizePx: number
  ): void {
    const size = (sizePx / this.scale);
    ctx.save();
    ctx.fillStyle = ctx.strokeStyle as string;
    ctx.strokeStyle = ctx.strokeStyle as string;

    if (style === "arrow") {
      ctx.beginPath();
      if (dir === "down") {
        ctx.moveTo(x, y);
        ctx.lineTo(x - size, y - size * 1.5);
        ctx.lineTo(x + size, y - size * 1.5);
      } else {
        ctx.moveTo(x, y);
        ctx.lineTo(x - size * 1.5, y - size);
        ctx.lineTo(x - size * 1.5, y + size);
      }
      ctx.closePath();
      ctx.fill();
      ctx.restore();
      return;
    }

    ctx.translate(x, y);
    if (dir === "down") {
      ctx.rotate(Math.PI / 2);
    }

    if (style === "square") {
      const half = size;
      ctx.beginPath();
      ctx.rect(-half, -half, half * 2, half * 2);
      ctx.fill();
    } else if (style === "diamond") {
      const half = size;
      ctx.rotate(Math.PI / 4);
      ctx.beginPath();
      ctx.rect(-half, -half, half * 2, half * 2);
      ctx.fill();
    } else if (style === "circle") {
      ctx.beginPath();
      ctx.arc(0, 0, size, 0, Math.PI * 2);
      ctx.fill();
    }

    ctx.restore();
  }

  private ellipsis(text: string, max: number): string {
    if (!text) return "";
    return text.length <= max ? text : text.slice(0, max - 1) + "…";
  }

  private formatValue(value: number | string | null | undefined): string {
    if (value === null || value === undefined) return "";
    if (typeof value === "number" && Number.isFinite(value)) return value.toLocaleString();
    return String(value);
  }

  private getTitleFontSpec(): { font: string; size: number; lineHeight: number } {
    const nodes = this.settings.nodes;
    const size = Math.max(6, nodes.titleFontSize || 11);
    const family = (nodes.titleFontFamily || "Segoe UI").trim() || "Segoe UI";
    let fontStyle = "normal";
    let weight = "600";

    switch (nodes.titleFontStyle) {
      case "bold":
        weight = "700";
        break;
      case "italic":
        fontStyle = "italic";
        break;
      case "boldItalic":
        fontStyle = "italic";
        weight = "700";
        break;
      case "normal":
      default:
        fontStyle = "normal";
        weight = "600";
        break;
    }

    const font = `${fontStyle} ${weight} ${size}px ${family}`;
    return {
      font,
      size,
      lineHeight: Math.max(10, size + 2)
    };
  }

  private drawSingleLineText(ctx: CanvasRenderingContext2D, text: string, x: number, y: number, maxWidth: number): void {
    if (!text) return;
    const value = String(text);
    if (ctx.measureText(value).width <= maxWidth) {
      ctx.fillText(value, x, y);
      return;
    }

    const ellipsis = "…";
    let lo = 0;
    let hi = value.length;

    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2);
      const test = value.slice(0, mid) + ellipsis;
      if (ctx.measureText(test).width <= maxWidth) {
        lo = mid + 1;
      } else {
        hi = mid;
      }
    }

    const cut = Math.max(0, lo - 1);
    ctx.fillText(value.slice(0, cut) + ellipsis, x, y);
  }

  private drawWrappedText(
    ctx: CanvasRenderingContext2D,
    text: string,
    x: number,
    y: number,
    maxWidth: number,
    lineHeight: number,
    maxLines: number,
    align: CanvasTextAlign = "left"
  ): void {
    const prevAlign = ctx.textAlign;
    ctx.textAlign = align;

    const words = (text ?? "").split(/\s+/).filter(Boolean);
    if (words.length === 0) {
      ctx.textAlign = prevAlign;
      return;
    }

    let line = "";
    let lines = 0;

    for (let i = 0; i < words.length; i++) {
      const w = words[i];
      const test = line ? line + " " + w : w;

      if (ctx.measureText(test).width > maxWidth && line) {
        ctx.fillText(line, x, y + lines * lineHeight);
        lines++;

        if (lines >= maxLines - 1) {
          const remaining = words.slice(i).join(" ");
          ctx.fillText(this.ellipsis(remaining, 40), x, y + lines * lineHeight);
          ctx.textAlign = prevAlign;
          return;
        }

        line = w;
      } else {
        line = test;
      }
    }

    if (line && lines < maxLines) ctx.fillText(line, x, y + lines * lineHeight);

    ctx.textAlign = prevAlign;
  }

  private buildTooltipItems(node: LayoutNode): Array<{ displayName: string; value: string }> {
    const items: Array<{ displayName: string; value: string }> = [];

    if (node.label) items.push({ displayName: this.labelDisplayName, value: node.label });

    const valueText = this.formatValue(node.value);
    if (valueText) items.push({ displayName: this.valueDisplayName, value: valueText });

    const sparkText = this.formatValue(node.sparkline);
    if (sparkText) items.push({ displayName: this.sparklineDisplayName, value: sparkText });

    const tooltipText = this.formatValue(node.tooltip);
    if (tooltipText) items.push({ displayName: this.tooltipDisplayName, value: tooltipText });

    return items;
  }

  private showTooltip(hit: Hit, e: PointerEvent): void {
    if (!this.tooltipService || !this.tooltipService.enabled()) return;
    const items = this.buildTooltipItems(hit.node);
    if (items.length === 0) return;

    this.tooltipService.show({
      coordinates: [e.clientX, e.clientY],
      isTouchEvent: e.pointerType === "touch",
      dataItems: items as any,
      identities: [hit.node.selectionId]
    });
  }

  private moveTooltip(hit: Hit, e: PointerEvent): void {
    if (!this.tooltipService || !this.tooltipService.enabled()) return;
    this.tooltipService.move({
      coordinates: [e.clientX, e.clientY],
      isTouchEvent: e.pointerType === "touch",
      identities: [hit.node.selectionId]
    });
  }

  private hideTooltip(): void {
    if (!this.tooltipService || !this.tooltipService.enabled()) return;
    this.tooltipService.hide({
      isTouchEvent: false,
      immediately: true
    });
  }

  // ---------------------------
  // Messages
  // ---------------------------
  private clearAndMessage(msg: string): void {
    this.ctx.clearRect(0, 0, this.lastViewportW, this.lastViewportH);
    if (this.settings?.appearance?.useBackground) {
      this.ctx.fillStyle = this.settings.appearance.backgroundColor || UI.bg;
      this.ctx.fillRect(0, 0, this.lastViewportW, this.lastViewportH);
    }

    this.canvas.style.display = "block";
    if (this.tableContainer) this.tableContainer.style.display = "none";
    this.showLandingPage(false);

    let div = this.root.querySelector(".pbi-msg") as HTMLDivElement | null;
    if (!div) {
      div = document.createElement("div");
      div.className = "pbi-msg";
      div.style.position = "absolute";
      div.style.left = "0";
      div.style.top = "0";
      div.style.padding = "10px";
      div.style.pointerEvents = "none";
      this.root.style.position = "relative";
      this.root.appendChild(div);
    }
    const palette = (this.host as any).colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette | undefined;
    const msgColor = palette?.isHighContrast ? palette.foreground?.value : UI.subtext;
    div.style.color = msgColor || UI.subtext;
    div.textContent = msg;
  }

  private clearMessage(): void {
    const div = this.root.querySelector(".pbi-msg") as HTMLDivElement | null;
    if (div) div.remove();
  }

  // ---------------------------
  // Interactions
  // ---------------------------
  private onKeyDown(e: KeyboardEvent): void {
    const list = this.viewMode === "table" ? this.tableRows : this.nodes;
    if (!list.length) return;

    let idx = list.findIndex(n => n.id === this.focusedNodeId);
    if (idx < 0) idx = 0;

    let handled = true;

    switch (e.key) {
      case "ArrowDown":
      case "ArrowRight":
        idx = Math.min(list.length - 1, idx + 1);
        break;
      case "ArrowUp":
      case "ArrowLeft":
        idx = Math.max(0, idx - 1);
        break;
      case "Home":
        idx = 0;
        break;
      case "End":
        idx = list.length - 1;
        break;
      case "Enter":
      case " ":
        if (!this.allowInteractions) return;
        this.focusedNodeId = list[idx].id;
        this.focusedIndex = idx;
        this.selectedIds.clear();
        this.selectedIds.add(list[idx].id);
        this.selectionManager.select(list[idx].selectionId, false);
        this.renderView();
        e.preventDefault();
        return;
      case "+":
      case "=":
        this.focusedNodeId = list[idx].id;
        this.focusedIndex = idx;
        if (this.collapsedIds.has(list[idx].id)) this.toggleCollapse(list[idx].id);
        e.preventDefault();
        return;
      case "-":
      case "_":
        this.focusedNodeId = list[idx].id;
        this.focusedIndex = idx;
        if (!this.collapsedIds.has(list[idx].id)) this.toggleCollapse(list[idx].id);
        e.preventDefault();
        return;
      default:
        handled = false;
    }

    if (!handled) return;
    e.preventDefault();
    this.focusedNodeId = list[idx].id;
    this.focusedIndex = idx;
    this.renderView();
  }

  private wireInteractions(): void {
    let isPanning = false;
    let lastX = 0;
    let lastY = 0;

    this.canvas.addEventListener("pointerdown", (e) => {
      if (this.viewMode !== "tree") return;
      this.canvas.setPointerCapture(e.pointerId);
      isPanning = e.button === 1 || e.shiftKey; // middle mouse or Shift+drag
      lastX = e.clientX;
      lastY = e.clientY;
      this.hideTooltip();
    });

    this.canvas.addEventListener("pointermove", (e) => {
      if (this.viewMode !== "tree") return;
      if (isPanning) {
        const dx = e.clientX - lastX;
        const dy = e.clientY - lastY;
        lastX = e.clientX;
        lastY = e.clientY;
        this.tx += dx;
        this.ty += dy;
        this.renderTree(this.lastViewportW, this.lastViewportH);
        return;
      }

      const hit = this.hitTest(e.clientX, e.clientY);
      const nextHover = hit ? hit.node.id : null;
      const prevHover = this.hoveredId;

      if (nextHover !== this.hoveredId) {
        this.hoveredId = nextHover;
        this.canvas.style.cursor = hit ? "pointer" : "default";
        this.renderTree(this.lastViewportW, this.lastViewportH);
      }

      if (hit) {
        if (prevHover !== hit.node.id) this.showTooltip(hit, e);
        else this.moveTooltip(hit, e);
      } else {
        this.hideTooltip();
      }
    });

    this.canvas.addEventListener("pointerup", () => {
      if (this.viewMode !== "tree") return;
      isPanning = false;
    });

    this.canvas.addEventListener("pointerleave", () => {
      if (this.viewMode !== "tree") return;
      this.hoveredId = null;
      this.hideTooltip();
      this.renderTree(this.lastViewportW, this.lastViewportH);
    });

    this.canvas.addEventListener("wheel", (e) => {
      if (this.viewMode !== "tree") return;
      e.preventDefault();
      this.hideTooltip();

      const delta = -e.deltaY;
      const zoomFactor = delta > 0 ? 1.1 : 0.9;

      const prev = this.scale;
      const next = Math.min(4, Math.max(0.2, this.scale * zoomFactor));

      const rect = this.canvas.getBoundingClientRect();
      const mx = e.clientX - rect.left;
      const my = e.clientY - rect.top;

      this.tx = mx - (mx - this.tx) * (next / prev);
      this.ty = my - (my - this.ty) * (next / prev);
      this.scale = next;
      this.updateZoomLabel();

      this.renderTree(this.lastViewportW, this.lastViewportH);
    }, { passive: false });

    this.canvas.addEventListener("dblclick", (e) => {
      if (this.viewMode !== "tree") return;
      if (!this.allowInteractions) return;
      const hit = this.hitTest(e.clientX, e.clientY);
      if (!hit) return;
      if (this.isToggleHit(hit.node.id, hit.worldX, hit.worldY)) return;

      const zoomPercent = Math.max(10, this.settings.controls.doubleClickZoomPercent);
      const factor = zoomPercent / 100;
      const next = Math.min(4, Math.max(0.2, this.scale * factor));

      const cx = this.lastViewportW / 2;
      const cy = this.lastViewportH / 2;
      this.tx = cx - hit.node.x * next;
      this.ty = cy - hit.node.y * next;
      this.scale = next;
      this.updateZoomLabel();
      this.renderTree(this.lastViewportW, this.lastViewportH);
    });

    // Right-click context menu (addresses "Context Menu" recommendation)
    this.canvas.addEventListener("contextmenu", async (e) => {
      e.preventDefault();
      if (!this.allowInteractions) return;
      if (this.viewMode !== "tree") return;

      const hit = this.hitTest(e.clientX, e.clientY);
      if (!hit) return;

      // ensure selection matches the context target
      this.selectedIds.clear();
      this.selectedIds.add(hit.node.id);
      await this.selectionManager.select(hit.node.selectionId, false);

      // show Power BI context menu at mouse position
      const point = { x: e.clientX, y: e.clientY } as any;
      (this.selectionManager as any).showContextMenu?.(hit.node.selectionId, point);

      this.renderTree(this.lastViewportW, this.lastViewportH);
    });

    // Left-click: collapse toggle OR select
    this.canvas.addEventListener("click", async (e) => {
      if (e.shiftKey) return; // shift reserved for panning
      if (!this.allowInteractions) return;
      if (this.viewMode !== "tree") return;

      const hit = this.hitTest(e.clientX, e.clientY);
      if (!hit) {
        this.selectedIds.clear();
        this.focusedNodeId = null;
        await this.selectionManager.clear();
        this.renderTree(this.lastViewportW, this.lastViewportH);
        return;
      }

      this.focusedNodeId = hit.node.id;
      this.focusedIndex = this.nodes.findIndex(n => n.id === hit.node.id);
      this.canvas.focus();

      // First: check toggle (+/-) hit
      if (this.isToggleHit(hit.node.id, hit.worldX, hit.worldY)) {
        this.toggleCollapse(hit.node.id);
        return;
      }

      // Otherwise: normal selection
      const isMulti = e.ctrlKey || e.metaKey;

      if (!isMulti) {
        this.selectedIds.clear();
        this.selectedIds.add(hit.node.id);
      } else {
        if (this.selectedIds.has(hit.node.id)) this.selectedIds.delete(hit.node.id);
        else this.selectedIds.add(hit.node.id);
      }

      await this.selectionManager.select(hit.node.selectionId, isMulti);
      this.renderTree(this.lastViewportW, this.lastViewportH);
    });

    this.canvas.addEventListener("keydown", (e) => this.onKeyDown(e));
    this.tableContainer.addEventListener("keydown", (e) => this.onKeyDown(e));

    this.canvas.addEventListener("focus", () => {
      if (this.viewMode !== "tree") return;
      this.ensureFocus();
      this.renderTree(this.lastViewportW, this.lastViewportH);
    });

    this.tableContainer.addEventListener("focus", () => {
      if (this.viewMode !== "table") return;
      this.ensureFocus();
      this.renderTable();
    });
  }

  private toggleCollapse(nodeId: string): void {
    const hasChildren = (this.childrenMap.get(nodeId)?.length ?? 0) > 0;
    if (!hasChildren) return;

    if (this.collapsedIds.has(nodeId)) this.collapsedIds.delete(nodeId);
    else this.collapsedIds.add(nodeId);

    // recompute layout based on collapsed state
    this.computeLayoutFromState(false, nodeId);
    this.renderView();
  }

  private isToggleHit(nodeId: string, worldX: number, worldY: number): boolean {
    for (let i = this.toggleRects.length - 1; i >= 0; i--) {
      const t = this.toggleRects[i];
      if (t.nodeId !== nodeId) continue;
      if (worldX >= t.x && worldX <= t.x + t.w && worldY >= t.y && worldY <= t.y + t.h) return true;
    }
    return false;
  }

  private hitTest(clientX: number, clientY: number): Hit | null {
    const rect = this.canvas.getBoundingClientRect();
    const sx = clientX - rect.left;
    const sy = clientY - rect.top;

    // screen -> world coords
    const wx = (sx - this.tx) / this.scale;
    const wy = (sy - this.ty) / this.scale;

    for (let i = this.nodeRects.length - 1; i >= 0; i--) {
      const r = this.nodeRects[i];
      if (wx >= r.x && wx <= r.x + r.w && wy >= r.y && wy <= r.y + r.h) {
        return {
          node: r.node,
          worldX: wx,
          worldY: wy,
          localX: wx - r.x,
          localY: wy - r.y
        };
      }
    }
    return null;
  }
}
