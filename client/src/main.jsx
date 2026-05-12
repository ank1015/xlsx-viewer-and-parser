import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { CompactSelection, DataEditor, GridCellKind } from "@glideapps/glide-data-grid";
import { HugeiconsIcon } from "@hugeicons/react";
import {
  Download01Icon,
  FunctionOfXIcon,
  MinusSignIcon,
  PlusSignIcon,
} from "@hugeicons/core-free-icons";
import "@glideapps/glide-data-grid/dist/index.css";
import "@fontsource/geist-sans/400.css";
import "@fontsource/geist-sans/500.css";
import "@fontsource/geist-sans/600.css";
import "@fontsource/geist-sans/700.css";
import "./style.css";

const API_BASE_URL = "http://localhost:8000";
const fileNumbers = [1, 2, 3, 4, 5, 6, 7, 8];
const defaultRedesignNumber = 8;
const workbookZoomStorageKey = "openxml-web-workbook-zoom";

const defaultGridTheme = {
  accentColor: "#19736a",
  accentFg: "#ffffff",
  accentLight: "#d9f0ed",
  bgCell: "#ffffff",
  bgHeader: "#f5f7fa",
  borderColor: "#d2d8e2",
  cellHorizontalPadding: 10,
  cellVerticalPadding: 4,
  fontFamily: "Geist Sans, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif",
  headerFontStyle: "600 13px",
  textDark: "#172033",
  textHeader: "#536174",
  textMedium: "#536174",
};

const borderWidthByStyle = {
  dashDot: 1,
  dashDotDot: 1,
  dashed: 1,
  dotted: 1,
  double: 3,
  hair: 1,
  medium: 2,
  mediumDashDot: 2,
  mediumDashDotDot: 2,
  mediumDashed: 2,
  slantDashDot: 2,
  thick: 3,
  thin: 1,
};

const drawingColors = {
  chart: {
    bg: "rgba(92, 95, 190, 0.08)",
    border: "#5c5fbe",
    labelBg: "#eef0ff",
    labelText: "#31346f",
  },
  image: {
    bg: "rgba(188, 109, 36, 0.08)",
    border: "#bc6d24",
    labelBg: "#fff3e7",
    labelText: "#7a3e0e",
  },
};

const chartPalette = [
  "#2f6f73",
  "#d88c40",
  "#5c5fbe",
  "#87a83b",
  "#bd5d76",
  "#4590b8",
  "#9f6ab7",
  "#c3a33a",
];

const defaultExcelThemeColors = {
  0: "#ffffff",
  1: "#000000",
  2: "#e7e6e6",
  3: "#44546a",
  4: "#4472c4",
  5: "#ed7d31",
  6: "#a5a5a5",
  7: "#ffc000",
  8: "#5b9bd5",
  9: "#70ad47",
  10: "#0563c1",
  11: "#954f72",
};

const indexedExcelColors = {
  0: "#000000",
  1: "#ffffff",
  2: "#ff0000",
  3: "#00ff00",
  4: "#0000ff",
  5: "#ffff00",
  6: "#ff00ff",
  7: "#00ffff",
  8: "#000000",
  9: "#ffffff",
  10: "#ff0000",
  11: "#00ff00",
  12: "#0000ff",
  13: "#ffff00",
  14: "#ff00ff",
  15: "#00ffff",
  16: "#800000",
  17: "#008000",
  18: "#000080",
  19: "#808000",
  20: "#800080",
  21: "#008080",
  22: "#c0c0c0",
  23: "#808080",
  24: "#9999ff",
  25: "#993366",
  26: "#ffffcc",
  27: "#ccffff",
  28: "#660066",
  29: "#ff8080",
  30: "#0066cc",
  31: "#ccccff",
  32: "#000080",
  33: "#ff00ff",
  34: "#ffff00",
  35: "#00ffff",
  36: "#800080",
  37: "#800000",
  38: "#008080",
  39: "#0000ff",
  40: "#00ccff",
  41: "#ccffff",
  42: "#ccffcc",
  43: "#ffff99",
  44: "#99ccff",
  45: "#ff99cc",
  46: "#cc99ff",
  47: "#ffcc99",
  48: "#3366ff",
  49: "#33cccc",
  50: "#99cc00",
  51: "#ffcc00",
  52: "#ff9900",
  53: "#ff6600",
  54: "#666699",
  55: "#969696",
  56: "#003366",
  57: "#339966",
  58: "#003300",
  59: "#333300",
  60: "#993300",
  61: "#993366",
  62: "#333399",
  63: "#333333",
};

const minDrawingSpanRowHeight = 8;

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function getRouteNumber() {
  const routeNumber = Number(window.location.pathname.slice(1));
  return Number.isInteger(routeNumber) && routeNumber > 0 ? routeNumber : 1;
}

function getRedesignRouteNumber() {
  const match = window.location.pathname.match(/^\/redesign(?:\/(\d+))?\/?$/);
  const routeNumber = match?.[1] ? Number(match[1]) : defaultRedesignNumber;

  return fileNumbers.includes(routeNumber) ? routeNumber : defaultRedesignNumber;
}

function isRedesignPath() {
  return /^\/redesign(?:\/\d+)?\/?$/.test(window.location.pathname);
}

function getStoredWorkbookZoom() {
  try {
    const value = Number(window.localStorage.getItem(workbookZoomStorageKey));

    return Number.isFinite(value) ? clamp(value, 50, 200) : 100;
  } catch {
    return 100;
  }
}

function getInitialTheme() {
  return window.localStorage.getItem("openxml-viewer-theme") === "dark" ? "dark" : "light";
}

function getWorkbookSheets(workbook) {
  return workbook?.workbook?.sheets || [];
}

function getWorkbookTitle(workbook, fallbackTitle = "Workbook.xlsx") {
  return (
    workbook?.source?.fileName ||
    workbook?.source?.filename ||
    workbook?.source?.name ||
    workbook?.workbook?.source?.fileName ||
    workbook?.workbook?.name ||
    workbook?.name ||
    getWorkbookSheets(workbook)[0]?.name ||
    fallbackTitle
  );
}

function createEmptyGridSelection() {
  return {
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  };
}

function selectedCellFromDetails(details) {
  return details
    ? {
        address: details.address,
        drawings: details.drawings,
        formula: details.formula,
        merge: details.merge,
        sourceAddress: details.sourceAddress,
        style: details.style,
        styleIndex: details.styleIndex,
        table: details.table,
        type: details.type,
        value: details.value,
      }
    : null;
}

function getFormulaBarText(selectedCell) {
  if (!selectedCell) {
    return "";
  }

  return getFormulaText(selectedCell.formula) || selectedCell.value || "";
}

function getSelectionLabel(selection, gridData, fallback = "A1") {
  const range = selection?.current?.range;

  if (!range) {
    return fallback;
  }

  const startCol = Math.max(1, range.x);
  const endCol = Math.max(1, range.x + range.width - 1);
  const start = gridData.getCellMeta(startCol, range.y);
  const end = gridData.getCellMeta(endCol, range.y + range.height - 1);

  if (!start?.address) {
    return fallback;
  }

  if (!end?.address || end.address === start.address) {
    return start.address;
  }

  return `${start.address}:${end.address}`;
}

function safeDownloadName(name) {
  return (name || "workbook")
    .replace(/\.[^.]+$/, "")
    .replace(/[^a-z0-9_-]+/gi, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase() || "workbook";
}

function downloadWorkbookJson(workbook, title) {
  if (!workbook) {
    return;
  }

  const blob = new Blob([JSON.stringify(workbook, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = `${safeDownloadName(title)}.json`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function selectionFillRegions(selection) {
  const range = selection?.current?.range;
  const activeCell = selection?.current?.cell;

  if (!range || !activeCell) {
    return [];
  }

  const x = Math.max(1, range.x);
  const y = range.y;
  const width = Math.max(0, range.x + range.width - x);
  const height = Math.max(0, range.height);

  if (width === 0 || height === 0) {
    return [];
  }

  const activeX = activeCell[0];
  const activeY = activeCell[1];
  const activeInside =
    activeX >= x &&
    activeX < x + width &&
    activeY >= y &&
    activeY < y + height;

  if (!activeInside) {
    return [
      {
        color: "#DFEDFF",
        range: { height, width, x, y },
        style: "no-outline",
      },
    ];
  }

  return [
    activeY > y
      ? {
          color: "#DFEDFF",
          range: { height: activeY - y, width, x, y },
          style: "no-outline",
        }
      : null,
    activeY + 1 < y + height
      ? {
          color: "#DFEDFF",
          range: {
            height: y + height - activeY - 1,
            width,
            x,
            y: activeY + 1,
          },
          style: "no-outline",
        }
      : null,
    activeX > x
      ? {
          color: "#DFEDFF",
          range: { height: 1, width: activeX - x, x, y: activeY },
          style: "no-outline",
        }
      : null,
    activeX + 1 < x + width
      ? {
          color: "#DFEDFF",
          range: {
            height: 1,
            width: x + width - activeX - 1,
            x: activeX + 1,
            y: activeY,
          },
          style: "no-outline",
        }
      : null,
  ].filter(Boolean);
}

function drawSelectionOutline(ctx, rect, col, row, selection) {
  const range = selection?.current?.range;

  if (!range || col === 0) {
    return;
  }

  const x = Math.max(1, range.x);
  const y = range.y;
  const width = Math.max(0, range.x + range.width - x);
  const height = Math.max(0, range.height);

  if (
    width === 0 ||
    height === 0 ||
    col < x ||
    col >= x + width ||
    row < y ||
    row >= y + height
  ) {
    return;
  }

  const left = col === x;
  const right = col === x + width - 1;
  const top = row === y;
  const bottom = row === y + height - 1;

  if (!left && !right && !top && !bottom) {
    return;
  }

  ctx.save();
  ctx.strokeStyle = "#0285FF";
  ctx.lineWidth = 2.5;
  ctx.setLineDash([]);

  if (top) {
    drawStraightLine(ctx, rect.x, rect.y + 1.25, rect.x + rect.width, rect.y + 1.25);
  }

  if (right) {
    drawStraightLine(
      ctx,
      rect.x + rect.width - 1.25,
      rect.y,
      rect.x + rect.width - 1.25,
      rect.y + rect.height,
    );
  }

  if (bottom) {
    drawStraightLine(
      ctx,
      rect.x,
      rect.y + rect.height - 1.25,
      rect.x + rect.width,
      rect.y + rect.height - 1.25,
    );
  }

  if (left) {
    drawStraightLine(ctx, rect.x + 1.25, rect.y, rect.x + 1.25, rect.y + rect.height);
  }

  ctx.restore();
}

function columnNameToNumber(columnName) {
  return [...columnName].reduce(
    (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
    0,
  );
}

function numberToColumnName(columnNumber) {
  let name = "";
  let current = columnNumber;

  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }

  return name;
}

function parseCellAddress(address) {
  const match = /^([A-Z]+)(\d+)$/i.exec(address);

  if (!match) {
    return null;
  }

  return {
    column: columnNameToNumber(match[1].toUpperCase()),
    row: Number(match[2]),
  };
}

function parseDimension(dimension) {
  const [start, end = start] = String(dimension || "").split(":");
  const startAddress = parseCellAddress(start);
  const endAddress = parseCellAddress(end);

  if (!startAddress || !endAddress) {
    return null;
  }

  return {
    minColumn: Math.min(startAddress.column, endAddress.column),
    maxColumn: Math.max(startAddress.column, endAddress.column),
    minRow: Math.min(startAddress.row, endAddress.row),
    maxRow: Math.max(startAddress.row, endAddress.row),
  };
}

function getSheetBounds(sheet) {
  const cells = Object.values(sheet?.cells || {});
  const dimensionBounds = parseDimension(sheet?.dimension);

  if (dimensionBounds) {
    return dimensionBounds;
  }

  if (cells.length === 0) {
    return {
      minColumn: 1,
      maxColumn: 1,
      minRow: 1,
      maxRow: 1,
    };
  }

  return cells.reduce(
    (bounds, cell) => ({
      minColumn: Math.min(bounds.minColumn, cell.column),
      maxColumn: Math.max(bounds.maxColumn, cell.column),
      minRow: Math.min(bounds.minRow, cell.row),
      maxRow: Math.max(bounds.maxRow, cell.row),
    }),
    {
      minColumn: Number.POSITIVE_INFINITY,
      maxColumn: 1,
      minRow: Number.POSITIVE_INFINITY,
      maxRow: 1,
    },
  );
}

function getCellDisplayValue(cell) {
  if (!cell) {
    return "";
  }

  if (cell.displayValue !== undefined && cell.displayValue !== null) {
    return String(cell.displayValue);
  }

  if (cell.value !== undefined && cell.value !== null) {
    return String(cell.value);
  }

  if (cell.rawValue !== undefined && cell.rawValue !== null) {
    return String(cell.rawValue);
  }

  if (cell.formula?.cachedDisplayValue !== undefined && cell.formula.cachedDisplayValue !== null) {
    return String(cell.formula.cachedDisplayValue);
  }

  if (cell.formula?.cachedValue !== undefined && cell.formula.cachedValue !== null) {
    return String(cell.formula.cachedValue);
  }

  if (cell.formula?.cachedRawValue !== undefined && cell.formula.cachedRawValue !== null) {
    return String(cell.formula.cachedRawValue);
  }

  return "";
}

function getFormulaSummary(cell) {
  if (!cell?.formula) {
    return null;
  }

  return {
    address: cell.address,
    baseAddress: cell.formula.baseAddress,
    baseText: cell.formula.baseText,
    cachedValue: getCellDisplayValue(cell),
    kind: cell.formula.kind,
    reference: cell.formula.reference,
    resolvedText: cell.formula.resolvedText,
    sharedIndex: cell.formula.sharedIndex,
    text: cell.formula.text,
  };
}

function getFormulaText(formula) {
  return formula?.resolvedText || formula?.text || "";
}

function normalizeHexColor(hex) {
  if (!hex) {
    return null;
  }

  const value = String(hex).replace("#", "");
  const normalized = value.length === 8 ? value.slice(2) : value;

  if (!/^[0-9a-f]{6}$/i.test(normalized)) {
    return null;
  }

  return `#${normalized.toLowerCase()}`;
}

function applyTint(hex, tint) {
  const normalized = normalizeHexColor(hex);

  if (!normalized || tint === undefined || tint === null || tint === 0) {
    return normalized;
  }

  const amount = Number(tint);

  if (!Number.isFinite(amount)) {
    return normalized;
  }

  const channels = [1, 3, 5].map((start) => parseInt(normalized.slice(start, start + 2), 16));
  const tintedChannels = channels.map((channel) => {
    const adjusted = amount < 0
      ? channel * (1 + amount)
      : channel * (1 - amount) + 255 * amount;

    return clamp(Math.round(adjusted), 0, 255).toString(16).padStart(2, "0");
  });

  return `#${tintedChannels.join("")}`;
}

function colorToCss(color) {
  if (!color || color.auto) {
    return null;
  }

  const resolvedRgb = normalizeHexColor(color.resolvedRgb);

  if (resolvedRgb) {
    return resolvedRgb;
  }

  const rgb = normalizeHexColor(color.rgb);

  if (rgb) {
    return applyTint(rgb, color.tint);
  }

  if (color.theme !== undefined && color.theme !== null) {
    return applyTint(defaultExcelThemeColors[color.theme], color.tint);
  }

  return applyTint(indexedExcelColors[color.indexed], color.tint);
}

function itemByIndex(items, index) {
  if (!Array.isArray(items) || index === undefined || index === null) {
    return null;
  }

  return items.find((item) => item.index === index) || items[index] || null;
}

function resolveStyle(workbookStyles, styleIndex) {
  if (styleIndex === undefined || styleIndex === null) {
    return null;
  }

  return itemByIndex(workbookStyles?.cellFormats, styleIndex);
}

function resolveEffectiveStyle(cell, rowMeta, columnMeta, workbookStyles) {
  return (
    cell?.style ||
    resolveStyle(workbookStyles, cell?.styleIndex) ||
    resolveStyle(workbookStyles, rowMeta?.styleIndex) ||
    resolveStyle(workbookStyles, columnMeta?.styleIndex) ||
    null
  );
}

function resolveStyleParts(style, workbookStyles) {
  return {
    border: itemByIndex(workbookStyles?.borders, style?.borderId),
    fill: itemByIndex(workbookStyles?.fills, style?.fillId),
    font: itemByIndex(workbookStyles?.fonts, style?.fontId),
  };
}

function borderLine(side) {
  if (!side?.style) {
    return null;
  }

  return {
    color: colorToCss(side.color) || "#cfd6e2",
    style: side.style,
    width: borderWidthByStyle[side.style] || 1,
  };
}

function hasWrapText(style) {
  const wrapText = style?.wrapText ?? style?.alignment?.wrapText;

  return wrapText === true || wrapText === "1";
}

function buildCellPresentation(cell, style, workbookStyles, merge, table) {
  const { border, fill, font } = resolveStyleParts(style, workbookStyles);
  const fontColor = colorToCss(font?.color);
  const fillColor = fill?.patternType === "solid" ? colorToCss(fill.foregroundColor) : null;
  const themeOverride = {};
  const fontSize = font?.size || 13;

  if (font?.name) {
    themeOverride.fontFamily = defaultGridTheme.fontFamily;
  }

  if (font || style?.applyFont) {
    themeOverride.baseFontStyle = `${font?.italic ? "italic " : ""}${font?.bold ? "700" : "400"} ${fontSize}px`;
  }

  if (fontColor) {
    themeOverride.textDark = fontColor;
    themeOverride.textMedium = fontColor;
  }

  if ((style?.applyFill || fill?.patternType === "solid") && fillColor) {
    themeOverride.bgCell = fillColor;
  }

  if (table && !themeOverride.bgCell) {
    if (table.role === "header") {
      themeOverride.bgCell = "#e5f1f9";
    } else if (table.role === "totals") {
      themeOverride.bgCell = "#edf4ea";
    } else if (table.isStriped) {
      themeOverride.bgCell = "#f8fbfd";
    }
  }

  if ((table?.role === "header" || table?.role === "totals") && !themeOverride.baseFontStyle) {
    themeOverride.baseFontStyle = `700 ${fontSize}px`;
  }

  return {
    border: {
      bottom: borderLine(border?.bottom),
      left: borderLine(border?.left),
      right: borderLine(border?.right),
      top: borderLine(border?.top),
    },
    font: {
      size: fontSize,
      strike: Boolean(font?.strike),
      underline: Boolean(font?.underline),
    },
    themeOverride,
    textAlign: getContentAlign(cell, style, merge),
    wrapText: hasWrapText(style),
  };
}

function getContentAlign(cell, style, merge) {
  const horizontal = style?.alignment?.horizontal || style?.horizontalAlignment;

  if (horizontal === "center" || merge) {
    return "center";
  }

  if (horizontal === "right" || cell?.type === "number" || cell?.type === "date") {
    return "right";
  }

  return "left";
}

function excelColumnWidthToPixels(width) {
  if (!Number.isFinite(width)) {
    return 132;
  }

  return clamp(Math.round(width * 7 + 5), 36, 420);
}

function rowHeightToPixels(height) {
  if (!Number.isFinite(height)) {
    return 32;
  }

  return clamp(Math.round(height * (4 / 3)), 22, 260);
}

function buildRowMetaByIndex(sheet) {
  return new Map((sheet?.rows || []).map((row) => [row.index, row]));
}

function findColumnMeta(sheet, columnNumber) {
  return (sheet?.columns || []).find(
    (column) => column.min <= columnNumber && column.max >= columnNumber,
  );
}

function buildMergeByAddress(sheet) {
  const mergeByAddress = new Map();

  for (const merge of sheet?.mergedCells || []) {
    for (let row = merge.startRow; row <= merge.endRow; row += 1) {
      for (let column = merge.startColumn; column <= merge.endColumn; column += 1) {
        mergeByAddress.set(`${numberToColumnName(column)}${row}`, merge);
      }
    }
  }

  return mergeByAddress;
}

function normalizeTables(sheet) {
  return (sheet?.tables || [])
    .map((table) => {
      const bounds = parseDimension(table.reference);
      const referenceParts = table.reference?.split(":") || [];
      const startAddress =
        table.startAddress || (bounds ? referenceParts[0] : null);
      const endAddress =
        table.endAddress || (bounds ? referenceParts.at(-1) : null);

      return {
        ...table,
        displayName: table.displayName || table.name || `Table ${table.id || ""}`.trim(),
        endAddress,
        endColumn: table.endColumn ?? bounds?.maxColumn,
        endRow: table.endRow ?? bounds?.maxRow,
        startAddress,
        startColumn: table.startColumn ?? bounds?.minColumn,
        startRow: table.startRow ?? bounds?.minRow,
      };
    })
    .filter(
      (table) =>
        Number.isFinite(table.startColumn) &&
        Number.isFinite(table.endColumn) &&
        Number.isFinite(table.startRow) &&
        Number.isFinite(table.endRow),
    );
}

function findTableInfo(tables, row, column) {
  for (const table of tables) {
    if (
      row < table.startRow ||
      row > table.endRow ||
      column < table.startColumn ||
      column > table.endColumn
    ) {
      continue;
    }

    const headerEndRow = table.startRow + Math.max(0, table.headerRowCount || 0) - 1;
    const totalsStartRow = table.endRow - Math.max(0, table.totalsRowCount || 0) + 1;
    const role =
      table.headerRowCount > 0 && row <= headerEndRow
        ? "header"
        : table.totalsRowCount > 0 && row >= totalsStartRow
          ? "totals"
          : "body";
    const bodyRowOffset = row - headerEndRow - 1;

    return {
      bodyRowOffset,
      columnOffset: column - table.startColumn,
      isBottom: row === table.endRow,
      isHeaderBottom: role === "header" && row === headerEndRow,
      isLeft: column === table.startColumn,
      isRight: column === table.endColumn,
      isStriped:
        role === "body" &&
        table.style?.showRowStripes &&
        bodyRowOffset >= 0 &&
        bodyRowOffset % 2 === 1,
      isTop: row === table.startRow,
      isTotalsTop: role === "totals" && row === totalsStartRow,
      role,
      table,
    };
  }

  return null;
}

function markerRow(marker) {
  return marker?.row === undefined || marker?.row === null ? null : Number(marker.row);
}

function anchorToBounds(anchor) {
  if (!anchor) {
    return null;
  }

  if (anchor.from) {
    const startRow = markerRow(anchor.from);
    const startColumn = Number(anchor.from.column);
    const endRow = markerRow(anchor.to) ?? startRow;
    const endColumn = anchor.to ? Number(anchor.to.column) : startColumn;

    if (
      !Number.isFinite(startColumn) ||
      !Number.isFinite(endColumn) ||
      !Number.isFinite(startRow) ||
      !Number.isFinite(endRow)
    ) {
      return null;
    }

    return {
      endColumn: Math.max(startColumn, endColumn),
      endRow: Math.max(startRow, endRow),
      startColumn: Math.min(startColumn, endColumn),
      startRow: Math.min(startRow, endRow),
    };
  }

  return null;
}

function parseCellReference(reference) {
  if (!reference) {
    return null;
  }

  const rangePart = String(reference).split("!").at(-1)?.replaceAll("$", "");

  if (!rangePart) {
    return null;
  }

  return parseDimension(rangePart) || parseDimension(`${rangePart}:${rangePart}`);
}

function valuesFromRange(cells, reference, coerceNumber) {
  const bounds = parseCellReference(reference);

  if (!bounds) {
    return [];
  }

  const values = [];

  for (let row = bounds.minRow; row <= bounds.maxRow; row += 1) {
    for (let column = bounds.minColumn; column <= bounds.maxColumn; column += 1) {
      const cell = cells?.[`${numberToColumnName(column)}${row}`];
      const value = getCellDisplayValue(cell);

      if (coerceNumber) {
        const number = Number(String(value).replaceAll(",", ""));

        values.push(Number.isFinite(number) ? number : Number.NaN);
      } else if (value) {
        values.push(value);
      }
    }
  }

  return values;
}

function cellsFromRange(cells, reference) {
  const bounds = parseCellReference(reference);

  if (!bounds) {
    return [];
  }

  const items = [];

  for (let row = bounds.minRow; row <= bounds.maxRow; row += 1) {
    for (let column = bounds.minColumn; column <= bounds.maxColumn; column += 1) {
      items.push(cells?.[`${numberToColumnName(column)}${row}`] || null);
    }
  }

  return items;
}

function cellFillColor(cell, workbookStyles) {
  const style = resolveEffectiveStyle(cell, null, null, workbookStyles);
  const { fill } = resolveStyleParts(style, workbookStyles);

  return fill?.patternType === "solid" ? colorToCss(fill.foregroundColor) : null;
}

function colorsFromRange(cells, reference, workbookStyles) {
  const colors = cellsFromRange(cells, reference).map((cell) => cellFillColor(cell, workbookStyles));

  return colors.some(Boolean) ? colors : [];
}

function hydrateChartSeries(chart, cells, workbookStyles) {
  return {
    ...chart,
    series: (chart.series || []).map((series) => {
      const categoryColors = colorsFromRange(cells, series.categoriesRange, workbookStyles);
      const valueColors = colorsFromRange(cells, series.valuesRange, workbookStyles);
      const pointColors = categoryColors.length > 0 ? categoryColors : valueColors;

      return {
        ...series,
        cachedCategories:
          series.cachedCategories?.length > 0
            ? series.cachedCategories
            : valuesFromRange(cells, series.categoriesRange, false),
        cachedValues:
          series.cachedValues?.length > 0
            ? series.cachedValues
            : valuesFromRange(cells, series.valuesRange, true),
        pointColors,
        seriesColor: pointColors[0],
      };
    }),
  };
}

function drawingLabel(object) {
  if (object.kind === "chart") {
    return object.title || `${object.type || "chart"} ${object.id}`;
  }

  return object.name || object.fileName || object.id;
}

function drawingImageUrl(drawing) {
  return drawing.kind === "image" ? drawing.assetUrl : null;
}

function normalizeDrawingObjects(sheet, workbookStyles) {
  const charts = (sheet?.charts || []).map((chart) => {
    const hydratedChart = hydrateChartSeries(chart, sheet?.cells, workbookStyles);

    return {
      ...hydratedChart,
      bounds: anchorToBounds(hydratedChart.anchor),
      kind: "chart",
      label: drawingLabel({ ...hydratedChart, kind: "chart" }),
    };
  });
  const images = (sheet?.images || []).map((image) => ({
    ...image,
    bounds: anchorToBounds(image.anchor),
    kind: "image",
    label: drawingLabel({ ...image, kind: "image" }),
  }));

  return [...charts, ...images];
}

function clipDrawingObjectsToGrid(
  drawings,
  visibleRows,
  visibleColumns,
  columnSizeForNumber,
  rowSizeForNumber,
) {
  if (!visibleRows.length || !visibleColumns.length) {
    return drawings;
  }

  const columnAxis = buildVisibleAxis(visibleColumns, "number", "width", columnSizeForNumber);
  const rowAxis = buildVisibleAxis(visibleRows, "number", "height", rowSizeForNumber);

  return drawings
    .map((drawing) => {
      if (!drawing.bounds) {
        return drawing;
      }

      const pixelRect = anchorToPixelRect(drawing.anchor, columnAxis, rowAxis);

      if (!pixelRect) {
        return {
          ...drawing,
          cellRects: new Map(),
          clipBounds: null,
        };
      }

      const clipRect = intersectRect(pixelRect, {
        height: rowAxis.total,
        width: columnAxis.total,
        x: 0,
        y: 0,
      });
      const cellRects = clipRect
        ? buildDrawingCellRects(pixelRect, clipRect, visibleRows, visibleColumns, rowAxis, columnAxis)
        : new Map();
      const rowRects = buildDrawingRowRects(cellRects);
      const intersectingColumns = [...cellRects.values()].map((item) => item.column);
      const intersectingRows = [...cellRects.values()].map((item) => item.row);

      return {
        ...drawing,
        cellRects,
        clipBounds:
          intersectingColumns.length > 0 && intersectingRows.length > 0
            ? {
                endColumn: Math.max(...intersectingColumns),
                endRow: Math.max(...intersectingRows),
                startColumn: Math.min(...intersectingColumns),
                startRow: Math.min(...intersectingRows),
              }
            : null,
        pixelRect,
        rowRects,
        totalHeight: pixelRect.height,
        totalWidth: pixelRect.width,
      };
    })
    .filter((drawing) => !drawing.bounds || drawing.clipBounds);
}

function buildVisibleAxis(items, keyField, sizeField, sizeForKey) {
  const metrics = new Map();
  let total = 0;

  for (const item of items) {
    const key = item[keyField];
    const size = item[sizeField] || 1;

    metrics.set(key, {
      offset: total,
      size,
    });
    total += size;
  }

  return {
    items: items.map((item) => ({
      key: item[keyField],
      size: item[sizeField] || 1,
    })),
    metrics,
    sizeForKey,
    total,
  };
}

function emuToPixels(value) {
  return (Number(value) || 0) / 9525;
}

function markerToPixel(marker, axis) {
  if (!marker) {
    return null;
  }

  const coordinate = Number(marker.column ?? marker.row);
  const offset = emuToPixels(marker.columnOffsetEmu ?? marker.rowOffsetEmu);
  const exactMetric = axis.metrics.get(coordinate);

  if (exactMetric) {
    return exactMetric.offset + offset;
  }

  const first = axis.items[0];
  const last = axis.items.at(-1);

  if (!first || !last) {
    return null;
  }

  if (coordinate < first.key) {
    let position = 0;

    for (let key = first.key - 1; key >= coordinate; key -= 1) {
      position -= axis.sizeForKey?.(key) || first.size;
    }

    return position + offset;
  }

  if (coordinate > last.key) {
    let position = axis.total;

    for (let key = last.key + 1; key < coordinate; key += 1) {
      position += axis.sizeForKey?.(key) || last.size;
    }

    return position + offset;
  }

  return null;
}

function anchorToPixelRect(anchor, columnAxis, rowAxis) {
  const x1 = markerToPixel(
    anchor?.from
      ? {
          column: anchor.from.column,
          columnOffsetEmu: anchor.from.columnOffsetEmu,
        }
      : null,
    columnAxis,
  );
  const y1 = markerToPixel(
    anchor?.from
      ? {
          row: anchor.from.row,
          rowOffsetEmu: anchor.from.rowOffsetEmu,
        }
      : null,
    rowAxis,
  );
  const x2 = markerToPixel(
    anchor?.to
      ? {
          column: anchor.to.column,
          columnOffsetEmu: anchor.to.columnOffsetEmu,
        }
      : null,
    columnAxis,
  ) ?? (Number.isFinite(x1) && anchor?.cxEmu ? x1 + emuToPixels(anchor.cxEmu) : null);
  const y2 = markerToPixel(
    anchor?.to
      ? {
          row: anchor.to.row,
          rowOffsetEmu: anchor.to.rowOffsetEmu,
        }
      : null,
    rowAxis,
  ) ?? (Number.isFinite(y1) && anchor?.cyEmu ? y1 + emuToPixels(anchor.cyEmu) : null);

  if (![x1, y1, x2, y2].every(Number.isFinite)) {
    return null;
  }

  return {
    height: Math.max(1, Math.abs(y2 - y1)),
    width: Math.max(1, Math.abs(x2 - x1)),
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
  };
}

function intersectRect(a, b) {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.width, b.x + b.width);
  const y2 = Math.min(a.y + a.height, b.y + b.height);

  if (x2 <= x1 || y2 <= y1) {
    return null;
  }

  return {
    height: y2 - y1,
    width: x2 - x1,
    x: x1,
    y: y1,
  };
}

function buildDrawingCellRects(pixelRect, clipRect, visibleRows, visibleColumns, rowAxis, columnAxis) {
  const cellRects = new Map();

  for (const column of visibleColumns) {
    const columnMetric = columnAxis.metrics.get(column.number);

    if (!columnMetric) {
      continue;
    }

    for (const row of visibleRows) {
      const rowMetric = rowAxis.metrics.get(row.number);

      if (!rowMetric) {
        continue;
      }

      const cellRect = {
        height: rowMetric.size,
        width: columnMetric.size,
        x: columnMetric.offset,
        y: rowMetric.offset,
      };
      const intersection = intersectRect(cellRect, clipRect);

      if (!intersection) {
        continue;
      }

      cellRects.set(`${column.number}:${row.number}`, {
        column: column.number,
        height: intersection.height,
        row: row.number,
        sourceOffsetX: intersection.x - pixelRect.x,
        sourceOffsetY: intersection.y - pixelRect.y,
        width: intersection.width,
        xOffset: intersection.x - cellRect.x,
        yOffset: intersection.y - cellRect.y,
      });
    }
  }

  return cellRects;
}

function findDrawingInfos(drawings, row, column) {
  return drawings
    .map((drawing) => {
      const rowRect = drawing.rowRects?.get(row);
      const rect =
        rowRect && column === rowRect.column
          ? rowRect
          : rowRect && column >= rowRect.startColumn && column <= rowRect.endColumn
            ? null
            : drawing.cellRects?.get(`${column}:${row}`);

      if (!rect) {
        return null;
      }

      return {
        drawing,
        isBottom: rect.sourceOffsetY + rect.height >= drawing.totalHeight - 0.5,
        isLeft: rect.sourceOffsetX <= 0.5,
        isRight: rect.sourceOffsetX + rect.width >= drawing.totalWidth - 0.5,
        isTop: rect.sourceOffsetY <= 0.5,
        rect,
      };
    })
    .filter(Boolean);
}

function buildDrawingRowRects(cellRects) {
  const byRow = new Map();

  for (const rect of cellRects.values()) {
    const current = byRow.get(rect.row);
    const startX = rect.sourceOffsetX;
    const endX = rect.sourceOffsetX + rect.width;

    if (!current) {
      byRow.set(rect.row, {
        column: rect.column,
        endColumn: rect.column,
        height: rect.height,
        row: rect.row,
        sourceEndX: endX,
        sourceOffsetX: startX,
        sourceOffsetY: rect.sourceOffsetY,
        startColumn: rect.column,
        width: rect.width,
        xOffset: rect.xOffset,
        yOffset: rect.yOffset,
      });
      continue;
    }

    current.endColumn = Math.max(current.endColumn, rect.column);
    current.height = Math.max(current.height, rect.height);
    current.sourceEndX = Math.max(current.sourceEndX, endX);
    current.sourceOffsetX = Math.min(current.sourceOffsetX, startX);
    current.sourceOffsetY = Math.min(current.sourceOffsetY, rect.sourceOffsetY);

    if (rect.column < current.column) {
      current.column = rect.column;
      current.startColumn = rect.column;
      current.xOffset = rect.xOffset;
      current.yOffset = rect.yOffset;
    }

    current.width = current.sourceEndX - current.sourceOffsetX;
  }

  for (const rect of byRow.values()) {
    delete rect.sourceEndX;
  }

  return byRow;
}

function getDrawingVisibleSpan(drawings, row, column, visibleColumnIndexByNumber) {
  let bestSpan;

  for (const drawing of drawings) {
    const rowRect = drawing.rowRects?.get(row);

    if (
      !rowRect ||
      rowRect.height < minDrawingSpanRowHeight ||
      column < rowRect.startColumn ||
      column > rowRect.endColumn
    ) {
      continue;
    }

    const visibleIndexes = [];

    for (let item = rowRect.startColumn; item <= rowRect.endColumn; item += 1) {
      const index = visibleColumnIndexByNumber.get(item);

      if (index !== undefined) {
        visibleIndexes.push(index + 1);
      }
    }

    if (visibleIndexes.length < 2) {
      continue;
    }

    const span = [Math.min(...visibleIndexes), Math.max(...visibleIndexes)];

    if (!bestSpan || span[1] - span[0] > bestSpan[1] - bestSpan[0]) {
      bestSpan = span;
    }
  }

  return bestSpan;
}

function getVisibleSpan(merge, visibleColumnIndexByNumber) {
  const visibleIndexes = [];

  for (let column = merge.startColumn; column <= merge.endColumn; column += 1) {
    const index = visibleColumnIndexByNumber.get(column);

    if (index !== undefined) {
      visibleIndexes.push(index + 1);
    }
  }

  if (visibleIndexes.length < 2) {
    return undefined;
  }

  return [Math.min(...visibleIndexes), Math.max(...visibleIndexes)];
}

function estimateTextWidth(cell, text, rowInfo, columnInfo, workbookStyles) {
  const style = resolveEffectiveStyle(cell, rowInfo?.meta, columnInfo?.meta, workbookStyles);
  const { font } = resolveStyleParts(style, workbookStyles);
  const fontSize = font?.size || 13;
  const averageGlyphWidth = font?.bold ? 0.62 : 0.55;

  return text.length * fontSize * averageGlyphWidth + 20;
}

function buildTextOverflowSpans(
  sheet,
  visibleRows,
  visibleColumns,
  visibleColumnIndexByNumber,
  mergeByAddress,
  clippedDrawingObjects,
  workbookStyles,
) {
  const spans = new Map();

  for (const rowInfo of visibleRows) {
    for (let columnIndex = 0; columnIndex < visibleColumns.length; columnIndex += 1) {
      const columnInfo = visibleColumns[columnIndex];
      const address = `${columnInfo.name}${rowInfo.number}`;

      if (spans.has(address) || mergeByAddress.has(address)) {
        continue;
      }

      const cell = sheet.cells?.[address];
      const text = getCellDisplayValue(cell);

      if (!text) {
        continue;
      }

      const style = resolveEffectiveStyle(cell, rowInfo.meta, columnInfo.meta, workbookStyles);

      if (hasWrapText(style)) {
        continue;
      }

      if (getContentAlign(cell, style, null) !== "left") {
        continue;
      }

      const requiredWidth = estimateTextWidth(cell, text, rowInfo, columnInfo, workbookStyles);
      let availableWidth = columnInfo.width;
      let endColumnIndex = columnIndex;

      while (availableWidth < requiredWidth && endColumnIndex + 1 < visibleColumns.length) {
        const nextColumnIndex = endColumnIndex + 1;
        const nextColumn = visibleColumns[nextColumnIndex];
        const nextAddress = `${nextColumn.name}${rowInfo.number}`;

        if (
          mergeByAddress.has(nextAddress) ||
          getCellDisplayValue(sheet.cells?.[nextAddress]) ||
          getDrawingVisibleSpan(
            clippedDrawingObjects,
            rowInfo.number,
            nextColumn.number,
            visibleColumnIndexByNumber,
          )
        ) {
          break;
        }

        availableWidth += nextColumn.width;
        endColumnIndex = nextColumnIndex;
      }

      if (endColumnIndex <= columnIndex) {
        continue;
      }

      const span = [columnIndex + 1, endColumnIndex + 1];

      for (let item = columnIndex; item <= endColumnIndex; item += 1) {
        const spanAddress = `${visibleColumns[item].name}${rowInfo.number}`;

        spans.set(spanAddress, {
          role: item === columnIndex ? "start" : "covered",
          sourceAddress: address,
          span,
        });
      }
    }
  }

  return spans;
}

function buildGridData(sheet, workbookStyles) {
  if (!sheet) {
    return {
      columns: [],
      getCellContent: () => blankCell(),
      getCellMeta: () => null,
      drawingObjects: [],
      mergedCells: [],
      rows: [],
      tables: [],
    };
  }

  const bounds = getSheetBounds(sheet);
  const rowMetaByIndex = buildRowMetaByIndex(sheet);
  const mergeByAddress = buildMergeByAddress(sheet);
  const tables = normalizeTables(sheet);
  const drawingObjects = normalizeDrawingObjects(sheet, workbookStyles);
  const visibleColumnIndexByNumber = new Map();
  const visibleRows = [];
  const visibleColumns = [];

  for (let column = bounds.minColumn; column <= bounds.maxColumn; column += 1) {
    const columnMeta = findColumnMeta(sheet, column);

    if (columnMeta?.hidden) {
      continue;
    }

    visibleColumnIndexByNumber.set(column, visibleColumns.length);
    visibleColumns.push({
      meta: columnMeta,
      name: numberToColumnName(column),
      number: column,
      width: excelColumnWidthToPixels(columnMeta?.width),
    });
  }

  for (let row = bounds.minRow; row <= bounds.maxRow; row += 1) {
    const rowMeta = rowMetaByIndex.get(row);

    if (!rowMeta?.hidden) {
      visibleRows.push({
        height: rowHeightToPixels(rowMeta?.height),
        meta: rowMeta,
        number: row,
      });
    }
  }

  const clippedDrawingObjects = clipDrawingObjectsToGrid(
    drawingObjects,
    visibleRows,
    visibleColumns,
    (column) => excelColumnWidthToPixels(findColumnMeta(sheet, column)?.width),
    (row) => rowHeightToPixels(rowMetaByIndex.get(row)?.height),
  );
  const textOverflowByAddress = buildTextOverflowSpans(
    sheet,
    visibleRows,
    visibleColumns,
    visibleColumnIndexByNumber,
    mergeByAddress,
    clippedDrawingObjects,
    workbookStyles,
  );

  const columns = [
    {
      id: "__rowNumber",
      title: "#",
      width: 64,
    },
    ...visibleColumns.map((column) => ({
      id: column.name,
      title: column.name,
      width: column.width,
    })),
  ];

  const getCellDetails = (col, row) => {
    const rowInfo = visibleRows[row];

    if (!rowInfo) {
      return null;
    }

    if (col === 0) {
      return {
        address: String(rowInfo.number),
        cell: null,
        displayValue: String(rowInfo.number),
        isRowHeader: true,
        type: "row",
      };
    }

    const visibleColumn = visibleColumns[col - 1];

    if (!visibleColumn) {
      return null;
    }

    const address = `${visibleColumn.name}${rowInfo.number}`;
    const merge = mergeByAddress.get(address);
    const textOverflow = !merge ? textOverflowByAddress.get(address) : null;
    const table = findTableInfo(tables, rowInfo.number, visibleColumn.number);
    const drawings = findDrawingInfos(
      clippedDrawingObjects,
      rowInfo.number,
      visibleColumn.number,
    );
    const drawingSpan = !merge
      ? getDrawingVisibleSpan(
          clippedDrawingObjects,
          rowInfo.number,
          visibleColumn.number,
          visibleColumnIndexByNumber,
        )
      : undefined;
    const sourceAddress = merge?.startAddress || textOverflow?.sourceAddress || address;
    const sourceCell = sheet.cells?.[sourceAddress];
    const isTopMergeRow = !merge || rowInfo.number === merge.startRow;
    const displayValue = isTopMergeRow ? getCellDisplayValue(sourceCell) : "";
    const style = resolveEffectiveStyle(
      sourceCell,
      rowInfo.meta,
      visibleColumn.meta,
      workbookStyles,
    );
    const presentation = buildCellPresentation(sourceCell, style, workbookStyles, merge, table);
    const span = merge
      ? getVisibleSpan(merge, visibleColumnIndexByNumber)
      : drawingSpan || textOverflow?.span;
    const role = !merge
      ? null
      : address === merge.startAddress
        ? "start"
        : "covered";

    return {
      address,
      cell: sourceCell,
      displayValue,
      drawings,
      formula: getFormulaSummary(sourceCell),
      merge: merge
        ? {
            ...merge,
            role,
            sourceAddress: merge.startAddress,
          }
        : null,
      presentation,
      sourceAddress,
      span,
      style,
      styleIndex: sourceCell?.styleIndex ?? style?.index,
      table,
      type: sourceCell?.type || "blank",
      value: displayValue,
    };
  };

  return {
    columns,
    drawingObjects: clippedDrawingObjects,
    getCellContent: (item) => {
      const [col, row] = item;
      const details = getCellDetails(col, row);

      if (!details) {
        return blankCell();
      }

      if (details.isRowHeader) {
        return {
          allowOverlay: false,
          contentAlign: "right",
          data: details.displayValue,
          displayData: details.displayValue,
          kind: GridCellKind.Text,
          readonly: true,
          themeOverride: {
            bgCell: "#f5f7fa",
            textDark: "#536174",
          },
        };
      }

      return {
        allowOverlay: false,
        contentAlign: details.presentation.textAlign,
        copyData: details.displayValue,
        data: details.displayValue,
        displayData: details.displayValue,
        kind: GridCellKind.Text,
        readonly: true,
        allowWrapping: details.presentation.wrapText,
        span: details.span,
        themeOverride: details.presentation.themeOverride,
      };
    },
    getCellMeta: getCellDetails,
    mergedCells: sheet.mergedCells || [],
    rows: visibleRows,
    tables,
  };
}

function blankCell() {
  return {
    allowOverlay: false,
    data: "",
    displayData: "",
    kind: GridCellKind.Text,
    readonly: true,
  };
}

function lineDashForBorder(style) {
  if (style === "dotted") {
    return [1, 3];
  }

  if (style?.toLowerCase().includes("dash")) {
    return [6, 4];
  }

  return [];
}

function drawBorderLine(ctx, rect, side, border) {
  if (!border) {
    return;
  }

  const half = border.width / 2;
  ctx.beginPath();
  ctx.lineWidth = border.width;
  ctx.strokeStyle = border.color;
  ctx.setLineDash(lineDashForBorder(border.style));

  if (side === "top") {
    ctx.moveTo(rect.x, rect.y + half);
    ctx.lineTo(rect.x + rect.width, rect.y + half);
  } else if (side === "right") {
    ctx.moveTo(rect.x + rect.width - half, rect.y);
    ctx.lineTo(rect.x + rect.width - half, rect.y + rect.height);
  } else if (side === "bottom") {
    ctx.moveTo(rect.x, rect.y + rect.height - half);
    ctx.lineTo(rect.x + rect.width, rect.y + rect.height - half);
  } else if (side === "left") {
    ctx.moveTo(rect.x + half, rect.y);
    ctx.lineTo(rect.x + half, rect.y + rect.height);
  }

  ctx.stroke();
}

function drawTextDecoration(ctx, rect, details, theme) {
  if (!details?.displayValue || (!details.presentation.font.underline && !details.presentation.font.strike)) {
    return;
  }

  const padding = theme.cellHorizontalPadding || 10;
  const textWidth = Math.min(
    ctx.measureText(details.displayValue).width,
    Math.max(0, rect.width - padding * 2),
  );
  let x = rect.x + padding;

  if (details.presentation.textAlign === "center") {
    x = rect.x + (rect.width - textWidth) / 2;
  } else if (details.presentation.textAlign === "right") {
    x = rect.x + rect.width - padding - textWidth;
  }

  ctx.save();
  ctx.strokeStyle = theme.textDark;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);

  if (details.presentation.font.underline) {
    const y = rect.y + rect.height / 2 + details.presentation.font.size / 2 - 2;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + textWidth, y);
    ctx.stroke();
  }

  if (details.presentation.font.strike) {
    const y = rect.y + rect.height / 2;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + textWidth, y);
    ctx.stroke();
  }

  ctx.restore();
}

function drawTableOutline(ctx, rect, table) {
  if (!table) {
    return;
  }

  ctx.save();
  ctx.strokeStyle = "#19736a";
  ctx.lineWidth = 2;
  ctx.setLineDash([]);

  if (table.isTop) {
    drawStraightLine(ctx, rect.x, rect.y + 1, rect.x + rect.width, rect.y + 1);
  }

  if (table.isRight) {
    drawStraightLine(
      ctx,
      rect.x + rect.width - 1,
      rect.y,
      rect.x + rect.width - 1,
      rect.y + rect.height,
    );
  }

  if (table.isBottom) {
    drawStraightLine(
      ctx,
      rect.x,
      rect.y + rect.height - 1,
      rect.x + rect.width,
      rect.y + rect.height - 1,
    );
  }

  if (table.isLeft) {
    drawStraightLine(ctx, rect.x + 1, rect.y, rect.x + 1, rect.y + rect.height);
  }

  if (table.isHeaderBottom || table.isTotalsTop) {
    ctx.strokeStyle = "#2b8f86";
    ctx.lineWidth = 1.5;
    const y = table.isHeaderBottom ? rect.y + rect.height - 1 : rect.y + 1;
    drawStraightLine(ctx, rect.x, y, rect.x + rect.width, y);
  }

  ctx.restore();
}

function drawDrawingOverlays(ctx, rect, drawingInfos, theme, imageCache) {
  if (!drawingInfos?.length) {
    return;
  }

  for (const drawingInfo of drawingInfos) {
    const drawing = drawingInfo.drawing;
    const colors = drawingColors[drawing.kind] || drawingColors.chart;
    const imageUrl = drawingImageUrl(drawing);
    const imageEntry = imageUrl ? imageCache[imageUrl] : null;
    const hasLoadedImage = imageEntry?.status === "loaded" && imageEntry.image;
    const hasRenderableChart = drawing.kind === "chart" && getRenderableSeries(drawing).length > 0;

    ctx.save();

    if (hasLoadedImage) {
      drawImageTile(ctx, rect, drawingInfo, imageEntry.image);
    } else if (hasRenderableChart) {
      drawChartTile(ctx, rect, drawingInfo, theme);
    } else {
      ctx.fillStyle = colors.bg;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    }

    if (!hasLoadedImage && !hasRenderableChart && drawingInfo.isTop && drawingInfo.isLeft) {
      drawObjectLabel(ctx, rect, drawing, colors, theme);
    }

    ctx.restore();
  }
}

function getColumnOffset(columns, index) {
  let offset = 0;

  for (let item = 0; item < index && item < columns.length; item += 1) {
    offset += columns[item]?.width || 0;
  }

  return offset;
}

function getRowOffset(rows, index) {
  let offset = 0;

  for (let item = 0; item < index && item < rows.length; item += 1) {
    offset += rows[item]?.height || 0;
  }

  return offset;
}

function intersectsViewport(rect, width, height) {
  return (
    rect.x < width &&
    rect.y < height &&
    rect.x + rect.width > 0 &&
    rect.y + rect.height > 0
  );
}

function drawFullDrawingOverlay(ctx, drawing, rect, theme, imageCache) {
  const colors = drawingColors[drawing.kind] || drawingColors.chart;
  const imageUrl = drawingImageUrl(drawing);
  const imageEntry = imageUrl ? imageCache[imageUrl] : null;
  const hasLoadedImage = imageEntry?.status === "loaded" && imageEntry.image;
  const hasRenderableChart = drawing.kind === "chart" && getRenderableSeries(drawing).length > 0;

  ctx.save();

  if (hasLoadedImage) {
    ctx.drawImage(imageEntry.image, rect.x, rect.y, rect.width, rect.height);
  } else if (hasRenderableChart) {
    drawChart(ctx, drawing, rect.x, rect.y, rect.width, rect.height, theme);
  } else {
    ctx.fillStyle = colors.bg;
    ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    drawObjectLabel(ctx, rect, drawing, colors, theme);
  }

  ctx.restore();
}

function drawDrawingOverlayCanvas(canvas, gridData, visibleRegion, theme, imageCache, headerHeight = 32) {
  if (!canvas) {
    return;
  }

  const width = canvas.clientWidth;
  const height = canvas.clientHeight;
  const dpr = window.devicePixelRatio || 1;
  const pixelWidth = Math.max(1, Math.round(width * dpr));
  const pixelHeight = Math.max(1, Math.round(height * dpr));

  if (canvas.width !== pixelWidth || canvas.height !== pixelHeight) {
    canvas.width = pixelWidth;
    canvas.height = pixelHeight;
  }

  const ctx = canvas.getContext("2d");

  if (!ctx) {
    return;
  }

  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, width, height);

  const drawings = gridData.drawingObjects || [];

  if (drawings.length === 0 || width <= 0 || height <= 0) {
    return;
  }

  const frozenWidth = gridData.columns[0]?.width || 0;
  const firstVisibleColumn = Math.max(1, visibleRegion?.x ?? 1);
  const firstVisibleRow = Math.max(0, visibleRegion?.y ?? 0);
  const tx = visibleRegion?.tx ?? 0;
  const ty = visibleRegion?.ty ?? 0;
  const scrollX = getColumnOffset(gridData.columns, firstVisibleColumn) - frozenWidth - tx;
  const scrollY = getRowOffset(gridData.rows, firstVisibleRow) - ty;

  ctx.save();
  ctx.beginPath();
  ctx.rect(frozenWidth, headerHeight, Math.max(0, width - frozenWidth), Math.max(0, height - headerHeight));
  ctx.clip();

  for (const drawing of drawings) {
    if (!drawing.pixelRect) {
      continue;
    }

    const rect = {
      height: drawing.pixelRect.height,
      width: drawing.pixelRect.width,
      x: frozenWidth + drawing.pixelRect.x - scrollX,
      y: headerHeight + drawing.pixelRect.y - scrollY,
    };

    if (!intersectsViewport(rect, width, height)) {
      continue;
    }

    drawFullDrawingOverlay(ctx, drawing, rect, theme, imageCache);
  }

  ctx.restore();
}

function DrawingOverlay({ gridData, imageCache, theme, visibleRegion }) {
  const canvasRef = useRef(null);
  const [size, setSize] = useState({ height: 0, width: 0 });

  useEffect(() => {
    const canvas = canvasRef.current;
    const parent = canvas?.parentElement;

    if (!parent) {
      return undefined;
    }

    const resizeObserver = new ResizeObserver((entries) => {
      const rect = entries[0]?.contentRect;

      if (!rect) {
        return;
      }

      setSize({
        height: rect.height,
        width: rect.width,
      });
    });

    resizeObserver.observe(parent);

    return () => resizeObserver.disconnect();
  }, []);

  useEffect(() => {
    const canvas = canvasRef.current;
    let animationFrame = window.requestAnimationFrame(() => {
      drawDrawingOverlayCanvas(canvas, gridData, visibleRegion, theme, imageCache);
    });

    return () => window.cancelAnimationFrame(animationFrame);
  }, [gridData, imageCache, size, theme, visibleRegion]);

  return (
    <canvas
      aria-hidden="true"
      className="web-workbook-drawing-overlay"
      ref={canvasRef}
    />
  );
}

function expandedClipRect(rect, target) {
  const bleed = 1;

  return {
    height: Math.min(rect.height, target.height + bleed * 2),
    width: Math.min(rect.width, target.width + bleed * 2),
    x: rect.x + target.xOffset - bleed,
    y: rect.y + target.yOffset - bleed,
  };
}

function drawChartTile(ctx, rect, drawingInfo, theme) {
  const drawing = drawingInfo.drawing;
  const target = drawingInfo.rect;

  if (!target || !drawing.totalWidth || !drawing.totalHeight) {
    drawChart(ctx, drawing, rect.x, rect.y, rect.width, rect.height, theme);
    return;
  }

  ctx.save();
  const clip = expandedClipRect(rect, target);
  ctx.beginPath();
  ctx.rect(clip.x, clip.y, clip.width, clip.height);
  ctx.clip();
  ctx.translate(rect.x + target.xOffset - target.sourceOffsetX, rect.y + target.yOffset - target.sourceOffsetY);
  drawChart(ctx, drawing, 0, 0, drawing.totalWidth, drawing.totalHeight, theme);
  ctx.restore();
}

function drawImageTile(ctx, rect, drawingInfo, image) {
  const drawing = drawingInfo.drawing;
  const target = drawingInfo.rect;

  if (
    !target ||
    !drawing.totalWidth ||
    !drawing.totalHeight ||
    !image.naturalWidth ||
    !image.naturalHeight
  ) {
    ctx.drawImage(image, rect.x, rect.y, rect.width, rect.height);
    return;
  }

  const sourceX = (target.sourceOffsetX / drawing.totalWidth) * image.naturalWidth;
  const sourceY = (target.sourceOffsetY / drawing.totalHeight) * image.naturalHeight;
  const sourceWidth = (target.width / drawing.totalWidth) * image.naturalWidth;
  const sourceHeight = (target.height / drawing.totalHeight) * image.naturalHeight;

  ctx.drawImage(
    image,
    sourceX,
    sourceY,
    sourceWidth,
    sourceHeight,
    rect.x + target.xOffset,
    rect.y + target.yOffset,
    target.width,
    target.height,
  );
}

function getRenderableSeries(chart) {
  return (chart.series || [])
    .map((series) => ({
      ...series,
      cachedValues: series.cachedValues || [],
    }))
    .filter((series) => series.cachedValues.some(Number.isFinite));
}

function getChartCategories(series) {
  const source = series.find((item) => item.cachedCategories?.length > 0);
  const categoryCount = Math.max(0, ...series.map((item) => item.cachedValues.length));

  return Array.from({ length: categoryCount }, (_, index) => {
    return source?.cachedCategories?.[index] || `Item ${index + 1}`;
  });
}

function chartSeriesColor(series, seriesIndex) {
  return colorToCss(series.color) || series.seriesColor || chartPalette[seriesIndex % chartPalette.length];
}

function chartPointColor(series, pointIndex, seriesIndex) {
  return series.pointColors?.[pointIndex] || chartSeriesColor(series, seriesIndex);
}

function chartBarDirection(chart) {
  return chart.barDirection || chart.barDir || "col";
}

function chartBarGrouping(chart) {
  return chart.barGrouping || chart.grouping || "clustered";
}

function drawChart(ctx, chart, x, y, width, height, theme) {
  if (width < 48 || height < 36) {
    return;
  }

  const series = getRenderableSeries(chart);

  if (series.length === 0) {
    return;
  }

  ctx.save();
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(x, y, width, height);
  ctx.strokeStyle = "rgba(92, 95, 190, 0.18)";
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, width - 1, height - 1);

  const titleHeight = chart.title ? 24 : 8;
  const legendHeight = height > 120 ? 24 : 0;
  const plot = {
    height: Math.max(20, height - titleHeight - legendHeight - 14),
    width: Math.max(24, width - 22),
    x: x + 12,
    y: y + titleHeight,
  };

  if (chart.title) {
    ctx.fillStyle = theme.textDark || "#172033";
    ctx.font = `700 13px ${theme.fontFamily}`;
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    ctx.fillText(chart.title, x + width / 2, y + 7, width - 18);
  }

  if (chart.type === "pie" || chart.type === "doughnut") {
    if (hasInvalidPieValues(series[0])) {
      drawChartError(ctx, plot, "Pie values must be non-negative numbers", theme);
      ctx.restore();
      return;
    }

    drawPieChart(ctx, series[0], plot, chart.type);
  } else if (chart.type === "line" || chart.type === "area" || chart.type === "scatter") {
    drawLineChart(ctx, series, plot, chart.type);
  } else if (chartBarGrouping(chart) === "percentStacked" || chartBarGrouping(chart) === "stacked") {
    drawStackedBarChart(ctx, series, plot, theme, chartBarDirection(chart), chartBarGrouping(chart));
  } else if (chartBarDirection(chart) === "bar") {
    drawHorizontalBarChart(ctx, series, plot, theme);
  } else {
    drawBarChart(ctx, series, plot, theme);
  }

  if (legendHeight > 0) {
    drawLegend(ctx, series, x + 12, y + height - legendHeight, width - 24, legendHeight, theme);
  }

  ctx.restore();
}

function hasInvalidPieValues(series) {
  const values = series?.cachedValues || [];

  return values.length === 0 || values.some((value) => !Number.isFinite(value) || value < 0);
}

function drawChartError(ctx, plot, message, theme) {
  const centerX = plot.x + plot.width / 2;
  const centerY = plot.y + plot.height / 2;
  const iconSize = Math.min(34, Math.max(18, Math.min(plot.width, plot.height) * 0.22));

  ctx.save();
  ctx.strokeStyle = "rgba(180, 70, 58, 0.25)";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(plot.x, centerY);
  ctx.lineTo(plot.x + plot.width, centerY);
  ctx.stroke();

  ctx.fillStyle = "#d84a3a";
  ctx.beginPath();
  ctx.arc(centerX - Math.min(92, plot.width * 0.35), plot.y + 20, 8, 0, Math.PI * 2);
  ctx.fill();
  ctx.fillStyle = "#ffffff";
  ctx.font = `700 11px ${theme.fontFamily}`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("!", centerX - Math.min(92, plot.width * 0.35), plot.y + 20);

  ctx.fillStyle = "#d84a3a";
  ctx.font = `600 12px ${theme.fontFamily}`;
  ctx.textAlign = "left";
  ctx.fillText(message, centerX - Math.min(72, plot.width * 0.26), plot.y + 20, plot.width * 0.7);

  ctx.strokeStyle = "rgba(83, 97, 116, 0.18)";
  ctx.lineWidth = 5;
  ctx.strokeRect(centerX - iconSize / 2, centerY - iconSize / 2, iconSize, iconSize);
  ctx.lineWidth = 3;
  ctx.beginPath();
  ctx.moveTo(centerX - iconSize * 0.22, centerY + iconSize * 0.22);
  ctx.lineTo(centerX - iconSize * 0.22, centerY);
  ctx.moveTo(centerX, centerY + iconSize * 0.22);
  ctx.lineTo(centerX, centerY - iconSize * 0.18);
  ctx.moveTo(centerX + iconSize * 0.22, centerY + iconSize * 0.22);
  ctx.lineTo(centerX + iconSize * 0.22, centerY - iconSize * 0.05);
  ctx.stroke();
  ctx.restore();
}

function drawPieChart(ctx, series, plot, type) {
  const values = series.cachedValues.map((value) => Math.max(0, value));
  const total = values.reduce((sum, value) => sum + value, 0);
  const radius = Math.max(8, Math.min(plot.width, plot.height) * 0.38);
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;

  if (total <= 0) {
    drawEmptyPieChart(ctx, cx, cy, radius);
    return;
  }

  let angle = -Math.PI / 2;

  values.forEach((value, index) => {
    const slice = (value / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, angle, angle + slice);
    ctx.closePath();
    ctx.fillStyle = chartPointColor(series, index, index);
    ctx.fill();
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 1.5;
    ctx.stroke();
    angle += slice;
  });

  if (type === "doughnut") {
    ctx.beginPath();
    ctx.fillStyle = "#ffffff";
    ctx.arc(cx, cy, radius * 0.52, 0, Math.PI * 2);
    ctx.fill();
  }
}

function drawEmptyPieChart(ctx, cx, cy, radius) {
  const slice = (Math.PI * 2) / 3;
  let angle = -Math.PI / 2;

  for (let index = 0; index < 3; index += 1) {
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, angle, angle + slice);
    ctx.closePath();
    ctx.fillStyle = index % 2 === 0 ? "#eef2f7" : "#f8fafc";
    ctx.fill();
    ctx.strokeStyle = "#cfd6e2";
    ctx.lineWidth = 1;
    ctx.stroke();
    angle += slice;
  }

  ctx.beginPath();
  ctx.arc(cx, cy, radius, 0, Math.PI * 2);
  ctx.strokeStyle = "#9aa7b8";
  ctx.lineWidth = 1.5;
  ctx.stroke();
}

function chartValueExtent(series) {
  const values = series.flatMap((item) => item.cachedValues).filter(Number.isFinite);
  const min = Math.min(0, ...values);
  const max = Math.max(0, ...values);
  const padding = Math.max(1, (max - min) * 0.08);

  return {
    max: max + padding,
    min: min - padding,
  };
}

function valueToY(value, plot, extent) {
  const range = extent.max - extent.min || 1;
  return plot.y + plot.height - ((value - extent.min) / range) * plot.height;
}

function valueToX(value, plot, extent) {
  const range = extent.max - extent.min || 1;
  return plot.x + ((value - extent.min) / range) * plot.width;
}

function drawAxes(ctx, plot, extent) {
  const zeroY = valueToY(0, plot, extent);

  ctx.save();
  ctx.strokeStyle = "rgba(83, 97, 116, 0.24)";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(plot.x, plot.y);
  ctx.lineTo(plot.x, plot.y + plot.height);
  ctx.lineTo(plot.x + plot.width, plot.y + plot.height);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(plot.x, zeroY);
  ctx.lineTo(plot.x + plot.width, zeroY);
  ctx.stroke();
  ctx.restore();
}

function drawHorizontalAxes(ctx, plot, extent) {
  const zeroX = valueToX(0, plot, extent);

  ctx.save();
  ctx.strokeStyle = "rgba(83, 97, 116, 0.24)";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(plot.x, plot.y);
  ctx.lineTo(plot.x, plot.y + plot.height);
  ctx.lineTo(plot.x + plot.width, plot.y + plot.height);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(zeroX, plot.y);
  ctx.lineTo(zeroX, plot.y + plot.height);
  ctx.stroke();
  ctx.restore();
}

function drawCategoryLabels(ctx, categories, plot, theme, orientation) {
  ctx.save();
  ctx.fillStyle = theme.textMedium || "#536174";
  ctx.font = `600 9px ${theme.fontFamily}`;
  ctx.textBaseline = "middle";

  if (orientation === "horizontal") {
    const categoryHeight = plot.height / Math.max(1, categories.length);

    ctx.textAlign = "right";
    categories.forEach((category, index) => {
      ctx.fillText(
        category,
        plot.x - 5,
        plot.y + index * categoryHeight + categoryHeight / 2,
        Math.max(10, plot.labelWidth - 8),
      );
    });
  } else {
    const categoryWidth = plot.width / Math.max(1, categories.length);

    ctx.textAlign = "center";
    categories.forEach((category, index) => {
      ctx.fillText(
        category,
        plot.x + index * categoryWidth + categoryWidth / 2,
        plot.y + plot.height + plot.labelHeight / 2,
        Math.max(10, categoryWidth - 3),
      );
    });
  }

  ctx.restore();
}

function drawHorizontalBarChart(ctx, series, plot, theme) {
  const categories = getChartCategories(series);
  const labelWidth = plot.width > 82 ? clamp(plot.width * 0.32, 38, 96) : 0;
  const innerPlot = {
    ...plot,
    labelWidth,
    width: Math.max(12, plot.width - labelWidth),
    x: plot.x + labelWidth,
  };
  const extent = chartValueExtent(series);
  const zeroX = valueToX(0, innerPlot, extent);
  const categoryHeight = innerPlot.height / Math.max(1, categories.length);
  const groupHeight = Math.max(4, categoryHeight * 0.72);
  const barHeight = Math.max(2, groupHeight / Math.max(1, series.length));

  if (labelWidth > 0) {
    drawCategoryLabels(ctx, categories, innerPlot, theme, "horizontal");
  }

  drawHorizontalAxes(ctx, innerPlot, extent);

  categories.forEach((_, categoryIndex) => {
    const groupStart = innerPlot.y + categoryIndex * categoryHeight + (categoryHeight - groupHeight) / 2;

    series.forEach((item, seriesIndex) => {
      const value = item.cachedValues[categoryIndex];

      if (!Number.isFinite(value)) {
        return;
      }

      const valueX = valueToX(value, innerPlot, extent);
      const barX = Math.min(valueX, zeroX);
      const barY = groupStart + seriesIndex * barHeight;
      const barWidth = Math.max(1, Math.abs(valueX - zeroX));

      ctx.fillStyle = series.length === 1
        ? chartPointColor(item, categoryIndex, seriesIndex)
        : chartSeriesColor(item, seriesIndex);
      ctx.fillRect(barX, barY, barWidth, Math.max(1, barHeight - 1));
    });
  });
}

function stackTotal(series, categoryIndex) {
  return series.reduce((sum, item) => {
    const value = item.cachedValues[categoryIndex];

    return sum + (Number.isFinite(value) && value > 0 ? value : 0);
  }, 0);
}

function formatChartValue(value) {
  if (!Number.isFinite(value)) {
    return "";
  }

  return `$${Math.round(value).toLocaleString("en-US")}`;
}

function drawStackedBarChart(ctx, series, plot, theme, direction, grouping) {
  const categories = getChartCategories(series);

  if (direction === "bar") {
    drawHorizontalStackedBarChart(ctx, series, categories, plot, theme, grouping);
    return;
  }

  drawVerticalStackedBarChart(ctx, series, categories, plot, theme, grouping);
}

function drawVerticalStackedBarChart(ctx, series, categories, plot, theme, grouping) {
  const labelHeight = plot.height > 48 ? clamp(plot.height * 0.16, 14, 24) : 0;
  const innerPlot = {
    ...plot,
    height: Math.max(12, plot.height - labelHeight),
    labelHeight,
  };
  const categoryWidth = innerPlot.width / Math.max(1, categories.length);
  const barWidth = Math.max(5, categoryWidth * 0.58);
  const maxTotal = Math.max(1, ...categories.map((_, index) => stackTotal(series, index)));

  drawAxes(ctx, innerPlot, {
    max: grouping === "percentStacked" ? 1 : maxTotal,
    min: 0,
  });

  if (labelHeight > 0) {
    drawCategoryLabels(ctx, categories, innerPlot, theme, "vertical");
  }

  categories.forEach((_, categoryIndex) => {
    const total = stackTotal(series, categoryIndex);
    let yCursor = innerPlot.y + innerPlot.height;
    const barX = innerPlot.x + categoryIndex * categoryWidth + (categoryWidth - barWidth) / 2;

    series.forEach((item, seriesIndex) => {
      const value = item.cachedValues[categoryIndex];

      if (!Number.isFinite(value) || value <= 0 || total <= 0) {
        return;
      }

      const ratio = grouping === "percentStacked" ? value / total : value / maxTotal;
      const segmentHeight = Math.max(1, innerPlot.height * ratio);
      yCursor -= segmentHeight;

      ctx.fillStyle = chartSeriesColor(item, seriesIndex);
      ctx.fillRect(barX, yCursor, barWidth, segmentHeight);

      if (segmentHeight > 16 && barWidth > 34) {
        ctx.fillStyle = "#ffffff";
        ctx.font = `600 10px ${theme.fontFamily}`;
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(formatChartValue(value), barX + barWidth / 2, yCursor + segmentHeight / 2, barWidth - 4);
      }
    });
  });
}

function drawHorizontalStackedBarChart(ctx, series, categories, plot, theme, grouping) {
  const labelWidth = plot.width > 82 ? clamp(plot.width * 0.32, 38, 96) : 0;
  const innerPlot = {
    ...plot,
    labelWidth,
    width: Math.max(12, plot.width - labelWidth),
    x: plot.x + labelWidth,
  };
  const categoryHeight = innerPlot.height / Math.max(1, categories.length);
  const barHeight = Math.max(5, categoryHeight * 0.58);
  const maxTotal = Math.max(1, ...categories.map((_, index) => stackTotal(series, index)));

  drawHorizontalAxes(ctx, innerPlot, {
    max: grouping === "percentStacked" ? 1 : maxTotal,
    min: 0,
  });

  if (labelWidth > 0) {
    drawCategoryLabels(ctx, categories, innerPlot, theme, "horizontal");
  }

  categories.forEach((_, categoryIndex) => {
    const total = stackTotal(series, categoryIndex);
    let xCursor = innerPlot.x;
    const barY = innerPlot.y + categoryIndex * categoryHeight + (categoryHeight - barHeight) / 2;

    series.forEach((item, seriesIndex) => {
      const value = item.cachedValues[categoryIndex];

      if (!Number.isFinite(value) || value <= 0 || total <= 0) {
        return;
      }

      const ratio = grouping === "percentStacked" ? value / total : value / maxTotal;
      const segmentWidth = Math.max(1, innerPlot.width * ratio);

      ctx.fillStyle = chartSeriesColor(item, seriesIndex);
      ctx.fillRect(xCursor, barY, segmentWidth, barHeight);

      if (segmentWidth > 42 && barHeight > 13) {
        ctx.fillStyle = "#ffffff";
        ctx.font = `600 10px ${theme.fontFamily}`;
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(formatChartValue(value), xCursor + segmentWidth / 2, barY + barHeight / 2, segmentWidth - 4);
      }

      xCursor += segmentWidth;
    });
  });
}

function drawBarChart(ctx, series, plot, theme) {
  const categories = getChartCategories(series);
  const labelHeight = plot.height > 48 ? clamp(plot.height * 0.16, 14, 24) : 0;
  const innerPlot = {
    ...plot,
    height: Math.max(12, plot.height - labelHeight),
    labelHeight,
  };
  const extent = chartValueExtent(series);
  const zeroY = valueToY(0, innerPlot, extent);
  const categoryWidth = innerPlot.width / Math.max(1, categories.length);
  const groupWidth = Math.max(4, categoryWidth * 0.72);
  const barWidth = Math.max(2, groupWidth / Math.max(1, series.length));

  drawAxes(ctx, innerPlot, extent);

  if (labelHeight > 0) {
    drawCategoryLabels(ctx, categories, innerPlot, theme, "vertical");
  }

  categories.forEach((_, categoryIndex) => {
    const groupStart = innerPlot.x + categoryIndex * categoryWidth + (categoryWidth - groupWidth) / 2;

    series.forEach((item, seriesIndex) => {
      const value = item.cachedValues[categoryIndex];

      if (!Number.isFinite(value)) {
        return;
      }

      const valueY = valueToY(value, innerPlot, extent);
      const barX = groupStart + seriesIndex * barWidth;
      const barY = Math.min(valueY, zeroY);
      const barHeight = Math.max(1, Math.abs(zeroY - valueY));

      ctx.fillStyle = series.length === 1
        ? chartPointColor(item, categoryIndex, seriesIndex)
        : chartSeriesColor(item, seriesIndex);
      ctx.fillRect(barX, barY, Math.max(1, barWidth - 1), barHeight);
    });
  });
}

function drawLineChart(ctx, series, plot, type) {
  const categories = getChartCategories(series);
  const extent = chartValueExtent(series);
  const step = categories.length > 1 ? plot.width / (categories.length - 1) : plot.width;

  drawAxes(ctx, plot, extent);

  series.forEach((item, seriesIndex) => {
    const points = item.cachedValues
      .map((value, index) => ({
        x: plot.x + index * step,
        y: valueToY(value, plot, extent),
      }))
      .filter((point) => Number.isFinite(point.y));

    if (points.length === 0) {
      return;
    }

    ctx.strokeStyle = chartSeriesColor(item, seriesIndex);
    ctx.fillStyle = chartSeriesColor(item, seriesIndex);
    ctx.lineWidth = 2;
    ctx.beginPath();
    points.forEach((point, index) => {
      if (index === 0) {
        ctx.moveTo(point.x, point.y);
      } else {
        ctx.lineTo(point.x, point.y);
      }
    });
    ctx.stroke();

    if (type === "area") {
      ctx.lineTo(points.at(-1).x, plot.y + plot.height);
      ctx.lineTo(points[0].x, plot.y + plot.height);
      ctx.closePath();
      ctx.globalAlpha = 0.18;
      ctx.fill();
      ctx.globalAlpha = 1;
    }

    for (const point of points) {
      ctx.beginPath();
      ctx.arc(point.x, point.y, type === "scatter" ? 3 : 2.5, 0, Math.PI * 2);
      ctx.fill();
    }
  });
}

function drawLegend(ctx, series, x, y, width, height, theme) {
  const itemWidth = width / Math.min(series.length, 3);

  ctx.save();
  ctx.font = `600 10px ${theme.fontFamily}`;
  ctx.textAlign = "left";
  ctx.textBaseline = "middle";

  series.slice(0, 3).forEach((item, index) => {
    const itemX = x + index * itemWidth;

    ctx.fillStyle = chartSeriesColor(item, index);
    ctx.fillRect(itemX, y + height / 2 - 4, 8, 8);
    ctx.fillStyle = theme.textMedium || "#536174";
    ctx.fillText(item.name || `Series ${index + 1}`, itemX + 12, y + height / 2, itemWidth - 16);
  });

  ctx.restore();
}

function drawObjectLabel(ctx, rect, drawing, colors, theme) {
  const label = `${drawing.kind === "chart" ? "Chart" : "Image"}: ${drawing.label}`;
  const maxWidth = Math.max(0, rect.width - 12);

  if (!label || maxWidth < 24) {
    return;
  }

  ctx.save();
  ctx.font = `700 11px ${theme.fontFamily}`;
  const text = label.length > 34 ? `${label.slice(0, 31)}...` : label;
  const textWidth = Math.min(ctx.measureText(text).width, maxWidth - 12);
  const chipWidth = clamp(textWidth + 12, 24, maxWidth);
  const chipHeight = 22;
  const x = rect.x + 6;
  const y = rect.y + 6;

  ctx.fillStyle = colors.labelBg;
  ctx.fillRect(x, y, chipWidth, chipHeight);
  ctx.strokeStyle = colors.border;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);
  ctx.strokeRect(x + 0.5, y + 0.5, chipWidth - 1, chipHeight - 1);
  ctx.fillStyle = colors.labelText;
  ctx.textBaseline = "middle";
  ctx.fillText(text, x + 6, y + chipHeight / 2, chipWidth - 12);
  ctx.restore();
}

function drawStraightLine(ctx, x1, y1, x2, y2) {
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
}

function tableDetailText(table) {
  const parts = [
    table.reference,
    `${table.columns?.length || 0} column(s)`,
    table.autoFilter ? "auto-filter" : "",
    table.style?.name || "",
    table.totalsRowCount ? `${table.totalsRowCount} totals row(s)` : "",
  ].filter(Boolean);

  return parts.join(" | ");
}

function drawingDetailText(drawing) {
  if (drawing.kind === "chart") {
    const parts = [
      drawing.type || "chart",
      drawing.title || "",
      `${drawing.series?.length || 0} series`,
      drawing.assetPath || "",
    ].filter(Boolean);

    return parts.join(" | ");
  }

  return [drawing.contentType, drawing.fileName, drawing.description, drawing.assetPath]
    .filter(Boolean)
    .join(" | ");
}

function TableStrip({ tables }) {
  if (!tables.length) {
    return null;
  }

  return (
    <div className="table-strip" aria-label="Excel tables">
      {tables.map((table) => (
        <div
          className="table-pill"
          key={`${table.relationshipId}-${table.displayName}-${table.reference}`}
          title={tableDetailText(table)}
        >
          <span>{table.displayName}</span>
          <code>{table.reference}</code>
        </div>
      ))}
    </div>
  );
}

function DrawingStrip({ drawings }) {
  if (!drawings.length) {
    return null;
  }

  return (
    <div className="drawing-strip" aria-label="Worksheet drawings">
      {drawings.map((drawing) => (
        <div
          className={`drawing-pill ${drawing.kind}`}
          key={`${drawing.kind}-${drawing.id}-${drawing.relationshipId}`}
          title={drawingDetailText(drawing)}
        >
          <span>{drawing.kind === "chart" ? "Chart" : "Image"}</span>
          <code>{drawing.label}</code>
        </div>
      ))}
    </div>
  );
}

function FormulaBar({ gridData, selectedCell }) {
  if (!selectedCell) {
    return (
      <div className="formula-bar" data-testid="formula-bar">
        <span>Select a cell to inspect its value or formula.</span>
        {gridData.mergedCells.length > 0 ? (
          <span className="formula-meta">{gridData.mergedCells.length} merged range(s)</span>
        ) : null}
        {gridData.tables.length > 0 ? (
          <span className="formula-meta">{gridData.tables.length} table(s)</span>
        ) : null}
        {gridData.drawingObjects.length > 0 ? (
          <span className="formula-meta">
            {gridData.drawingObjects.filter((drawing) => drawing.kind === "chart").length} chart(s)
            {" | "}
            {gridData.drawingObjects.filter((drawing) => drawing.kind === "image").length} image(s)
          </span>
        ) : null}
      </div>
    );
  }

  const formulaText = getFormulaText(selectedCell.formula);

  return (
    <div className="formula-bar" data-testid="formula-bar">
      <span className="formula-address">{selectedCell.address}</span>
      {formulaText ? (
        <>
          <code>{formulaText}</code>
          {selectedCell.formula?.kind === "shared" &&
          selectedCell.formula?.baseAddress &&
          selectedCell.formula.baseAddress !== selectedCell.address ? (
            <span className="formula-meta">shared from {selectedCell.formula.baseAddress}</span>
          ) : null}
        </>
      ) : selectedCell.formula ? (
        <span>
          Shared formula #{selectedCell.formula.sharedIndex}; cached value{" "}
          <strong>{selectedCell.formula.cachedValue || "(blank)"}</strong>
        </span>
      ) : (
        <span>
          Value <strong>{selectedCell.value || "(blank)"}</strong>
        </span>
      )}
      {selectedCell.merge ? (
        <span className="formula-meta">
          merged {selectedCell.merge.reference}
          {selectedCell.merge.role === "covered" ? ` from ${selectedCell.merge.sourceAddress}` : ""}
        </span>
      ) : null}
      {selectedCell.table ? (
        <span className="formula-meta">
          table {selectedCell.table.table.displayName} | {selectedCell.table.role} |{" "}
          {selectedCell.table.table.reference}
        </span>
      ) : null}
      {selectedCell.drawings?.length > 0 ? (
        <span className="formula-meta">
          {selectedCell.drawings
            .map(({ drawing }) => `${drawing.kind} ${drawing.label}`)
            .join(" | ")}
        </span>
      ) : null}
      {selectedCell.style ? (
        <span className="formula-meta">
          {selectedCell.type}
          {selectedCell.style.numberFormatCode ? ` | ${selectedCell.style.numberFormatCode}` : ""}
          {selectedCell.styleIndex !== undefined ? ` | style ${selectedCell.styleIndex}` : ""}
        </span>
      ) : null}
    </div>
  );
}

function WebWorkbook({
  darkBgColor = "#0F0F11",
  formulaValue = "",
  iconButtonHoverStyle = {},
  iconButtonStyle = {},
  lightBgColor = "#FFFFFF",
  selectedRange = "A1",
  title = "Workbook.xlsx",
  titleStyle = {},
  workbook = null,
  onDownload = null,
  zoomTextStyle = {},
}) {
  const [zoom, setZoom] = useState(getStoredWorkbookZoom);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState(0);
  const [selectedCell, setSelectedCell] = useState(null);
  const [gridSelection, setGridSelection] = useState(createEmptyGridSelection);
  const [visibleRegion, setVisibleRegion] = useState(null);
  const [imageCache, setImageCache] = useState({});
  const imageLoadRef = useRef(new Set());
  const sheets = getWorkbookSheets(workbook);
  const workbookStyles = workbook?.workbook?.styles || workbook?.styles || {};
  const selectedSheet = sheets[selectedSheetIndex] || sheets[0];
  const gridData = useMemo(
    () => buildGridData(selectedSheet, workbookStyles),
    [selectedSheet, workbookStyles],
  );
  const zoomScale = zoom / 100;
  const gridTheme = useMemo(
    () => ({
      ...defaultGridTheme,
      accentColor: "#0285FF",
      accentLight: "rgba(2, 133, 255, 0)",
      fontFamily: defaultGridTheme.fontFamily,
      headerFontStyle: defaultGridTheme.headerFontStyle,
    }),
    [],
  );
  const imageUrlsKey = useMemo(
    () =>
      gridData.drawingObjects
        .map(drawingImageUrl)
        .filter(Boolean)
        .sort()
        .join("|"),
    [gridData],
  );
  const workbookTitle = getWorkbookTitle(workbook, title);
  const formulaBarValue = selectedCell ? getFormulaBarText(selectedCell) : formulaValue;
  const selectionLabel = getSelectionLabel(gridSelection, gridData, selectedRange);
  const highlightRegions = useMemo(() => selectionFillRegions(gridSelection), [gridSelection]);
  const handleDownload = useCallback(() => {
    if (typeof onDownload === "function") {
      onDownload({
        selectedSheet,
        selectedSheetIndex,
        title: workbookTitle,
        workbook,
      });
      return;
    }

    downloadWorkbookJson(workbook, workbookTitle);
  }, [onDownload, selectedSheet, selectedSheetIndex, workbook, workbookTitle]);

  useEffect(() => {
    try {
      window.localStorage.setItem(workbookZoomStorageKey, String(zoom));
    } catch {
      // Ignore storage failures; zoom still works for the current session.
    }
  }, [zoom]);

  useEffect(() => {
    setSelectedSheetIndex(0);
  }, [workbook]);

  useEffect(() => {
    setSelectedCell(null);
    setGridSelection(createEmptyGridSelection());
    setVisibleRegion(null);
  }, [selectedSheet]);

  useEffect(() => {
    const urls = imageUrlsKey ? imageUrlsKey.split("|").filter(Boolean) : [];

    for (const url of urls) {
      if (imageLoadRef.current.has(url)) {
        continue;
      }

      imageLoadRef.current.add(url);
      setImageCache((current) => ({
        ...current,
        [url]: {
          status: "loading",
        },
      }));

      const image = new Image();
      image.crossOrigin = "anonymous";
      image.onload = () => {
        setImageCache((current) => ({
          ...current,
          [url]: {
            image,
            status: "loaded",
          },
        }));
      };
      image.onerror = () => {
        setImageCache((current) => ({
          ...current,
          [url]: {
            status: "error",
          },
        }));
      };
      image.src = url;
    }
  }, [imageUrlsKey]);

  const updateSelectedCellFromSelection = useCallback(
    (selection) => {
      const current = selection?.current;
      const activeCell = current?.cell;
      const range = current?.range;
      const col = activeCell?.[0] ?? Math.max(1, range?.x ?? 1);
      const row = activeCell?.[1] ?? range?.y ?? 0;

      if (col === 0) {
        setSelectedCell(null);
        return;
      }

      setSelectedCell(selectedCellFromDetails(gridData.getCellMeta(col, row)));
    },
    [gridData],
  );

  const handleGridSelectionChange = useCallback(
    (newSelection) => {
      setGridSelection(newSelection);
      updateSelectedCellFromSelection(newSelection);
    },
    [updateSelectedCellFromSelection],
  );

  const handleCellClicked = useCallback(
    (item) => {
      const [col, row] = item;

      if (col === 0) {
        setSelectedCell(null);
        return;
      }

      setSelectedCell(selectedCellFromDetails(gridData.getCellMeta(col, row)));
    },
    [gridData],
  );

  const handleVisibleRegionChanged = useCallback((region, tx = 0, ty = 0) => {
    setVisibleRegion({
      ...region,
      tx,
      ty,
    });
  }, []);

  const drawCell = useCallback(
    (args, drawContent) => {
      const details = gridData.getCellMeta(args.col, args.row);

      drawContent();

      if (!details || details.isRowHeader) {
        return;
      }

      drawTextDecoration(args.ctx, args.rect, details, args.theme);

      const border = details.presentation.border;
      args.ctx.save();
      drawBorderLine(args.ctx, args.rect, "top", border.top);
      drawBorderLine(args.ctx, args.rect, "right", border.right);
      drawBorderLine(args.ctx, args.rect, "bottom", border.bottom);
      drawBorderLine(args.ctx, args.rect, "left", border.left);
      args.ctx.restore();

      drawTableOutline(args.ctx, args.rect, details.table);
      drawSelectionOutline(args.ctx, args.rect, args.col, args.row, gridSelection);
    },
    [gridData, gridSelection],
  );

  return (
    <section
      aria-label="Web workbook"
      className="web-workbook"
      style={{
        "--web-workbook-dark-bg": darkBgColor,
        "--web-workbook-icon-hover-bg":
          iconButtonHoverStyle.backgroundColor || iconButtonHoverStyle.background,
        "--web-workbook-icon-hover-color": iconButtonHoverStyle.color,
        "--web-workbook-light-bg": lightBgColor,
      }}
    >
      <header className="web-workbook-header">
        <div className="web-workbook-title" style={titleStyle} title={workbookTitle}>
          {workbookTitle}
        </div>

        <div className="web-workbook-actions" aria-label="Workbook controls">
          <div className="web-workbook-zoom" aria-label="Zoom controls">
            <button
              aria-label="Zoom in"
              className="web-workbook-icon-button"
              onClick={() => setZoom((current) => clamp(current + 10, 50, 200))}
              style={iconButtonStyle}
              type="button"
            >
              <HugeiconsIcon icon={PlusSignIcon} size={16} color="currentColor" strokeWidth={1.8} />
            </button>
            <span className="web-workbook-zoom-value" style={zoomTextStyle}>
              {zoom}%
            </span>
            <button
              aria-label="Zoom out"
              className="web-workbook-icon-button"
              onClick={() => setZoom((current) => clamp(current - 10, 50, 200))}
              style={iconButtonStyle}
              type="button"
            >
              <HugeiconsIcon icon={MinusSignIcon} size={16} color="currentColor" strokeWidth={1.8} />
            </button>
          </div>

          <button
            aria-label="Download workbook"
            className="web-workbook-icon-button"
            disabled={!workbook}
            onClick={handleDownload}
            style={iconButtonStyle}
            type="button"
          >
            <HugeiconsIcon
              icon={Download01Icon}
              size={17}
              color="currentColor"
              strokeWidth={1.8}
            />
          </button>
        </div>
      </header>
      <div className="web-workbook-formula-bar" aria-label="Formula bar">
        <div className="web-workbook-selection" title={selectionLabel}>
          {selectionLabel}
        </div>
        <div className="web-workbook-function-icon" aria-hidden="true">
          <HugeiconsIcon
            icon={FunctionOfXIcon}
            size={16}
            color="currentColor"
            strokeWidth={1.7}
          />
        </div>
        <div className="web-workbook-formula-value" title={formulaBarValue}>
          {formulaBarValue}
        </div>
      </div>
      <div className="web-workbook-body">
        {sheets.length > 0 ? (
          <div
            className="web-workbook-grid-shell"
            data-testid="web-workbook-grid"
            style={{ zoom: zoomScale }}
          >
            <DataEditor
              columns={gridData.columns}
              drawFocusRing={false}
              drawCell={drawCell}
              freezeColumns={1}
              getCellContent={gridData.getCellContent}
              getCellsForSelection
              gridSelection={gridSelection}
              headerHeight={32}
              height="100%"
              highlightRegions={highlightRegions}
              key={`${selectedSheet?.id}-${selectedSheet?.name}`}
              onCellClicked={handleCellClicked}
              onGridSelectionChange={handleGridSelectionChange}
              onVisibleRegionChanged={handleVisibleRegionChanged}
              rowHeight={(row) => gridData.rows[row]?.height || 32}
              rowMarkers="none"
              rows={gridData.rows.length}
              smoothScrollX
              smoothScrollY
              spanRangeBehavior="allowPartial"
              theme={gridTheme}
              width="100%"
            />
            <DrawingOverlay
              gridData={gridData}
              imageCache={imageCache}
              theme={gridTheme}
              visibleRegion={visibleRegion}
            />
          </div>
        ) : null}
      </div>
      <footer className="web-workbook-footer" aria-label="Workbook sheets">
        <div className="web-workbook-footer-label">Sheets</div>
        <div className="web-workbook-sheet-strip">
          {sheets.map((sheet, index) => (
            <button
              className={index === selectedSheetIndex ? "active" : ""}
              key={`${sheet.id || index}-${sheet.name || "sheet"}`}
              onClick={() => setSelectedSheetIndex(index)}
              title={sheet.name || `Sheet ${index + 1}`}
              type="button"
            >
              {sheet.name || `Sheet ${index + 1}`}
            </button>
          ))}
        </div>
      </footer>
    </section>
  );
}

function RedesignWorkspace({ workbook = null }) {
  const containerRef = useRef(null);
  const [leftWidth, setLeftWidth] = useState(50);
  const [isResizing, setIsResizing] = useState(false);

  const updateLeftWidth = useCallback((clientX) => {
    const rect = containerRef.current?.getBoundingClientRect();

    if (!rect?.width) {
      return;
    }

    setLeftWidth(clamp(((clientX - rect.left) / rect.width) * 100, 16, 84));
  }, []);

  useEffect(() => {
    if (!isResizing) {
      return undefined;
    }

    const handlePointerMove = (event) => {
      updateLeftWidth(event.clientX);
    };
    const handlePointerUp = () => {
      setIsResizing(false);
    };

    document.body.classList.add("is-resizing-redesign");
    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);

    return () => {
      document.body.classList.remove("is-resizing-redesign");
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
    };
  }, [isResizing, updateLeftWidth]);

  return (
    <section
      aria-label="Redesign workspace"
      className="redesign-page"
      ref={containerRef}
      style={{ "--left-pane-width": `${leftWidth}%` }}
    >
      <div className="redesign-pane redesign-pane-left">
        <WebWorkbook
          darkBgColor="#0F0F11"
          lightBgColor="#FFFFFF"
          workbook={workbook}
        />
      </div>
      <div
        aria-label="Resize spreadsheet areas"
        aria-orientation="vertical"
        aria-valuemax={84}
        aria-valuemin={16}
        aria-valuenow={Math.round(leftWidth)}
        className="redesign-divider"
        onKeyDown={(event) => {
          if (event.key === "ArrowLeft") {
            event.preventDefault();
            setLeftWidth((current) => clamp(current - 2, 16, 84));
          }

          if (event.key === "ArrowRight") {
            event.preventDefault();
            setLeftWidth((current) => clamp(current + 2, 16, 84));
          }
        }}
        onPointerDown={(event) => {
          event.preventDefault();
          setIsResizing(true);
          updateLeftWidth(event.clientX);
        }}
        role="separator"
        tabIndex={0}
      />
      <div className="redesign-pane" />
    </section>
  );
}

function WorkbookGrid({ workbook }) {
  const sheets = workbook?.workbook?.sheets || [];
  const workbookStyles = workbook?.workbook?.styles || workbook?.styles || {};
  const [selectedSheetIndex, setSelectedSheetIndex] = useState(0);
  const [selectedCell, setSelectedCell] = useState(null);
  const [imageCache, setImageCache] = useState({});
  const imageLoadRef = useRef(new Set());
  const selectedSheet = sheets[selectedSheetIndex] || sheets[0];
  const gridData = useMemo(
    () => buildGridData(selectedSheet, workbookStyles),
    [selectedSheet, workbookStyles],
  );
  const imageUrlsKey = useMemo(
    () =>
      gridData.drawingObjects
        .map(drawingImageUrl)
        .filter(Boolean)
        .sort()
        .join("|"),
    [gridData],
  );

  useEffect(() => {
    setSelectedCell(null);
  }, [selectedSheet]);

  useEffect(() => {
    const urls = imageUrlsKey ? imageUrlsKey.split("|").filter(Boolean) : [];

    for (const url of urls) {
      if (imageLoadRef.current.has(url)) {
        continue;
      }

      imageLoadRef.current.add(url);
      setImageCache((current) => ({
        ...current,
        [url]: {
          status: "loading",
        },
      }));

      const image = new Image();
      image.crossOrigin = "anonymous";
      image.onload = () => {
        setImageCache((current) => ({
          ...current,
          [url]: {
            image,
            status: "loaded",
          },
        }));
      };
      image.onerror = () => {
        setImageCache((current) => ({
          ...current,
          [url]: {
            status: "error",
          },
        }));
      };
      image.src = url;
    }
  }, [imageUrlsKey]);

  const handleCellClicked = useCallback(
    (item) => {
      const [col, row] = item;

      if (col === 0) {
        setSelectedCell(null);
        return;
      }

      const details = gridData.getCellMeta(col, row);

      setSelectedCell(
        details
          ? {
              address: details.address,
              drawings: details.drawings,
              formula: details.formula,
              merge: details.merge,
              sourceAddress: details.sourceAddress,
              style: details.style,
              styleIndex: details.styleIndex,
              table: details.table,
              type: details.type,
              value: details.value,
            }
          : null,
      );
    },
    [gridData],
  );

  const drawCell = useCallback(
    (args, drawContent) => {
      const details = gridData.getCellMeta(args.col, args.row);

      if (details?.drawings?.length > 0) {
        drawDrawingOverlays(args.ctx, args.rect, details.drawings, args.theme, imageCache);
        return;
      }

      drawContent();

      if (!details || details.isRowHeader) {
        return;
      }

      drawTextDecoration(args.ctx, args.rect, details, args.theme);

      const border = details.presentation.border;
      args.ctx.save();
      drawBorderLine(args.ctx, args.rect, "top", border.top);
      drawBorderLine(args.ctx, args.rect, "right", border.right);
      drawBorderLine(args.ctx, args.rect, "bottom", border.bottom);
      drawBorderLine(args.ctx, args.rect, "left", border.left);
      args.ctx.restore();

      drawTableOutline(args.ctx, args.rect, details.table);
      drawDrawingOverlays(args.ctx, args.rect, details.drawings, args.theme, imageCache);
    },
    [gridData, imageCache],
  );

  if (sheets.length === 0) {
    return <div className="empty-state">No sheets found in this workbook.</div>;
  }

  return (
    <>
      <div className="sheet-tabs" aria-label="Workbook sheets">
        {sheets.map((sheet, index) => (
          <button
            className={index === selectedSheetIndex ? "active" : ""}
            key={`${sheet.id}-${sheet.name}`}
            onClick={() => setSelectedSheetIndex(index)}
            type="button"
          >
            {sheet.name || `Sheet ${index + 1}`}
          </button>
        ))}
      </div>

      <FormulaBar gridData={gridData} selectedCell={selectedCell} />
      <TableStrip tables={gridData.tables} />
      <DrawingStrip drawings={gridData.drawingObjects} />

      <div className="grid-box" data-testid="grid-box">
        <DataEditor
          columns={gridData.columns}
          drawCell={drawCell}
          freezeColumns={1}
          getCellContent={gridData.getCellContent}
          getCellsForSelection
          headerHeight={32}
          height="100%"
          key={`${selectedSheet?.id}-${selectedSheet?.name}`}
          onCellClicked={handleCellClicked}
          rowHeight={(row) => gridData.rows[row]?.height || 32}
          rowMarkers="none"
          rows={gridData.rows.length}
          smoothScrollX
          smoothScrollY
          spanRangeBehavior="allowPartial"
          theme={defaultGridTheme}
          width="100%"
        />
      </div>
    </>
  );
}

function App() {
  const isRedesignPage = isRedesignPath();
  const [activeNumber] = useState(getRouteNumber);
  const [redesignNumber] = useState(getRedesignRouteNumber);
  const [workbook, setWorkbook] = useState(null);
  const [redesignWorkbook, setRedesignWorkbook] = useState(null);
  const [status, setStatus] = useState(`/${activeNumber}`);
  const [error, setError] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isRedesignLoading, setIsRedesignLoading] = useState(false);
  const [themeName, setThemeName] = useState(getInitialTheme);

  if (window.location.pathname === "/") {
    window.history.replaceState(null, "", `/${activeNumber}`);
  }

  useEffect(() => {
    window.localStorage.setItem("openxml-viewer-theme", themeName);
  }, [themeName]);

  async function handleFetch() {
    setIsLoading(true);
    setError("");
    setStatus(`/${activeNumber} loading`);
    setWorkbook(null);

    try {
      const response = await fetch(`${API_BASE_URL}/xlsx/${activeNumber}`, {
        method: "POST",
        headers: {
          Accept: "application/json",
        },
      });
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload?.error || `Request failed with ${response.status}`);
      }

      setWorkbook(payload);
      setStatus(`/${activeNumber} loaded`);
    } catch (fetchError) {
      setError(fetchError.message);
      setStatus(`/${activeNumber} error`);
    } finally {
      setIsLoading(false);
    }
  }

  async function handleRedesignFetch() {
    setIsRedesignLoading(true);
    setRedesignWorkbook(null);

    try {
      const response = await fetch(`${API_BASE_URL}/xlsx/${redesignNumber}`, {
        method: "POST",
        headers: {
          Accept: "application/json",
        },
      });
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload?.error || `Request failed with ${response.status}`);
      }

      setRedesignWorkbook(payload);
    } catch (fetchError) {
      console.error(fetchError);
    } finally {
      setIsRedesignLoading(false);
    }
  }

  return (
    <section className={`shell theme-${themeName}`}>
      <header className="topbar" aria-label="Spreadsheet routes">
        <nav className="routes">
          {fileNumbers.map((number) => (
            <a
              aria-label={`Open file ${number}`}
              className={
                isRedesignPage
                  ? number === redesignNumber
                    ? "active"
                    : ""
                  : number === activeNumber
                    ? "active"
                    : ""
              }
              href={isRedesignPage ? `/redesign/${number}` : `/${number}`}
              key={number}
            >
              {number}
            </a>
          ))}
          <a
            aria-label="Open redesign page"
            className={isRedesignPage ? "active redesign-link" : "redesign-link"}
            href="/redesign"
          >
            Redesign
          </a>
        </nav>

        {isRedesignPage ? (
          <button
            className="fetch-button"
            data-testid="fetch-button"
            disabled={isRedesignLoading}
            onClick={handleRedesignFetch}
            type="button"
          >
            {isRedesignLoading ? "Fetching" : "Fetch"}
          </button>
        ) : (
          <button
            className="fetch-button"
            data-testid="fetch-button"
            disabled={isLoading}
            onClick={handleFetch}
            type="button"
          >
            {isLoading ? "Fetching" : "Fetch"}
          </button>
        )}

        <button
          aria-label={`Switch to ${themeName === "dark" ? "light" : "dark"} theme`}
          className="theme-toggle"
          onClick={() => setThemeName((current) => (current === "dark" ? "light" : "dark"))}
          type="button"
        >
          {themeName === "dark" ? "Light" : "Dark"}
        </button>
      </header>

      {isRedesignPage ? (
        <RedesignWorkspace workbook={redesignWorkbook} />
      ) : (
        <section className="workspace" aria-live="polite">
          <div className="status" data-testid="status">
            {status}
          </div>

          {error ? (
            <div className="grid-box error-box" data-testid="grid-box">
              {error}
            </div>
          ) : workbook ? (
            <WorkbookGrid workbook={workbook} />
          ) : (
            <div className="grid-box empty-state" data-testid="grid-box">
              Click Fetch to load the workbook grid.
            </div>
          )}
        </section>
      )}
    </section>
  );
}

createRoot(document.querySelector("#app")).render(<App />);
