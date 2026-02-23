import { execSync } from "node:child_process";

export interface CellCoordinate {
  col: string;
  row: number;
}

export interface RangeConfig {
  start: CellCoordinate;
  finish?: Partial<CellCoordinate>;
}

export interface FormulaGroup {
  mappingMode?: "cartesian" | "parallel";
  ranges: RangeConfig[];
}

export interface ExcelTemplate {
  wrapper: {
    start?: string;
    end?: string;
  };
  delimiter: string;
  formatCells: (...cells: string[]) => string;
}

export class ExcelToolkit {
  private static colToInt(col: string): number {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.toUpperCase().charCodeAt(i) - 64);
    }
    return num;
  }

  private static intToCol(num: number): string {
    let col = "";
    let n = num;
    while (n > 0) {
      const rem = (n - 1) % 26;
      col = String.fromCharCode(65 + rem) + col;
      n = Math.floor((n - 1) / 26);
    }
    return col;
  }

  private static expandRange(range: RangeConfig): string[] {
    const startRow = range.start.row;
    const endRow = range.finish?.row ?? startRow;

    const startCol = ExcelToolkit.colToInt(range.start.col);
    const endCol = ExcelToolkit.colToInt(range.finish?.col ?? range.start.col);

    if (startRow > endRow || startCol > endCol) {
      throw new Error(
        `[DX Error] Invalid range: start > finish at ${range.start.col}${startRow}`,
      );
    }

    const cells: string[] = [];
    for (let c = startCol; c <= endCol; c++) {
      const colStr = ExcelToolkit.intToCol(c);
      for (let r = startRow; r <= endRow; r++) {
        cells.push(`${colStr}${r}`);
      }
    }
    return cells;
  }

  private static cartesianProduct(arrays: string[][]): string[][] {
    if (arrays.length === 0) return [];
    return arrays.reduce<string[][]>(
      (acc, curr) => acc.flatMap((a) => curr.map((c) => [...a, c])),
      [[]],
    );
  }

  public static generateSuperFormula(
    groups: FormulaGroup[],
    template: ExcelTemplate,
  ): string {
    if (!groups.length)
      throw new Error("[DX Error] groups array cannot be empty!");

    const allPairs: string[] = [];

    for (const group of groups) {
      if (!group.ranges.length) continue;

      const expandedRanges = group.ranges.map((r) =>
        ExcelToolkit.expandRange(r),
      );
      const mode = group.mappingMode ?? "cartesian";

      let combinations: string[][] = [];

      if (mode === "cartesian") {
        combinations = ExcelToolkit.cartesianProduct(expandedRanges);
      } else if (mode === "parallel") {
        const minLength = Math.min(...expandedRanges.map((arr) => arr.length));
        for (let i = 0; i < minLength; i++) {
          combinations.push(expandedRanges.map((arr) => arr[i] ?? "undefined"));
        }
      }

      for (const combo of combinations) {
        allPairs.push(template.formatCells(...combo));
      }
    }

    if (!allPairs.length) return "=";

    const parts: string[] = [];
    if (template.wrapper.start) parts.push(template.wrapper.start);

    parts.push(allPairs.join(` & ${template.delimiter} & `));

    if (template.wrapper.end) parts.push(template.wrapper.end);

    return `=${parts.join(" & ")}`;
  }
}
