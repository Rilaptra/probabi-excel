/** biome-ignore-all assist/source/useSortedKeys: <explanation> */
/** biome-ignore-all lint/complexity/noStaticOnlyClass: <explanation> */
import { execSync } from "node:child_process";
import { ExcelToolkit } from "./utils/excel";

// ==========================================
// 🚀 EXECUTION
// ==========================================

const superFormula = ExcelToolkit.generateSuperFormula(
  [
    {
      mappingMode: "parallel", // <-- THIS IS THE MAGIC ⚡
      ranges: [
        {
          start: { col: "M", row: 34 },
          finish: { row: 37 },
        },
        {
          start: { col: "N", row: 34 },
          finish: { row: 37 },
        },
      ],
    },
  ],
  {
    delimiter: '" + "', // Fix delimiter string formatting for Excel
    formatCells: (cell1, cell2) => `FIXED(${cell1};4)&" ⋅ "&FIXED(${cell2};4)`,
    wrapper: {
      start: "",
      end: "",
    },
  },
);

console.log("\n=== 🎯 GENERATED EXCEL FORMULA ===");
console.log(superFormula);

try {
  execSync("clip", { input: superFormula });
  console.log(
    "\n📋 [SUCCESS] Formula automatically copied to clipboard! (Ready to paste 🚀)",
  );
} catch (error) {
  console.error("\n❌ [ERROR] Failed to auto-copy to clipboard.", error);
}
