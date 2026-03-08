/** biome-ignore-all assist/source/useSortedKeys: <explanatiaon> */
/** biome-ignore-all lint/complexity/noStaticOnlyClass: <explanatiaon> */
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
          start: { col: "J", row: 6 },
          finish: { row: 17 },
        },
      ],
    },
  ],
  {
    delimiter: '"+"', // Fix delimiter string formatting for Excel
    formatCells: (cell1) => cell1,
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
