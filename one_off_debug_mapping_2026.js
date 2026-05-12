/**
 * ****************************************
 * Google Apps Script - One-Off Mapping Debug
 * File: one_off_debug_mapping_2026.js
 * Description: Simulates the column mapping logic to debug index jumps.
 * ****************************************
 */

function debugMapping() {
  const productColMap = {};
  let currentDeepDiveColIdx = 5; // A=1(Profile), B=2(SubRegion), C=3(JobTitle), D=4(Tier1 Count)

  PRODUCT_SCHEMA.forEach(group => {
    Logger.log(`Processing Solution: ${group.solution}, Start Index: ${currentDeepDiveColIdx}`);
    currentDeepDiveColIdx++; // Skip Spacer
    group.products.forEach(prod => {
      const colLetter = columnToLetter(currentDeepDiveColIdx);
      productColMap[prod.toLowerCase()] = colLetter;
      Logger.log(`  Mapped Product: ${prod} -> ${colLetter} (${currentDeepDiveColIdx})`);
      currentDeepDiveColIdx++;
    });
  });
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
