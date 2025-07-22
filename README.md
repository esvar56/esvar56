function updateOrInsertRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Source");
  const destSheet = ss.getSheetByName("Destination");

  const sourceData = sourceSheet.getDataRange().getValues();
  const destData = destSheet.getDataRange().getValues();

  const destMap = new Map();

  // Build a map for quick access from destination (use JSON string of row as key if no unique ID)
  destData.forEach((row, index) => {
    const key = row.join("|||"); // Change this if you want a unique field instead
    destMap.set(key, index);
  });

  sourceData.forEach((srcRow) => {
    const srcKey = srcRow.join("|||");

    let found = false;

    // Check for similar row in dest by matching partial key (you can optimize this)
    for (let i = 0; i < destData.length; i++) {
      const destRow = destData[i];

      // If the row matches entirely, skip
      if (JSON.stringify(destRow) === JSON.stringify(srcRow)) {
        found = true;
        break;
      }

      // If the row partially matches (same key column), update
      if (destRow[0] === srcRow[0]) {
        destSheet.getRange(i + 1, 1, 1, srcRow.length).setValues([srcRow]);
        found = true;
        break;
      }
    }

    // If not found, append to the bottom
    if (!found) {
      destSheet.appendRow(srcRow);
    }
  });

  Logger.log("Sync complete.");
}
