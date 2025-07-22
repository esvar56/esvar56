function updateOrInsertRowsWithCompositeKey() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Source");
  const destSheet = ss.getSheetByName("Destination");

  const sourceData = sourceSheet.getDataRange().getValues();
  const destData = destSheet.getDataRange().getValues();

  const keyColumns = [0, 1, 2]; // Use columns A, B, C as composite key (0-based index)

  const destMap = new Map();

  // Build key map from destination sheet
  destData.forEach((row, index) => {
    const key = keyColumns.map(i => row[i]).join('|');
    destMap.set(key, { index: index, row: row });
  });

  sourceData.forEach((srcRow) => {
    const key = keyColumns.map(i => srcRow[i]).join('|');
    const match = destMap.get(key);

    if (match) {
      const destRow = match.row;
      const isDifferent = JSON.stringify(destRow) !== JSON.stringify(srcRow);

      if (isDifferent) {
        // Update the row if contents are different
        destSheet.getRange(match.index + 1, 1, 1, srcRow.length).setValues([srcRow]);
      }
    } else {
      // Append the new row
      destSheet.appendRow(srcRow);
    }
  });

  Logger.log("Update or insert completed.");
}
