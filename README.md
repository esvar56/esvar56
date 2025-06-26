function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  // Define the range to search in (e.g., entire sheet)
  const range = sheet.getUsedRange();

  // Text to search for
  const searchText = "Invoice";

  // Perform the find operation (first match only)
  const foundCell = range.find(searchText, {
    completeMatch: false,  // Use true for exact match
    matchCase: false       // Use true to match case
  });

  if (foundCell) {
    console.log(`Found "${searchText}" at cell: ${foundCell.getAddress()}`);
    console.log(`Value in cell: ${foundCell.getValue()}`);
  } else {
    console.log(`"${searchText}" not found.`);
  }
}
