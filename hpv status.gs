function onEdit(e) {
  if (!e) return; // Ignore manual runs

  const sheetName = "DATA";
  const rawSheetName = "RAW";
  const idCol = 1;        // Student_Naitonal ID - Column A
  const dobCol = 3;       // DOB (DD/MM/YYYY) - Column C
  const outputCol = 7;    // Vaccination Status - Column G

  const sheet = e.source.getSheetByName(sheetName);
  if (!sheet || sheet.getName() !== sheetName) return;

  const range = e.range;
  if (range.getRow() < 2) return; // Skip header

  const row = range.getRow();
  updateRowStatus(sheet, row, idCol, dobCol, outputCol, rawSheetName, e.source);
}

function updateAllVaccinationStatus() {
  const sheetName = "DATA";
  const rawSheetName = "RAW";
  const idCol = 1;        // Student_Naitonal ID
  const dobCol = 3;       // DOB
  const outputCol = 7;    // Vaccination Status

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    updateRowStatus(sheet, row, idCol, dobCol, outputCol, rawSheetName, ss);
  }
}

// Helper function to update status for a single row
function updateRowStatus(sheet, row, idCol, dobCol, outputCol, rawSheetName, ss) {
  const id = sheet.getRange(row, idCol).getValue();
  const dob = sheet.getRange(row, dobCol).getValue();

  if (!id || !dob || isNaN(new Date(dob))) {
    sheet.getRange(row, outputCol).setValue("");
    return;
  }

  const today = new Date();
  const dobDate = new Date(dob);
  let years = today.getFullYear() - dobDate.getFullYear();
  let months = today.getMonth() - dobDate.getMonth();
  if (months < 0 || (months === 0 && today.getDate() < dobDate.getDate())) {
    years--;
    months += 12;
  }

  const isUnder10 = years < 10;
  const isBetween10And10Half = (years === 10 && months < 6);

  const rawSheet = ss.getSheetByName(rawSheetName);
  const rawData = rawSheet.getRange("C2:N" + rawSheet.getLastRow()).getValues();
  const match = rawData.find(rowData => rowData[0] === id); // RAW!C is column 0 in rawData

  let status = "";

  if (isUnder10) {
    status = "Not eligible";
  } else if (isBetween10And10Half) {
    status = match ? (match[11] ? "Vaccinated" : "Pending Vaccination") : "Pending Vaccination";
  } else {
    status = match ? (match[11] ? "Vaccinated" : "Pending Vaccination") : "No Records Found";
  }

  sheet.getRange(row, outputCol).setValue(status);
}
