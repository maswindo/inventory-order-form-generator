// Order Form Generator
// Updates inventory levels + generates company-wide stock levels

function updateInventoryLevels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const databaseSheet = ss.getSheetByName("Database");
  const locationSheet = ss.getSheetByName("Location-Based Inventory");

  const dbData = databaseSheet.getDataRange().getValues();
  const locData = locationSheet.getDataRange().getValues();
  
  const dbHeaders = dbData[0];
  const locHeaders = locData[0];

  const dbMaterialIndex = dbHeaders.indexOf("Material");
  const dbLocationIndex = dbHeaders.indexOf("Name");
  const dbQuantityIndex = dbHeaders.indexOf("Qty");
  const dbUOMIndex = dbHeaders.indexOf("UOM");

  const locMaterialIndex = locHeaders.indexOf("Material Name");
  const locLocationIndex = locHeaders.indexOf("Location");
  const locQuantityIndex = locHeaders.indexOf("Quantity");
  const locUOMIndex = locHeaders.indexOf("Unit");

  let dbStockMap = {};
  for (let i = 1; i < dbData.length; i++) {
    let material = dbData[i][dbMaterialIndex];
    let location = dbData[i][dbLocationIndex];
    let quantity = dbData[i][dbQuantityIndex];
    let uom = dbData[i][dbUOMIndex];
    dbStockMap[`${location}|${material}`] = { quantity, uom };
  }

  for (let i = 1; i < locData.length; i++) {
    let location = locData[i][locLocationIndex];
    let material = locData[i][locMaterialIndex];
    let oldQuantity = locData[i][locQuantityIndex];
    let key = `${location}|${material}`;

    if (dbStockMap[key]) {
      locationSheet.getRange(i + 1, locQuantityIndex + 1).setValue(dbStockMap[key].quantity);
      locationSheet.getRange(i + 1, locUOMIndex + 1).setValue(dbStockMap[key].uom);
    }
  }
}

function calculateTotalStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const locationSheet = ss.getSheetByName("Location-Based Inventory");
  const companySheet = ss.getSheetByName("Company-Wide Inventory");

  const locData = locationSheet.getDataRange().getValues();
  const locHeaders = locData[0];

  const locMaterialIndex = locHeaders.indexOf("Material Name");
  const locQuantityIndex = locHeaders.indexOf("Quantity");
  const locUOMIndex = locHeaders.indexOf("Unit");

  let companyStockMap = {};
  let uomMap = {};

  for (let i = 1; i < locData.length; i++) {
    let material = locData[i][locMaterialIndex];
    let quantity = locData[i][locQuantityIndex];
    let uom = locData[i][locUOMIndex];

    if (!companyStockMap[material]) {
      companyStockMap[material] = 0;
      uomMap[material] = uom;
    }
    companyStockMap[material] += quantity;
  }

  companySheet.getDataRange().offset(1, 0).clearContent();
  let outputData = Object.entries(companyStockMap).map(([material, totalStock]) => [material, totalStock, uomMap[material] || "N/A"]);
  companySheet.getRange(2, 1, outputData.length, 3).setValues(outputData);
}
