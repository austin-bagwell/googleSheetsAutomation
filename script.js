const odDataSheet = SpreadsheetApp.getActive().getSheetByName("ODFL Data");
const nsDataSheet = SpreadsheetApp.getActive().getSheetByName("NS Data");
const datasetSheet = SpreadsheetApp.getActive().getSheetByName("Dataset");
const dcInfoSheet = SpreadsheetApp.getActive().getSheetByName("DC Info");

const datasetData = SpreadsheetApp.getActive()
  .getSheetByName("Dataset")
  .getDataRange()
  .getValues();
const datasetHeaders = datasetData[0];

// TODO - so much error handling to add good god

function main() {
  // GET AND PARSE CSV REPORTS FROM GMAIL
  function getCSVFromGmail(label) {
    const gmailThread = GmailApp.search(`label:${label}`, 0, 1)[0];
    const attachments = gmailThread.getMessages()[0].getAttachments();
    const attachmentNames = [];
    attachments.forEach((attachment) =>
      attachmentNames.push(attachment.getName())
    );
    const csv = attachmentNames.indexOf(
      attachmentNames.find((attach) => attach.toLowerCase().endsWith(".csv"))
    );
    const parsedCSV = Utilities.parseCsv(attachments[csv].getDataAsString());
    return parsedCSV;
  }

  // DO INITIAL DATA CLEANUP ON REPORTS -aka remove # from OD PO nums
  function cleanODReport(arr) {
    const poNumber = arr[0].indexOf("Purchase Order Number");
    const body = arr.slice(1);
    body.forEach((row) => {
      row[poNumber] = row[poNumber].replace("#", "");
    });
    return arr;
  }

  // NORMALIZE REPORT HEADERS
  function replaceSpaces(arr) {
    return arr.map((header) => header.replaceAll(" ", "_"));
  }

  function normalizeNetsuiteHeaders(arr) {
    const headers = arr.slice(0, 1).flat();

    const location = headers.indexOf("Location");
    const fillDate = headers.indexOf("Fulfillment Date");
    const nsMemo = headers.indexOf("Memo");
    const name = headers.indexOf("Name");
    const po = headers.indexOf("Customer PO #");
    headers[location] = "Ship From";
    headers[fillDate] = "Scheduled Ship Date";
    headers[nsMemo] = "Due Date";
    headers[name] = "DC Name";
    headers[po] = "PO #";

    const cleanHeaders = replaceSpaces(headers);
    arr[0] = cleanHeaders;
    return arr;
  }

  function normalizeODHeaders(arr) {
    const headers = arr.slice(0, 1).flat();
    const body = arr.slice(0, 1);

    const poNum = headers.indexOf("Purchase Order Number");
    const shipDate = headers.indexOf("Actual Pickup Date");
    const pro = headers.indexOf("OD Pro#");
    const arriveAtYard = headers.indexOf("Arrival Date");
    const delivery = headers.indexOf("Delivery Date");
    const pallets = headers.indexOf("Pieces (skids/pallets)");

    headers[poNum] = "PO #";
    headers[shipDate] = "Actual Ship Date";
    headers[pro] = "PRO Number";
    headers[arriveAtYard] = "Arrived At Carrier Yard";
    headers[delivery] = "Delivery Date";
    headers[pallets] = "Pallet Count";

    const cleanHeaders = replaceSpaces(headers);
    arr[0] = cleanHeaders;
    return arr;
  }

  function normalizeEstesHeaders(arr) {}

  // CONVERT NETSUITE REPORT INTO ARRAY OF OBJECTS
  // CONVERT NORMALIZED SHIPMET DETAILS INTO ARRAY OF OBJECTS

  // TODO
  // modify func so that it accepts N number of arrays
  // that way, I can pass it it multiple reports at once
  // and the normalized data in those reports will be
  // all pushed into the same array at once
  // Ex: function two2Arr(targetHeaders, ...arrays) {};

  function convert2DArrayToArrayOfObjects(targetHeaders, arr) {
    const objects = [];
    const headers = arr.slice(0, 1).flat();
    const body = arr.slice(1);

    for (let row = 0; row < body.length; row++) {
      const shipmentDetail = {};
      for (let col = 0; col < headers.length; col++) {
        const key = headers[col];
        const val = body[row][col];

        if (targetHeaders.includes(key)) {
          shipmentDetail[key] = val;
        } else {
          continue;
        }
      }
      objects.push(shipmentDetail);
    }
    return objects;
  }

  // EXAMPLE WORKFLOW
  // getting, cleaning, normalizing CSV, distilling all relevant
  // shipment details into an array of objects
  const gmailLabels = [
    "estes-ship-report",
    "odfl-ship-report",
    "ns-ltl-report",
  ];
  const shipmentDetailHeaders = replaceSpaces([
    "Actual Ship Date",
    "Arrived At Carrier Yard",
    "PO #",
    "Delivery Date",
    "PRO Number",
    "Pallet Count",
    "Weight",
  ]);
  const oldDominionReport = getCSVFromGmail(gmailLabels[1]);
  // TODO add data cleanup to normalization function?
  const cleanOD = cleanODReport(oldDominionReport);
  const normalOD = normalizeODHeaders(cleanOD);
  const netsuiteReport = getCSVFromGmail(gmailLabels[2]);
  const normalNetsuite = normalizeNetsuiteHeaders(netsuiteReport);
  const nsHeaders = replaceSpaces(
    SpreadsheetApp.getActive()
      .getSheetByName("Dataset")
      .getDataRange()
      .getValues()
      .slice(0, 1)
      .flat()
  );
  const odShipmentDetails = convert2DArrayToArrayOfObjects(
    shipmentDetailHeaders,
    normalOD
  );
  const netsuiteObjects = convert2DArrayToArrayOfObjects(
    nsHeaders,
    normalNetsuite
  );
  // Logger.log(odShipmentDetails[0]);
  // Logger.log(netsuiteObjects);

  // CONVERT ACTIVE SHIPMENT RANGE ROWS INTO ARRAY OF OBJECTS
  // I want everything to be a string, but I'm getting different
  // datatypes from datasetRange

  function getOldestActiveShipmentIndex() {
    const deliveryDate = datasetHeaders.indexOf("Delivery Date");
    let indexOfOldestActiveShipment;
    for (let i = 0; i < datasetData.length; i++) {
      if (datasetData[i][deliveryDate] === "") {
        indexOfOldestActiveShipment = i;
        break;
      }
    }
    return indexOfOldestActiveShipment;
  }

  // const startOfActiveShipmentRange = getOldestActiveShipmentIndex();
  const shipmentHeaders = replaceSpaces(datasetData[0].slice(0, 13));
  const activeShipmentsArray = datasetData.slice(
    getOldestActiveShipmentIndex()
  );
  // this removes the 'right hand formula' columns from activeShipmentsArray
  for (row of activeShipmentsArray) {
    row.splice(shipmentHeaders.length);
  }
  // this adds a header row which is needed for the fancy formula
  activeShipmentsArray.unshift(shipmentHeaders);

  // Logger.log(activeShipmentsArray[0]);
  // Logger.log(activeShipmentsArray[1]);

  const activeShipmentObjects = convert2DArrayToArrayOfObjects(
    shipmentHeaders,
    activeShipmentsArray
  );
  Logger.log(activeShipmentObjects[0]);

  // PUSH NEW NETSUITE ORDER OBJECTS INTO ACTIVE SHIPMENTS ARRAY
  // UPDATE ACTIVE SHIPMENT OBJECTS WITH NEW SHIPPING DETAILS

  // GET ACTIVE SHIPMENTS AND CONVERT TO SHIPMENT OBJECTS

  // STEP 4 - UPDATE ACTIVESHIPMENTRANGE SHIPMENT OBJECTS WITH NORMALIZED SHIPMENT DATA
  // -- if dataset[row][colToUpdate] === '', add the normalData[colToUpdate] value (could sitll be "", just means no new data)
  // STEP 5 - PUSH UPDATED DATASET BACK TO SPREADSHEET
  // STEP 6 - AUTOFILL 'RIGHT HAND' SHEETS FORMULAS
  // -- These formulas feed the reporting dashboard stuff

  // ACTUALLY CALL/USE ALL FUNCTIONS DEFINED ABOVE IN THIS AREA
  // const gmailLabels = ['estes-ship-report','odfl-ship-report','ns-ltl-report'];
  const estesReport = getCSVFromGmail(gmailLabels[0]);
  // const oldDominionReport = getCSVFromGmail(gmailLabels[1])
  // const netsuiteReport = getCSVFromGmail(gmailLabels[2]);
}

// autofills the 8 columns of sheets formulas on the right hand side of the dataset sheet.
// # of columns is hardcoded at the moment, if formulas are added this will need to be updated
function fillRightHandFormulas() {
  const firstRow = datasetSheet.getRange(["N2:U2"]);
  const formulaRows = datasetSheet.getRange(
    2,
    14,
    datasetSheet.getLastRow() - 2,
    8
  );
  firstRow.autoFill(formulaRows, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

// Get rid of formulas from Netsuite and Shipping details section while keeping values
// ie. locks in the data for first 11 columns without any pesky formulas sticking around
// will be the last function called
function pasteValsOnlyEquiv() {
  const rngCopyValsOnly = datasetSheet
    .getRange(2, 1, datasetSheet.getLastRow(), 13)
    .getValues();
  datasetSheet
    .getRange(2, 1, datasetSheet.getLastRow(), 13)
    .setValues(rngCopyValsOnly);
}
