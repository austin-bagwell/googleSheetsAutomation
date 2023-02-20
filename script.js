const odDataSheet = SpreadsheetApp.getActive().getSheetByName("ODFL Data");
const nsDataSheet = SpreadsheetApp.getActive().getSheetByName("NS Data");
const datasetSheet = SpreadsheetApp.getActive().getSheetByName("Dataset");
const dcInfoSheet = SpreadsheetApp.getActive().getSheetByName("DC Info");

const datasetData = SpreadsheetApp.getActive()
  .getSheetByName("Dataset")
  .getDataRange()
  .getDisplayValues();
const datasetHeaders = datasetData[0];

// TODO
// add normalization for Estes data
// combine OD/Estes data into a 'normalizedShippingDetails' array for which to update shipments with
// add a ton of error handling/logging

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
  function cleanOldDominionReport(arr) {
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

  // maybe wrap header normalization && general cleanup here?
  // might be better to return a smaller initial array
  function normalizeEstesHeaders(arr) {
    const headers = arr.slice(0, 1).flat();
    const body = arr.slice(0, 1);

    const poNum = headers.indexOf("Purchase Order #");
    const shipDate = headers.indexOf("Pickup Date");
    const pro = headers.indexOf("Pro #");
    const arriveAtYard = headers.indexOf("Arrival Date");
    const delivery = headers.indexOf("Delivery Date");
    const pallets = headers.indexOf("Pallets");
    const weight = headers.indexOf("Weight*");

    const columnsIWant = [
      pro,
      poNum,
      shipDate,
      arriveAtYard,
      delivery,
      pallets,
      weight,
    ];

    const shorterArr = [];
    for (const row of arr) {
      const newRow = [];
      const iterator = columnsIWant.values();
      for (const i of iterator) {
        const val = row[i];
        newRow.push(val);
      }
      shorterArr.push(newRow);
    }

    // i'm sorry for this and will clean it later
    const headersToKeep = shorterArr[0];
    headersToKeep[0] = "PRO Number";
    headersToKeep[1] = "PO #";
    headersToKeep[2] = "Actual Ship Date";
    headersToKeep[3] = "Arrived At Carrier Yard";
    headersToKeep[4] = "Delivery Date";
    headersToKeep[5] = "Pallet Count";
    headersToKeep[6] = "Weight";

    const cleanHeaders = replaceSpaces(headersToKeep);
    shorterArr[0] = cleanHeaders;
    return shorterArr;
  }

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

  // CONVERT ACTIVE SHIPMENT RANGE ROWS INTO ARRAY OF OBJECTS
  function getOldestActiveShipmentIndex() {
    const deliveryDate = datasetHeaders.indexOf("Delivery Date");
    let indexOfOldestActiveShipment;
    for (const [i, row] of datasetData.entries()) {
      if (row[deliveryDate] === "") {
        indexOfOldestActiveShipment = i;
        break;
      }
    }
    return indexOfOldestActiveShipment;
  }
  // 'PO_#' is the common key between target and source data
  function updateShipmentsWithNewData(target, source, key) {
    for (let existing of target) {
      for (let updated of source) {
        if (existing[key] === updated[key]) {
          Object.assign(existing, updated);
          break;
        }
      }
    }
    return target;
  }

  function makeUpdatedShipmentsArray(shipmentObs, headers) {
    const newRows = [];

    for (const shipment of shipmentObs) {
      const row = [];
      for (const header of headers) {
        // Logger.log(shipment[header])
        row.push(shipment[header]);
      }
      newRows.push(row);
    }

    return newRows;
  }

  // ACTUALLY CALL/USE ALL FUNCTIONS DEFINED ABOVE
  const estesReport = getCSVFromGmail("estes-ship-report");
  // Logger.log(estesReport);
  const oldDominionReport = getCSVFromGmail("odfl-ship-report");
  const netsuiteReport = getCSVFromGmail("ns-ltl-report");

  const cleanedOD = cleanOldDominionReport(oldDominionReport);
  const normalizedOldDominionReport = normalizeODHeaders(cleanedOD);
  const normalizedNetsuiteReport = normalizeNetsuiteHeaders(netsuiteReport);
  const normalizedEstesReport = normalizeEstesHeaders(estesReport);

  const netsuiteInfoHeaders = replaceSpaces(datasetHeaders.slice(0, 7));
  const carrierInfoHeaders = replaceSpaces(datasetHeaders.slice(7, 13));
  // must add this header back in to have a lookup key ... UGLY
  carrierInfoHeaders.push("PO_#");
  // fullShipmentHeaders? this defines headers for all the cols in Dataset I'm updating
  const shipmentHeaders = replaceSpaces(datasetHeaders.slice(0, 13));

  const activeShipmentsArray = datasetData.slice(
    getOldestActiveShipmentIndex()
  );
  // removes the 'right hand formula' columns from activeShipmentsArray
  for (row of activeShipmentsArray) {
    row.splice(shipmentHeaders.length);
  }
  // adds a header row which is needed for convert2DArrayToArrayOfObjects()
  activeShipmentsArray.unshift(shipmentHeaders);

  const odShipmentDetails = convert2DArrayToArrayOfObjects(
    carrierInfoHeaders,
    normalizedOldDominionReport
  );
  const estesShipmentDetails = convert2DArrayToArrayOfObjects(
    carrierInfoHeaders,
    normalizedEstesReport
  );
  // make this more programatic at some point
  const allNewShipmentDetailObjects =
    odShipmentDetails.concat(estesShipmentDetails);

  const existingActiveShipments = convert2DArrayToArrayOfObjects(
    shipmentHeaders,
    activeShipmentsArray
  );
  const newNetsuiteOrders = convert2DArrayToArrayOfObjects(
    netsuiteInfoHeaders,
    normalizedNetsuiteReport
  );

  const activeShipmentsObjectArray =
    existingActiveShipments.concat(newNetsuiteOrders);

  // TODO combine OD/estes shipping stuff into one normalizedShippingDetails array/object array
  const updatedShipments = updateShipmentsWithNewData(
    activeShipmentsObjectArray,
    allNewShipmentDetailObjects,
    "PO_#"
  );

  const updatedShipmentsArray = makeUpdatedShipmentsArray(
    updatedShipments,
    shipmentHeaders
  );

  const oldestActiveShipment = getOldestActiveShipmentIndex() + 1;
  // IT WORKS!
  // commenting out the final value reset so I don't keep updating shit
  // I still need to add Estes handling but whooooooo
  const rangeOfShipmentsToUpdate = datasetSheet.getRange(
    oldestActiveShipment,
    1,
    updatedShipmentsArray.length,
    shipmentHeaders.length
  );
  rangeOfShipmentsToUpdate.setValues(updatedShipmentsArray);
}
