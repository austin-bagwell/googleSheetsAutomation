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

  class NormalizationConfig {
    constructor(rawHeaders = [], nrmlHeaders = []) {
      this.rawHeaders = rawHeaders;
      this.headerMap = new Map(
        rawHeaders.map((val, i) => [val, nrmlHeaders[i]])
      );
    }
  }
  function normalizeCSVReport(report = [[]], config = {}, cleanDataCallback) {
    const { rawHeaders, headerMap } = config;

    const getRawIndexes = () => {
      const reportHeaders = report[0];
      const indexes = [];

      for (const [i, header] of reportHeaders.entries()) {
        if (rawHeaders.includes(header)) {
          indexes.push(i);
        }
      }
      return indexes;
    };

    const rawIndexes = getRawIndexes();

    const normalizedReport = [];
    for (const row of report) {
      const normalizedRow = [];
      const iterator = rawIndexes;
      for (const i of iterator) {
        const val = row[i];
        normalizedRow.push(val);
      }
      normalizedReport.push(normalizedRow);
    }

    const reportHeaders = normalizedReport[0];
    const normalizedHeaders = reportHeaders.map(
      (header) => (header = headerMap.get(header))
    );

    normalizedReport.splice(0, 1, normalizedHeaders);
    if (cleanDataCallback) {
      return cleanDataCallback(normalizedReport);
    }
    return normalizedReport;
  }

  // NORMALIZE REPORT HEADERS
  function replaceSpaces(arr) {
    return arr.map((header) => header.replaceAll(" ", "_"));
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

  // CUSTOMER SPECIFIC DATA CLEANUP FUNCTIONS
  function cleanOldDominionReport(arr) {
    const poNumber = arr[0].indexOf("PO_#");
    const body = arr.slice(1);
    body.forEach((row) => {
      row[poNumber] = row[poNumber].replaceAll("#", "");
    });
    return arr;
  }

  // DEFINE CONFIGS FOR EACH CSV REPORT
  const netsuiteReport = getCSVFromGmail("ns-ltl-report");
  const netsuiteNormalizedHeaders = replaceSpaces([
    "Ship From",
    "Scheduled Ship Date",
    "Due Date",
    "DC Name",
    "PO #",
  ]);
  const netsuiteRawHeaders = [
    "Location",
    "Fulfillment Date",
    "Memo",
    "Name",
    "Customer PO #",
  ];
  const netsuiteConfig = new NormalizationConfig(
    netsuiteRawHeaders,
    netsuiteNormalizedHeaders
  );

  const carrierNormalizedHeaders = replaceSpaces([
    "PO #",
    "Actual Ship Date",
    "PRO Number",
    "Arrived At Carrier Yard",
    "Delivery Date",
    "Pallet Count",
    "Weight",
  ]);

  const oldDominionReport = getCSVFromGmail("odfl-ship-report");
  const oldDominonRawHeaders = [
    "Purchase Order Number",
    "Actual Pickup Date",
    "OD Pro#",
    "Arrival Date",
    "Delivery Date",
    "Pieces (skids/pallets)",
    "Weight",
  ];
  const oldDominionConfig = new NormalizationConfig(
    oldDominonRawHeaders,
    carrierNormalizedHeaders
  );

  const estesReport = getCSVFromGmail("estes-ship-report");
  const estesRawHeaders = [
    "Purchase Order #",
    "Pickup Date",
    "Pro #",
    "Arrival Date",
    "Delivery Date",
    "Pallets",
    "Weight*",
  ];
  const estesConfig = new NormalizationConfig(
    estesRawHeaders,
    carrierNormalizedHeaders
  );

  const normalizedNetsuiteReport = normalizeCSVReport(
    netsuiteReport,
    netsuiteConfig
  );
  const normalizedOldDominionReport = normalizeCSVReport(
    oldDominionReport,
    oldDominionConfig,
    cleanOldDominionReport
  );
  const normalizedEstesReport = normalizeCSVReport(estesReport, estesConfig);

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
  // rangeOfShipmentsToUpdate.setValues(updatedShipmentsArray);
}
