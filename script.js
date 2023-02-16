const odDataSheet = SpreadsheetApp.getActive().getSheetByName("ODFL Data");
const nsDataSheet = SpreadsheetApp.getActive().getSheetByName("NS Data");
const datasetSheet = SpreadsheetApp.getActive().getSheetByName("Dataset");
const dcInfoSheet = SpreadsheetApp.getActive().getSheetByName("DC Info");

// TODO
// add some actual error handling - try/catch, maybe a way to email me or report owner when errors occur

// WORKFLOW
/*
get CSVs from gmail
do data cleanup on them as needed (Ex: remove # from OD PRO number)
NORMALIZE shipment data (OD/Estes/future carriers)
put normalized shipping info into one array || array of objects
pull dataset into working memory
add new shipments from NS report into dataset
use normalized shipping data to update the dataset with new info (ex. order was delivered)
push updated dataset back into spreadsheet
autofill any 'right hand' formula columns (don't like this approach but... meh)
*/

function main() {
  // TODO move this dataset data to a more relevant place
  const datasetData = SpreadsheetApp.getActive()
    .getSheetByName("Dataset")
    .getDataRange()
    .getValues();

  // STEP 0 - GET DATA FROM CSV REPORTS
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

  const gmailLabels = [
    "estes-ship-report",
    "odfl-ship-report",
    "ns-ltl-report",
  ];
  const estesReport = getCSVFromGmail(gmailLabels[0]);
  const oldDominionReport = getCSVFromGmail(gmailLabels[1]);
  const netsuiteReport = getCSVFromGmail(gmailLabels[2]);

  // FIXME - may not need this-updated NS report to use 'shipping addressee' instead, which is already cleaner
  // will require me to do some find and replace on actual dataset but that's easier than maintaining a helper sheet
  // STEP 1 - CLEAN DATA FROM CSV REPORTS
  // -- Netsuite customer names as shown on the CSV report need tidying up
  // -- this uses the names found in DC Info instead of the ugly Netsuite customer names
  function cleanNetsuiteCSVReport() {
    const dcInfo = SpreadsheetApp.getActive()
      .getSheetByName("DC Info")
      .getDataRange()
      .getValues();
    const dcInfoCustID = dcInfo[0].indexOf("Internal ID");
    const dcInfoName = dcInfo[0].indexOf("Name");

    const netsuiteHeaders = netsuiteReport.slice(0, 1).flat();
    const internalID = netsuiteHeaders.indexOf("Internal ID");
    const name = netsuiteHeaders.indexOf("Name");
    const netsuiteBody = netsuiteReport.slice(1);

    const dcInfoObject = {};
    dcInfo.forEach((row) => {
      dcInfoObject[row[dcInfoCustID].toString()] = row[dcInfoName];
    });
    netsuiteBody.forEach((row) => {
      row[name] = dcInfoObject[row[internalID.toString()]];
    });
    return netsuiteBody;
  }

  // OD sends purchase order field with a '#' prefix for some reason
  function cleanOldDominionCSVReport() {
    const oldDominionHeaders = oldDominionReport.slice(0, 1).flat();
    const oldDominionBody = oldDominionReport.slice(1);
    const poNumber = oldDominionHeaders.indexOf("Purchase Order Number");
    oldDominionBody.forEach((row) => {
      row[poNumber] = row[poNumber].replace("#", "");
    });
    return oldDominionBody;
  }

  // TODO determine what cleanup needs to happen for Estes
  // function cleanEstesCSVReport() {};

  const cleanedNetsuiteData = cleanNetsuiteCSVReport();
  const cleanedOldDominionData = cleanOldDominionCSVReport();
  // const cleanedEstesData = cleanEstesCSVReport();

  // STEP 2 - NORMALIZE SHIPMENT DATA (OD/Estes/other carriers)
  // STEP a - define a class that describes 'normalized' data that is expected
  // aka info that relates to these headers:
  // [actual ship, arrived at yard, delivery date, proNum, pallet count, weight];
  // -- OD/estes/any carrier will require their own normalization function

  class NormalizedShipment {
    constructor(
      poNum,
      shipDate,
      arrivedAtYard,
      delivered,
      pro,
      pallets,
      weight
    ) {
      this.poNum = poNum;
      this.shipDate = shipDate;
      this.arrivedAtYard = arrivedAtYard;
      this.delivered = delivered;
      this.pro = pro;
      this.pallets = pallets;
      this.weight = weight;
    }
  }

  // 2/15 evening - it works but it is so ugly
  // extracting specific indexes of data from a 2D array
  // gotta be a way to extract this logic to make it reusable for other carrier data

  function normalizeOldDominionData() {
    // for every row of the body, I want to return a NormalizedShipment object
    // those objects get pushed into an array
    const normalized = [];
    const odHeaders = oldDominionReport.slice(0, 1).flat();

    const headersIWant = [
      "Delivery Date",
      "Arrival Date",
      "Actual Pickup Date",
      "OD Pro#",
      "Purchase Order Number",
      "Pieces (skids/pallets)",
      "Weight",
    ];
    const indexesOfHeadersIWant = [];
    headersIWant.forEach((header) => {
      indexesOfHeadersIWant.push(odHeaders.indexOf(header));
    });
    const odBody = oldDominionReport.slice(1);
    const testI = indexesOfHeadersIWant[4];

    for (let row = 0; row < odBody.length; row++) {
      const newRow = [];
      const oldRow = odBody[row];
      for (let col = 0; col < indexesOfHeadersIWant.length; col++) {
        const magicIndex = indexesOfHeadersIWant[col];
        newRow.push(oldRow[magicIndex]);
      }
      normalized.push(newRow);
    }
    return normalized;
  }
  // Logger.log(normalizeOldDominionData())
  normalizeOldDominionData();

  // STEP 3 - ADD NEW NS DATA TO DATASET
  // -- Array.of(dataset.length) => [,,,,,, etc.] then add NS data into the correct index

  // STEP 4 - UPDATE DATASET WITH NORMALIZED SHIPMENT DATA
  // -- if dataset[row][colToUpdate] === '', add the normalData[colToUpdate] value (could sitll be "", just means no new data)
  // STEP 5 - PUSH UPDATED DATASET BACK TO SPREADSHEET
  // STEP 6 - AUTOFILL 'RIGHT HAND' SHEETS FORMULAS
  // -- These formulas feed the reporting dashboard stuff
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

/*

class Shipment(...[headers,from,datasetSheet,excluding,sheetsFormulas]) {
  // should they be named exactly as they appear (PO #) O
  // should they be normalized?
  // could do an inbetween and 'PO #'.replace(" ", "_"), convert back for checking index equality
}

// would require manually setting my headers range & all updatable cols would 
// need to be contiguous to avoid errors probably - slightly dangerous?
const relevantDatasetHeaders = datasetHeaders.slice(0,14)

datasetRows.push(all new NS orders in as an array[id,name,etc., fill remaining length with '' if needed])
then 
beginnng at the oldest in-transit shipment, convert all rows to new Shipments - push those into an array
then
use Object.assign(new Shipments, normalizedShippingData), where the object properties match exactly
as in, i have my new Shipment {name=kroger,poNum=4321, ship_date=1/1/23, deliver=''}
where NormalizedShippingData {poNum=4321,ship_date=1/1/23, deliver='1/3/23'};
common key is poNum to get the objects to line up correctly - Ex:
if Shipment.poNum === Normalized.poNum, Object.assign(Shipment, Normalized)
the only *new* data for any given shipment should be the updated date fields
and I think Object.assign will overwrite source data w/ target data
and would thus update deliveryDate (and poNum technically) while leaving the rest alone

// TODO dig up the 'instantiate a Class with n number of params... got that saved in my old ltlDashboard thing I think
*/

// proof of concept
// I CAN use Object.assign() to update by Shipment object
// using a Normalized object.
// Shipment and Normalized need to have matching keys
// so I should source my keys directly from my Dataset headers
// with some considerations ease of access (aka using _ instead of spaces)
// function shipmentTest() {
//   const shipment = {
//     'Internal_ID': '12345',
//     'DC_Name': 'kroger',
//     'PO_#': 'P9876',
//     'Scheduled_Ship_Date': '1/1/23',
//     'Delivery_Date': ''
//   }

//   const normalized = {
//     'PO_#': 'P9876',
//     'Actual_Ship_Date': '1/1/23',
//     'Delivery_Date': '1/4/23'
//   }

//   return Object.assign(shipment, normalized);
// }

// okay, got a decent working start here
function createConstructor(...propNames) {
  // TRY -- required to pass [] of strings into constructor
  // CATCH -- if anything else is passed, explode
  const cleanedProps = [...propNames].map((prop) => prop.replaceAll(" ", "_"));
  return class {
    constructor(...propValues) {
      cleanedProps.forEach((name, idx) => {
        this[name] = propValues[idx];
      });
    }
  };
}

function testShipment() {
  const testHeaders = ["PO #", "Actual Ship Date"];
  // const cleanHeaders = testHeaders.map(head => head.replaceAll(" ", "_"));
  // Logger.log(cleanHeaders);
  const Shipment = createConstructor(...testHeaders);
  const testShipment = new Shipment();
  Logger.log(testShipment);
}
