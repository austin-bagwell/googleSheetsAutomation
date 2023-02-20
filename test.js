// a significantly more generalized/abstracted normalization sequence
// where report === [[]]:
// pulls specific columns from a report with n rows and x columns
// normalized the header text to match a set of 'normal' headers
// so you can make the headers match whatever you have in your existing sheet

// will need to do some testing to make sure everything actually works
// will also need to add some error handing up in this B
class NormalizationConfig {
  constructor(rawHeaders = [], nrmlHeaders = []) {
    this.rawHeaders = rawHeaders;
    this.normalHeaders = nrmlHeaders;
  }
}

function normalizeCSVReport(report = [[]], config = {}) {
  const { rawHeaders, normalHeaders } = config;

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

  //   replace remaining raw headers with normalized headers
  normalizedReport.splice(0, 1, normalHeaders);
  return normalizedReport;
}

// I think it will be important that the indexes of these two arrays line up
// EG raw Purchase Order will need to line up with normalized PO_#, etc.
// FIXME indexes of raw/normal headers must line up 1:1
// will be replaced as raw[i] === normal[i]
const testRawHeaders = ["Purchase Order Number", "PRO NUMBER", "pallets/skids"];
const testNormalHeaders = ["PO_#", "PRO_#", "Pallet_Count"];

const testConfig = new NormalizationConfig(testRawHeaders, testNormalHeaders);
const testReport = [
  [
    "gross",
    "Purchase Order Number",
    "PRO NUMBER",
    "pallets/skids",
    "not this col",
  ],
  ["", "PO1234", "Pro987665", "7", "dont want this"],
  ["", "order777", "xyzPRO100", "8", "dont want this"],
];

const test = normalizeCSVReport(testReport, testConfig);
console.log(test);
