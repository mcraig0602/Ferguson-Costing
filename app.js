const xl = require("excel4node");
const csv = require("csvtojson");
const costing = `assets/November_Costing.csv`;
let wb = new xl.Workbook({
  jszip: {
    compression: "DEFLATE"
  },
  defaultFont: {
    size: 12,
    name: "Calibri",
    color: "FFFFFFFF"
  },
  dateFormat: "m/d/yy hh:mm:ss",
  workbookView: {
    activeTab: 1, // Specifies an unsignedInt that contains the index to the active sheet in this book view.
    autoFilterDateGrouping: true, // Specifies a boolean value that indicates whether to group dates when presenting the user with filtering options in the user interface.
    firstSheet: 1, // Specifies the index to the first sheet in this book view.
    minimized: false, // Specifies a boolean value that indicates whether the workbook window is minimized.
    showHorizontalScroll: true, // Specifies a boolean value that indicates whether to display the horizontal scroll bar in the user interface.
    showSheetTabs: true, // Specifies a boolean value that indicates whether to display the sheet tabs in the user interface.
    showVerticalScroll: true, // Specifies a boolean value that indicates whether to display the vertical scroll bar.
    tabRatio: 600, // Specifies ratio between the workbook tabs bar and the horizontal scroll bar.
    visibility: "visible", // Specifies visible state of the workbook window. ('hidden', 'veryHidden', 'visible') (§18.18.89)
    windowHeight: 17620, // Specifies the height of the workbook window. The unit of measurement for this value is twips.
    windowWidth: 28800, // Specifies the width of the workbook window. The unit of measurement for this value is twips..
    xWindow: 0, // Specifies the X coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
    yWindow: 440 // Specifies the Y coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
  },
  logLevel: 0, // 0 - 5. 0 suppresses all logs, 1 shows errors only, 5 is for debugging
  author: "Ballerific" // Name for use in features such as comments
});

let standard = wb.createStyle({
  font: {
    color: "#000000",
    size: 12
  }
});
let accounting = wb.createStyle({
  font: {
    color: "#000000",
    size: 12
  },
  numberFormat: "$#,##0.00; ($#,##0.00); -"
});
let number = wb.createStyle({
  font: {
    color: "#000000",
    size: 12
  },
  numberFormat: "####"
});
let floating = wb.createStyle({
  font: {
    color: "#000000",
    size: 12
  },
  numberFormat: "###.##"
});
let percent = wb.createStyle({
  font: {
    color: "#000000",
    size: 12
  },
  numberFormat: "#.00%; -#.00%; -"
});
let h1 = wb.createStyle({
  font: {
    color: "#000000",
    size: 48
  },
  alignment: { horizontal: "center" }
});
let h2 = wb.createStyle({
  font: {
    color: "#000000",
    size: 14
  },
  alignment: { horizontal: "center" }
});
const headers = {
  Master: {
    style: "standard",
    display: "Division",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: true
  },
  "Project: Invoice Number": {
    style: "number",
    display: "Inv #",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: true,
    Master: false
  },
  "Project: Project  Name": {
    style: "standard",
    display: "Project Name",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Work Order Name": {
    style: "standard",
    display: "WO Name",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Project: Work Site: Street": {
    style: "standard",
    display: "Address",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: false,
    Master: false
  },
  "Project: Estimated Project Scope": {
    style: "standard",
    display: "Work Being Done",
    Residential: true,
    Commercial: true,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Project: Final Sale Price": {
    style: "accounting",
    display: "Invoice Total",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: true
  },
  "Project: Incentive Amount": {
    style: "accounting",
    display: "CAC",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: true
  },
  "Project: Salesperson": {
    style: "standard",
    display: "Sales Rep",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: false,
    Master: false
  },
  "Project: Salesperson 2": {
    style: "standard",
    display: "Sales Rep 2",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  Crew: {
    style: "standard",
    display: "Crew",
    Residential: true,
    Commercial: true,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Soffit / Fascia Crew": {
    style: "standard",
    display: "Soffit/Fascia Crew",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Siding Crew": {
    style: "standard",
    display: "Siding Crew",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Gutter Crew": {
    style: "standard",
    display: "Gutter Crew",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  Vendor: {
    style: "standard",
    display: "Supplier",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Shortage?": {
    style: "bool",
    display: "Shortage",
    Residential: true,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Project: Total Shingle Squares": {
    style: "floating",
    display: "Total Bid Sq",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Total Shingle Squares": {
    style: "floating",
    display: "Need Title",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Conga - Total Shingles ordered to thirds": {
    style: "floating",
    display: "Total Sq",
    Residential: true,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Total Flat Roof Sq.": {
    style: "number",
    display: "Total MOD",
    Residential: true,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Project: Total Sq. of material": {
    style: "number",
    display: "Need Title",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Project: Invoiced Date": {
    style: "date",
    display: "Invoiced Date",
    Residential: false,
    Commercial: false,
    Exterior: false,
    Sales: false,
    Master: false
  },
  "Total Material Cost": {
    style: "accounting",
    display: "Material",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: true,
    Master: true
  },
  "Total Labor Rate": {
    style: "accounting",
    display: "Labor",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: true,
    Master: true
  },
  "Gross Profit": {
    style: "accounting",
    display: "Gross Profit",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: true,
    Master: true
  },
  "Gross Profit %": {
    style: "percent",
    display: "Gross Profit %",
    Residential: true,
    Commercial: true,
    Exterior: true,
    Sales: true,
    Master: true
  }
};
csv({
  noheader: true,
  output: "csv"
})
  .fromFile(costing)
  .then(jsonObj => {
    let mstrHead = jsonObj.splice(0, 1);
    let mstrCosting = extraLines(jsonObj);
    let monthyear = datemonth(mstrCosting);
    master(mstrCosting, monthyear);
    residential(mstrCosting, mstrHead, monthyear);
    commercial(mstrCosting, mstrHead, monthyear);
    exterior(mstrCosting, mstrHead, monthyear);
    salesmen(mstrCosting, mstrHead, monthyear);
    wb.write(`${monthyear[0]}-${monthyear[1]} Costing.xlsx`);
  });

extraLines = function(mstr) {
  const mstrLngth = mstr.length;
  mstr.splice(mstrLngth - 7, 7);
  return mstr;
};
datemonth = function(mstr) {
  let d = mstr[0][20].split("/");
  let m = d[0];
  let my = [];
  my[1] = d[2];

  switch (m) {
    case "1":
      my[0] = "January";
      break;
    case "2":
      my[0] = "February";
      break;
    case "3":
      my[0] = "March";
      break;
    case "4":
      my[0] = "April";
      break;
    case "5":
      my[0] = "May";
      break;
    case "6":
      my[0] = "June";
      break;
    case "7":
      my[0] = "July";
      break;
    case "8":
      my[0] = "August";
      break;
    case "9":
      my[0] = "September";
      break;
    case "10":
      my[0] = "October";
      break;
    case "11":
      my[0] = "November";
      break;
    case "12":
      my[0] = "December";
      break;
    default:
      my[0] = "Mike's Awesome";
  }
  return my;
};
master = function(mstr, my) {
  let ws = wb.addWorksheet("Master");
  ws.cell(1, 1, 1, 7, true)
    .string(`Exterior Cost and Profit Margin - ${my[0]} ${my[1]}`)
    .style(h1);
  let h = [
    "Division",
    "Invoice Total",
    "Customer Appreciation",
    "Labor",
    "Material",
    "Gross Profit",
    "Gross Profit %"
  ];
  let res = ["Residential", 0, 0, 0, 0, 0, 0];
  let comm = ["Commercial", 0, 0, 0, 0, 0, 0];
  let ext = ["Exterior", 0, 0, 0, 0, 0, 0];
  let total = ["Totals", 0, 0, 0, 0, 0, 0];
  let j = 1;
  h.forEach(head => {
    ws.cell(2, j)
      .string(head)
      .style(h2);
    j++;
  });
  mstr.forEach(master => {
    if (parseInt(master[0]) >= 100000) {
      if (master[5]) res[1] += parseFloat(master[5]);
      if (master[6]) res[2] += parseFloat(master[6]);
      if (master[22]) res[4] += parseFloat(master[22]);
      if (master[23]) res[3] += parseFloat(master[23]);
    } else if (parseInt(master[0]) <= 8000) {
      if (master[5]) comm[1] += parseFloat(master[5]);
      if (master[6]) comm[2] += parseFloat(master[6]);
      if (master[22]) comm[4] += parseFloat(master[22]);
      if (master[23]) comm[3] += parseFloat(master[23]);
    } else if (parseInt(master[0]) >= 8000 && parseInt(master[0]) <= 15000) {
      if (master[5]) ext[1] += parseFloat(master[5]);
      if (master[6]) ext[2] += parseFloat(master[6]);
      if (master[22]) ext[4] += parseFloat(master[22]);
      if (master[23]) ext[3] += parseFloat(master[23]);
    }
  });
  res[5] = res[1] - (res[3] + res[4]);
  res[6] = res[5] / res[1];
  comm[5] = comm[1] - (comm[3] + comm[4]);
  comm[6] = comm[5] / comm[1];
  ext[5] = ext[1] - (ext[3] + ext[4]);
  ext[6] = ext[5] / ext[1];

  total[1] = res[1] + comm[1] + ext[1];
  total[2] = res[2] + comm[2] + ext[2];
  total[3] = res[3] + comm[3] + ext[3];
  total[4] = res[4] + comm[4] + ext[4];
  total[5] = res[5] + comm[5] + ext[5];
  total[6] = total[5] / total[1];
  //Residential
  ws.cell(3, 1)
    .string(res[0])
    .style(standard);
  ws.cell(3, 2)
    .number(res[1])
    .style(accounting);
  ws.cell(3, 3)
    .number(res[2])
    .style(accounting);
  ws.cell(3, 4)
    .number(res[3])
    .style(accounting);
  ws.cell(3, 5)
    .number(res[4])
    .style(accounting);
  ws.cell(3, 6)
    .number(res[5])
    .style(accounting);
  ws.cell(3, 7)
    .number(res[6])
    .style(percent);
  //Commercial
  ws.cell(4, 1)
    .string(comm[0])
    .style(standard);
  ws.cell(4, 2)
    .number(comm[1])
    .style(accounting);
  ws.cell(4, 3)
    .number(comm[2])
    .style(accounting);
  ws.cell(4, 4)
    .number(comm[3])
    .style(accounting);
  ws.cell(4, 5)
    .number(comm[4])
    .style(accounting);
  ws.cell(4, 6)
    .number(comm[5])
    .style(accounting);
  ws.cell(4, 7)
    .number(comm[6])
    .style(percent);
  //Exterior
  ws.cell(5, 1)
    .string(ext[0])
    .style(standard);
  ws.cell(5, 2)
    .number(ext[1])
    .style(accounting);
  ws.cell(5, 3)
    .number(ext[2])
    .style(accounting);
  ws.cell(5, 4)
    .number(ext[3])
    .style(accounting);
  ws.cell(5, 5)
    .number(ext[4])
    .style(accounting);
  ws.cell(5, 6)
    .number(ext[5])
    .style(accounting);
  ws.cell(5, 7)
    .number(ext[6])
    .style(percent);
  //Totals
  ws.cell(6, 1)
    .string(total[0])
    .style(standard);
  ws.cell(6, 2)
    .number(total[1])
    .style(accounting);
  ws.cell(6, 3)
    .number(total[2])
    .style(accounting);
  ws.cell(6, 4)
    .number(total[3])
    .style(accounting);
  ws.cell(6, 5)
    .number(total[4])
    .style(accounting);
  ws.cell(6, 6)
    .number(total[5])
    .style(accounting);
  ws.cell(6, 7)
    .number(total[6])
    .style(percent);
};
salesmen = function(mstr, header, my) {
  let salesmanU = [];
  let salesmenInd = {};
  mstr.forEach(sale => {
    if (salesmanU.indexOf(sale[7]) === -1) {
      if (sale[7] !== "") salesmanU.push(sale[7]);
    }
  });
  salesmenInd.master = salesmanU;
  for (i = 0; i < salesmanU.length; i++) {
    console.log(salesmanU[i]);
    let ws = wb.addWorksheet(salesmanU[i]);
    ws.cell(1, 1, 1, 7, true)
      .string(`${salesmanU[i]}- ${my[0]} ${my[1]}`)
      .style(h1);
    let columns = [];
    let j = 1;
    header[0].forEach(head => {
      columns.push(head);
      if (headers[head]["Sales"]) {
        ws.cell(2, j)
          .string(headers[head]["display"])
          .style(standard);
        j++;
      }
    });
    let indS = [];
    mstr.forEach(row => {
      let k = 1;
      if (salesmanU[i] === row[7]) {
        indS.push(row);
        row.forEach((col, j) => {
          if (headers[columns[j]]["Sales"]) {
            console.log(indS.length);
            ws.cell(indS.length + 2, k)
              .string(col ? col : "")
              .style(standard);
            k++;
          }
        });
      }
    });
    /* mstr.forEach(row => {
      let k = 1;
      if (salesmanU[i] === row[7] && row[7] !== null && row[7] !== "") {
        indS.push(row);
        row.forEach((cols, j) => {
          if (headers[columns[j]]["Sales"]) {
            switch (headers[columns[j]]["style"]) {
              case "percent":
                ws.cell(indS.length + 2, k)
                  .number(cols ? parseFloat(cols) / 100 : 0)
                  .style(percent);
                k++;
                break;
              case "accounting":
                ws.cell(indS.length + 2, k)
                  .number(cols ? parseFloat(cols) : 0)
                  .style(accounting);
                k++;
                break;
              case "number":
                ws.cell(indS.length + 2, k)
                  .number(cols ? parseInt(cols) : 0)
                  .style(number);
                k++;
                break;
              case "floating":
                ws.cell(indS.length + 2, k)
                  .number(cols ? parseFloat(cols) : 0)
                  .style(floating);
                k++;
                break;
              case "bool":
                ws.cell(indS.length + 2, k)
                  .string(cols ? "Yes" : "No")
                  .style(standard);
                k++;
                break;
              default:
                ws.cell(indS.length + 2, k)
                  .string(cols ? cols : "")
                  .style(standard);
                k++;
            }
          }
        });
      }
    }); */
  }
};
residential = function(mstr, header, my) {
  let ws = wb.addWorksheet("Residential");
  ws.cell(1, 1, 1, 14, true)
    .string(`Residential Cost and Profit Margin - ${my[0]} ${my[1]}`)
    .style(h1);
  let resident = [];
  let srtInv = [];
  let sortRes = [];
  let columns = [];
  let j = 1;
  header[0].forEach(head => {
    columns.push(head);
    if (headers[head]["Residential"]) {
      ws.cell(2, j)
        .string(headers[head]["display"])
        .style(h2);
      j++;
    }
  });
  mstr.forEach(res => {
    if (parseInt(res[0]) >= 100000) {
      resident.push(res);
      srtInv.push(parseInt(res[0]));
    }
  });
  srtInv.forEach(s => {
    resident.forEach(res => {
      if (parseInt(res[0]) === s) sortRes.push(res);
    });
  });
  sortRes.forEach((row, i) => {
    let k = 1;
    row.forEach((cols, j) => {
      if (headers[columns[j]]["Residential"]) {
        switch (headers[columns[j]]["style"]) {
          case "percent":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) / 100 : 0)
              .style(percent);
            k++;
            break;
          case "accounting":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) : 0)
              .style(accounting);
            k++;
            break;
          case "number":
            ws.cell(i + 3, k)
              .number(cols ? parseInt(cols) : 0)
              .style(number);
            k++;
            break;
          case "floating":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) : 0)
              .style(floating);
            k++;
            break;
          case "bool":
            ws.cell(i + 3, k)
              .string(cols ? "Yes" : "No")
              .style(standard);
            k++;
            break;
          default:
            ws.cell(i + 3, k)
              .string(cols ? cols : "")
              .style(standard);
            k++;
        }
      }
    });
  });
};
commercial = function(mstr, header, my) {
  let ws = wb.addWorksheet("Commercial");
  ws.cell(1, 1, 1, 11, true)
    .string(`Commercial Cost and Profit Margin - ${my[0]} ${my[1]}`)
    .style(h1);
  let comm = [];
  let srtInv = [];
  let sortCom = [];
  let columns = [];
  let j = 1;
  header[0].forEach(head => {
    columns.push(head);
    if (headers[head]["Commercial"]) {
      ws.cell(2, j)
        .string(headers[head]["display"])
        .style(h2);
      j++;
    }
  });
  mstr.forEach(com => {
    if (parseInt(com[0]) <= 8000) {
      comm.push(com);
      srtInv.push(parseInt(com[0]));
    }
  });
  srtInv.forEach(s => {
    comm.forEach(com => {
      if (parseInt(com[0]) === s) sortCom.push(com);
    });
  });
  sortCom.forEach((row, i) => {
    let k = 1;
    row.forEach((cols, j) => {
      if (headers[columns[j]]["Commercial"]) {
        switch (headers[columns[j]]["style"]) {
          case "percent":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) / 100 : 0)
              .style(percent);
            k++;
            break;
          case "accounting":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) : 0)
              .style(accounting);
            k++;
            break;
          case "number":
            ws.cell(i + 3, k)
              .number(cols ? parseInt(cols) : 0)
              .style(number);
            k++;
            break;
          case "floating":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) : 0)
              .style(floating);
            k++;
            break;
          case "bool":
            ws.cell(i + 3, k)
              .string(cols ? "Yes" : "No")
              .style(standard);
            k++;
            break;
          default:
            ws.cell(i + 3, k)
              .string(cols ? cols : "")
              .style(standard);
            k++;
        }
      }
    });
  });
};
exterior = function(mstr, header, my) {
  let ws = wb.addWorksheet("Exterior");
  ws.cell(1, 1, 1, 10, true)
    .string(`Exterior Cost and Profit Margin - ${my[0]} ${my[1]}`)
    .style(h1);
  let ext = [];
  let srtInv = [];
  let sortExt = [];
  let columns = [];
  let j = 1;
  header[0].forEach(head => {
    columns.push(head);
    if (headers[head]["Exterior"]) {
      ws.cell(2, j)
        .string(headers[head]["display"])
        .style(h2);
      j++;
    }
  });
  mstr.forEach(extr => {
    if (parseInt(extr[0]) >= 8000 && parseInt(extr[0]) <= 15000) {
      ext.push(extr);
      srtInv.push(parseInt(extr[0]));
    }
  });
  srtInv.forEach(s => {
    ext.forEach(extr => {
      if (parseInt(extr[0]) === s) sortExt.push(extr);
    });
  });
  sortExt.forEach((row, i) => {
    let k = 1;
    row.forEach((cols, j) => {
      if (headers[columns[j]]["Exterior"]) {
        switch (headers[columns[j]]["style"]) {
          case "percent":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) / 100 : 0)
              .style(percent);
            k++;
            break;
          case "accounting":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) : 0)
              .style(accounting);
            k++;
            break;
          case "number":
            ws.cell(i + 3, k)
              .number(cols ? parseInt(cols) : 0)
              .style(number);
            k++;
            break;
          case "floating":
            ws.cell(i + 3, k)
              .number(cols ? parseFloat(cols) : 0)
              .style(floating);
            k++;
            break;
          case "bool":
            ws.cell(i + 3, k)
              .string(cols ? "Yes" : "No")
              .style(standard);
            k++;
            break;
          default:
            ws.cell(i + 3, k)
              .string(cols ? cols : "")
              .style(standard);
            k++;
        }
      }
    });
  });
};
