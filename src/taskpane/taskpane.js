/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-table").onclick = () => tryCatch(createTable);
    document.getElementById("filter-table").onclick = () => tryCatch(filterTable);
    document.getElementById("sort-table").onclick = () => tryCatch(sortTable);
    document.getElementById("create-chart").onclick = () => tryCatch(createChart);
    document.getElementById("freeze-header").onclick = () => tryCatch(freezeHeader);
    document.getElementById("custom-chart").onclick = customChart;
    document.getElementById("load-inforiver").onclick = loadInforiver;
  }
});

function loadInforiver() {
  document.getElementById("infoRiver").innerHTML = '<object type="text/html" data="test.html"></object>';
}

async function customChart() {
  await Excel.run((context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("B2:C1000");
    range.load("values");
    return context.sync().then(() => {
      function getData(n) {
        var arr = [],
          i,
          a,
          b,
          c,
          spike;
        for (i = 0; i < n; i = i + 1) {
          if (i % 100 === 0) {
            a = 2 * Math.random();
          }
          if (i % 1000 === 0) {
            b = 2 * Math.random();
          }
          if (i % 10000 === 0) {
            c = 2 * Math.random();
          }
          if (i % 50000 === 0) {
            spike = 10;
          } else {
            spike = 0;
          }
          arr.push([i, 2 * Math.sin(i / 100) + a + b + c + spike + Math.random()]);
        }
        return arr;
      }
      var data = getData(100000);

      // const chartContext = document.getElementById("myChart").getContext("2d");
      // Sample data for the chart
      // var chartData = [
      //   ["Apples", 10],
      //   ["Oranges", 5],
      //   ["Bananas", 7],
      //   ["Grapes", 8],
      //   ["Pears", 2],
      // ];

      // Create the chart
      Highcharts.chart("myChart", {
        chart: {
          type: "area",
          zoomType: "x",
          panning: true,
          panKey: "shift",
        },

        boost: {
          useGPUTranslations: true,
        },

        title: {
          text: "Highcharts drawing " + data.length + " points",
        },

        subtitle: {
          text: "Using the Boost module",
        },

        tooltip: {
          valueDecimals: 2,
        },

        series: [
          {
            data: data,
          },
        ],
      });
    });
  });
}

// function insertAddIn() {
//   const addInContent =
//     '<div id="myAddIn" style="width: 100%; height: 300px; background-color: yellow;">My Add-in Content</div>';

//   Office.context.document.setSelectedDataAsync(
//     addInContent,
//     { coercionType: Office.CoercionType.Ooxml },
//     function (result) {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         console.log("Add-in inserted successfully.");
//       } else {
//         console.error("Error inserting add-in:", result.error.message);
//       }
//     }
//   );
// }

async function createTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
    ]);

    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    await context.sync();
  });
}

async function filterTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    // eslint-disable-next-line office-addins/load-object-before-read
    const categoryFilter = expensesTable.columns.getItem("Category").filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);

    await context.sync();
  });
}

async function sortTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const sortFields = [
      {
        key: 1, // Merchant column
        ascending: false,
      },
    ];

    expensesTable.sort.apply(sortFields);

    await context.sync();
  });
}

async function createChart() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const dataRange = expensesTable.getDataBodyRange();

    const chart = currentWorksheet.charts.add("ColumnClustered", dataRange, "Auto");

    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "Right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 10;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = "Value in \u20AC";

    await context.sync();
  });
}

async function freezeHeader() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);

    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

let dialog = null;

function openDialog() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/popup.html",
    { height: 45, width: 55 },
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
  );
}

function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}
