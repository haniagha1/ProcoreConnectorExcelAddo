/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('fetch-data').onclick = fetchDataAndWriteToExcel;
  }
});

async function fetchDataAndWriteToExcel() {
  try {
    console.log('Fetching data');
    const response = await fetch('https://api.sampleapis.com/futurama/info'); // Replace with your API URL
    const data = await response.json();
    console.log(data.synopsis)
    writeDataToExcel(data);
  } catch (error) {
    console.error(error);
    document.getElementById('message').innerText = 'Error fetching data';
  }
}

function writeDataToExcel(data) {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1").getResizedRange(data[0].length - 1, data[0].length - 1);
    range.values = data[0].synopsis;
    range.format.autofitColumns();
    await context.sync();
  });
}
