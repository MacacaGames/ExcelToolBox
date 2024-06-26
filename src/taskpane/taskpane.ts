/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("selectRangeToJsonButton").onclick = selectRangeToJson;
  // document.getElementById("selectWholeSheetToJsonButton").onclick = selectWholeSheetToJson;
  //
  // document.getElementById("insertNoDataIntButton").onclick = insertNoDataInt;
  // document.getElementById("insertNoDataEnumButton").onclick = insertNoDataEnum;
  // document.getElementById("insertNoDataBoolButton").onclick = insertNoDataBool;
  // document.getElementById("insertNoDataArrayButton").onclick = insertNoDataArray;
  // document.getElementById("insertNoDataTextButton").onclick = insertNoDataText;
});

async function selectRangeToJson() {
  try {
    const rangeInput = (document.getElementById("rangeInput") as HTMLInputElement).value.trim();

    if (!rangeInput) {
      console.error('Please enter a valid range.');
      return;
    }
    
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const excelRange = sheet.getRange(rangeInput);

      excelRange.load("values");
      await context.sync();

      const data = excelRange.values;
      await dataArrayToJson(data);
    });
  } catch (error) {
    console.error(error);
  }
}

async function selectWholeSheetToJson() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const excelRange = sheet.getUsedRange();
      excelRange.load("values");
      await context.sync();
      
      const data = excelRange.values;
      await dataArrayToJson(data);
    });
  } catch (error) {
    console.error(error);
  }
}

async function dataArrayToJson(data: any[][]) {
  const headers = data[0];
  const rows = data.slice(1);

  let jsonArray = [];

  for (let i = 0; i < rows.length; i += 50) {
    const batch = rows.slice(i, i + 50);
    const jsonBatch = batch.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        let value = row[index];
        if(value === "NoData_Int" || value === "NoData_Enum") {
          value = 0;
        } else if (value === "NoData_Bool") {
          value = false;
        } else if (value === "NoData_Json") {
          value = {};
        } else if (value === "NoData_Array") {
          value = [];
        } else if (value === "NoData_Text") {
          value = "";
        } else {
          try {
            const parsedValue = JSON.parse(value);
            if (typeof parsedValue === 'object') {
              value = parsedValue;
            }
          } catch (e) {}
        }
        obj[header] = value;
      });
      return obj;
    });

    jsonArray = jsonArray.concat(jsonBatch);

    // Intermediate sync to keep the UI responsive
    await new Promise(resolve => setTimeout(resolve, 50));
  }

  const jsonString = JSON.stringify(jsonArray, null, 2);
  await openJsonDialog(jsonString);
}

async function openJsonDialog(jsonString: string) {
  localStorage.setItem('json', jsonString);

  Office.context.ui.displayDialogAsync(
      location.origin + '/taskpane/dialog.html',
      { height: 50, width: 50 },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error('Failed to open dialog:', result.error.message);
        }
      }
  );
}

async function insertNoDataInt() {
  await insertData("NoData_Int");
}

async function insertNoDataEnum() {
  await insertData("NoData_Enum");
}

async function insertNoDataBool() {
  await insertData("NoData_Bool");
}

async function insertNoDataJson() {
  await insertData("NoData_Json");
}

async function insertNoDataArray() {
  await insertData("NoData_Array");
}

async function insertNoDataText() {
  await insertData("NoData_Text");
}

async function insertData(data: any) {
  try {
    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.values = data;
      await context.sync();
    });
  } catch (error) {
    console.error("Error:", error);
  }
}