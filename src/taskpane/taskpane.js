/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    logic();
    
  }
});

function logic() {
  // Create myTable
  Excel.run(function (ctx) {
      return ctx.sync().then(function () {
          //Create a new binding for A1
          Office.context.document.bindings.addFromNamedItemAsync("Sheet1!A1", Office.BindingType.Text, { id: "A1" }, function (asyncResult) {
              if (asyncResult.status == "failed") {
                  console.log("Action failed with error: " + asyncResult.error.message);
              }
              else {
                  // If successful, add the event handler to the table binding.
                  Office.select("bindings#A1").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
              }
          });
      })
      .catch(function (error) {
          console.log(JSON.stringify(error));
      });
  });
};
  
  // When data in A1 is changed, this event is triggered.
  function onBindingDataChanged(_eventArgs) {
    Excel.run(function (ctx) {
        // Highlight the table in orange to indicate data changed.
        var sheet = ctx.workbook.worksheets.getItem("Sheet1");
        var cell = sheet.getRange("A1");
        var fill = cell.format.fill;
        fill.load("color");
        return ctx.sync().then(function () {
          if (fill.color != "Orange") {
            cell.format.fill.color = "Orange";
            console.log("The value in this table got changed!");
          }
        }).then(ctx.sync).catch(function (error) {
            console.log(JSON.stringify(error));
        });
    });
  } 

export async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.warn(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
