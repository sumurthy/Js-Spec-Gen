### Add Preset Databar
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    conditionalFormat.dataBarOrNull.positiveFormat = {
        color: "green",
        gradient: false,
    }; // does this make sense/easy to use???
    // perhaps: conditionalFormat.addDataBar("green", false);

    return ctx.sync().then(function () {
        console.log("Added green data bar, without gradient");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### Add Custom Databar
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    conditionalFormat.dataBarOrNull.showDataBarOnly = false;
    conditionalFormat.dataBarOrNull.Direction = "LeftToRight";
    conditionalFormat.dataBarOrNull.lowerBoundRule = { type: "percent", formula: "10" };
    conditionalFormat.dataBarOrNull.upperBoundRule = { type: "percent", formula: "90" };
    conditionalFormat.dataBarOrNull.positiveFormat = { color: "blue", gradient: false };

    return ctx.sync().then(function () {
        console.log("Added new blue 10%-90% databar format");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```