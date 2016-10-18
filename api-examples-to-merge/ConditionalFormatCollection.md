### getItemAt(index: int)
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormats = range.conditionalFormats;
    var conditionalFormat = conditionalFormats.getItemAt(3);
    return ctx.sync().then(function () {
        console.log("Conditional Format at Item 3 Loaded");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### clearAll()
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormats = range.conditionalFormats;
    var conditionalFormat = conditionalFormats.clearAll();
    return ctx.sync().then(function () {
        console.log("Cleared all conditional formats from this range.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### add(type: ConditionalFormatType)
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    conditionalFormat.iconOrNull.style = "YellowThreeArrows";
    return ctx.sync().then(function () {
        console.log("Added new yellow three arrow icon set.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### getCount()
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    conditionalFormat.iconOrNull.style = Excel.IconSet.fourTrafficLights;
    var cfCount = range.conditionalFormats.getCount(); 

    return ctx.sync().then(function () {
        console.log("Count: " + cfCount.value);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```