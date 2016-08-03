### Delete()
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.iconOrNull.style = Excel.IconSet.threeArrows;
    cf.delete();

    return ctx.sync().then(function (ctx) {
        console.log("Removed conditional format from worksheet.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### DeleteFromCurrentRange()
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);

    /// remove conditional format from just A2
    var subsetRangeAddress = "A2";
    var subsetRange = ctx.workbook.worksheets.getItem(sheetName).getRange(subsetRangeAddress);
    var subsetCF = range.conditionalFormats.getItemAt(0);
    subsetCF.deleteFromCurrentRange();
    return ctx.sync().then(function () {
        console.log("Removed conditional format just from A2.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### GetRangeOrNull()
#### Range Success:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    var rangeVal = cf.getRangeOrNull();
  
    return ctx.sync().then(function () {
        console.log("Range: " + rangeVal);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

#### Range = NULL
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);

    /// remove conditional format from just A2
    var subsetRangeAddress = "A2";
    var subsetRange = ctx.workbook.worksheets.getItem(sheetName).getRange(subsetRangeAddress);
    range.conditionalFormats.getItemAt(0).deleteFromCurrentRange();

    var rangeOnConditionalFormat = cf.getRangeOrNull();
    return ctx.sync().then(function () {
        console.log("Range is Discontiguous:" + rangeOnConditionalFormat);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Getters and Setters
#### Get Priority
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.iconOrNull.style = Excel.IconSet.threeArrows;

    cf.load('priority');
    return ctx.sync().then(function () {
        console.log("Local Priority" + cf.priority);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

#### StopIfTrue, get
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.iconOrNull.style = Excel.IconSet.threeArrows;

    cf.load('stopIfTrue');
    return ctx.sync().then(function () {
        console.log("Stop If True?" + cf.stopIfTrue);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

#### Set Reverse
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.reverse = true;
    return ctx.sync().then(function () {
        console.log("Reverse: true?" + cf.reverse);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
#### Get Type
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.load('type');
    return ctx.sync().then(function () {
        console.log("Error: " + Error);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```