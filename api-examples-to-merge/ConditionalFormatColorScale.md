### Add Preset Color Scale
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    cf.colorScaleOrNull.minimum.color = "red";
    cf.colorScaleOrNull.midpoint.color = "yellow";
    cf.colorScaleOrNull.maximum.color = "green"; 

    return ctx.sync().then(function () {
        console.log("Added red-yellow-green color scale.");
    })
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### Add Custom Color Scale
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    cf.colorScaleOrNull.minimum = {
        type: Excel.ConditionalFormatRuleType.percent,
        formula: 10,
        color: "red"
    };
    cf.colorScaleOrNull.maximum = {
        type: Excel.ConditionalFormatRuleType.percent,
        formula: 90,
        color: "blue"
    };
    return ctx.sync().then(function () {
        console.log("Added new red-blue 10%-90% color scale format");
    })
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```