### Add Custom IconSet
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet)
    conditionalFormat.iconOrNull.criteria = [
        { type: "percent", formula: 10, operator: "lessThan", customIcon: Excel.icons.fiveArrows.yellowDownInclineArrow },
        { type: "percent", formula: 30, operator: "lessThan", customIcon: Excel.icons.fourRating.fourBars },
        { type: "percent", formula: 50, operator: "lessThan", customIcon: null },
        { type: "percent", formula: 70, operator: "lessThan", customIcon: Excel.icons.fiveQuarters.blackCircle }
    ];
    
    return ctx.sync().then(function () {
        console.log("Added new custom icon conditional formatting.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Add Preset IconSet
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet).iconOrNull.style = Excel.IconSet.threeArrows;
    return ctx.sync().then(function () {
        console.log("Added new yellow three arrow icon set.");
    })
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```