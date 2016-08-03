## ConditionalFormatting Custom Types

### Average:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);

    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.average;
    conditionalFormat.customOrNull.rule.average.selection = Excel.ConditionalFormatAverageSelection.below;
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.bad;
    return ctx.sync().then(function () {
        console.log("Added new below average rule type, formatted bad.");
	});
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Between: 
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.between;
    conditionalFormat.customOrNull.rule.between = {
        inclusive: true,
        lowerBound: "A2",
        upperBound: "B3"
    };
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.good;
    return ctx.sync().then(function () {
        console.log("Added new between rule based on cells A2 and B3, formatted good");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### Formula: 
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.formula;
    conditionalFormat.customOrNull.rule.formula = "=COUNTIF($D$2:$D11,D2)>1";
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.neutral;
    return ctx.sync().then(function () {
        console.log("Added new custom formula rule with neutral formatting");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Top/Bottom Count:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.count;
    conditionalFormat.customOrNull.rule.count = {
        count: 10,
        direction: Excel.ConditionalFormatDirection.top
    };
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.heading1;
    return ctx.sync().then(function () {
        console.log("Added new top 10 items formatted like heading1");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Top/Bottom Percent:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.percent;
    conditionalFormat.customOrNull.rule.percent = {
        percent: 10,
        direction: Excel.ConditionalFormatDirection.bottom
    };
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.title;
    return ctx.sync().then(function () {
        console.log("Added new conditional format on bottom 10% formatted like a title");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### TextContains:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.textContains;
    conditionalFormat.customOrNull.rule.textContains = {
        type: Excel.StringMatchType.beginningWith,
        text: "G"
    };

    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.good;
    return ctx.sync().then(function () {
        console.log("Added new conditional format for strings beginning with 'G' - formatted 'good'");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### DatesOccurring:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.textContains;
    conditionalFormat.customOrNull.rule.textContains = {
        type: Excel.StringMatchType.beginningWith,
        text: "G"
    };

    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.good;
    return ctx.sync().then(function () {
        console.log("Added new conditional format for strings beginning with 'G' - formatted 'good'");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Blanks:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.blank;
    conditionalFormat.customOrNull.rule.blank = true;
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.bad;
    return ctx.sync().then(function () {
        console.log("Added new conditional format to format blank cells");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Errors:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.error;
    conditionalFormat.customOrNull.rule.error = true;
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.bad;
    return ctx.sync().then(function () {
        console.log("Added new conditional format to format cells with errors");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Unique: Formatting Duplicates
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.unique;
    conditionalFormat.customOrNull.rule.unique = false;
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.bad;
    return ctx.sync().then(function () {
        console.log("Added new conditional format to format duplicates");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```