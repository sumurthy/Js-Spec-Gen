# Chart Data Labels
### Getter

Make Series Name shown in Datalabels and set the `position` of datalabels to be "top";

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.datalabels.showValue = true;
	chart.datalabels.position = "top";
	chart.datalabels.showSeriesName = true;
	return ctx.sync().then(function() {
			console.log("Datalabels Shown");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
