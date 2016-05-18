# Chart Gridlines

### Getter and setter

Get the `visible` of Major Gridlines on value axis of Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var majGridlines = chart.axes.valueaxis.majorGridlines;
	majGridlines.load('visible');
	return ctx.sync().then(function() {
			console.log(majGridlines.visible);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set to show major gridlines on valueAxis of Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.majorgridlines.visible = true;
	return ctx.sync().then(function() {
			console.log("Axis Gridlines Added ");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
