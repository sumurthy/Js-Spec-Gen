# Chart Point Collection

### Getter 

Get the names of points in the points collection

```js
Excel.run(function (ctx) { 
	var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
	pointsCollection.load('items');
	return ctx.sync().then(function() {
		console.log("Points Collection loaded");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Get the number of points

```js
Excel.run(function (ctx) { 
	var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
	pointsCollection.load('count');
	return ctx.sync().then(function() {
		console.log("points: Count= " + pointsCollection.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getItemAt(index: number)
Set the border color for the first points in the points collection

```js
Excel.run(function (ctx) { 
	var point = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
	points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
	return ctx.sync().then(function() {
		console.log("Point Border Color Changed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```