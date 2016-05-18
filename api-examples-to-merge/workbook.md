# Workbook

### getSelectedRange()

```js
Excel.run(function (ctx) { 
	var selectedRange = ctx.workbook.getSelectedRange();
	selectedRange.load('address');
	return ctx.sync().then(function() {
			console.log(selectedRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```