# Worksheet Functions

### sum()
```js
Excel.run(function (ctx) { 
	var result = ctx.workbook.functions.sum(1, 2, 3.5, 10.333);
	result.load();
	return ctx.sync()
	.then(function(){
		console.log(result.value)
		}); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```