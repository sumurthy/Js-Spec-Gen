# NamedItem

### getRange()

Returns the Range object that is associated with the name. `null` if the name is not of the type `Range`. Note: This API currently supports only the Workbook scoped items.**

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var range = names.getItem('MyRange').getRange();
	range.load('address');
	return ctx.sync().then(function() {
			console.log(range.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Getter 

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	namedItem.load('type');
	return ctx.sync().then(function() {
			console.log(namedItem.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
