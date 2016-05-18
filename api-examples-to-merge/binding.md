
### getRange()
Below example uses binding object to get the associated range.

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var range = binding.getRange();
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getTable()
```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var table = binding.getTable();
	table.load('name');
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getText()

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var text = binding.getText();
	ctx.load('text');
	return ctx.sync().then(function() {
		console.log(text);
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
	var binding = ctx.workbook.bindings.getItemAt(0);
	binding.load('type');
	return ctx.sync().then(function() {
		console.log(binding.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
