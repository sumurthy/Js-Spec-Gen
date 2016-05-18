# Table Column
### delete() 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
	column.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getDataBodyRange() 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var dataBodyRange = column.getDataBodyRange();
	dataBodyRange.load('address');
	return ctx.sync().then(function() {
		console.log(dataBodyRange.address);
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getHeaderRowRange()

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var headerRowRange = columns.getHeaderRowRange();
	headerRowRange.load('address');
	return ctx.sync().then(function() {
		console.log(headerRowRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getRange() 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var columnRange = columns.getRange();
	columnRange.load('address');
	return ctx.sync().then(function() {
		console.log(columnRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getTotalRowRange() 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var totalRowRange = columns.getTotalRowRange();
	totalRowRange.load('address');
	return ctx.sync().then(function() {
		console.log(totalRowRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Getter and setter

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItem(0);
	column.load('index');
	return ctx.sync().then(function() {
		console.log(column.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
	column.values = newValues;
	column.load('values');
	return ctx.sync().then(function() {
		console.log(column.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```