# TableSort

### apply()
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var filter = table.columns.getItemAt(1).filter;
        filter.apply({
            filterOn: Excel.FilterOn.bottomItems,
            criterion1: "3"
        });
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```