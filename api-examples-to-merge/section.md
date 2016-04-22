
### getFooter(type: HeaderFooterType)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
	
	// Create a proxy sectionsCollection object.
	var mySections = context.document.sections;
	
	// Queue a commmand to load the sections.
	context.load(mySections, 'body/style');
	
	// Synchronize the document state by executing the queued commands, 
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		
		// Create a proxy object the primary footer of the first section. 
		// Note that the footer is a body object.
		var myFooter = mySections.items[0].getFooter("primary");
		
		// Queue a command to insert text at the end of the footer.
		myFooter.insertText("This is a footer.", Word.InsertLocation.end);
		
		// Queue a command to wrap the header in a content control.
		myFooter.insertContentControl();
							  
		// Synchronize the document state by executing the queued commands, 
		// and return a promise to indicate task completion.
		return context.sync().then(function () {
			console.log("Added a footer to the first section.");
		});                    
	});  
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
	}
});
```
### getHeader(type: HeaderFooterType)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary header of the first section. 
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");
        
        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
        });                    
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
