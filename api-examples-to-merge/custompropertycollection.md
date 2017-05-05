### getNextOrNullObject(key: string)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the custom property.
    var customProperty = context.document.properties.customProperties.getItemOrNullObject('MyProperty');

    // Queue a command to load the value of the custom property.
    context.load(customProperty, 'value');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (customProperty.isNullObject) {
            console.log('There is no property with that name.')
        } else {
            console.log('The property has been found in this document.');
        }
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```