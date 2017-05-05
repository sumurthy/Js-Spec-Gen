### getNextOrNullObject()

To use this snippet, add an inline picture to the document and assign it an alt text title.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the first inline picture.
	var firstPicture = context.document.body.inlinePictures.getFirstOrNullObject();

    // Queue a command to load the alternative text title of the picture.
    context.load(firstPicture, 'altTextTitle');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (firstPicture.isNullObject) {
            console.log('There are inline pictures in this document.')
        } else {
            console.log(firstPicture.altTextTitle);
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