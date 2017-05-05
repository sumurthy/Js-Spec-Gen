
### getById(id: number)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

	// Create a proxy object for the content control that contains a specific id.
	var contentControl = context.document.contentControls.getById(30086310);

	// Queue a command to load the text property for a content control.
	context.load(contentControl, 'text');

	// Synchronize the document state by executing the queued commands,
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		console.log('The content control with that Id has been found in this document.');
	});
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
	}
});
```

### getByIdOrNullObject(id: number)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

	// Create a proxy object for the content control that contains a specific id.
    var contentControl = context.document.contentControls.getByIdOrNullObject(30086310);

    // Queue a command to load the text property for a content control.
    context.load(contentControl, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControl.isNullObject) {
            console.log('There is no content control with that ID.')
        } else {
            console.log('The content control with that ID has been found in this document.');
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


### getByTag(tag: string)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');

    // Queue a command to load the text property for all of content controls with a specific tag.
    context.load(contentControlsWithTag, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log("There isn't a content control with a tag of Customer-Address in this document.");
        } else {
            console.log('The first content control with the tag of Customer-Address has this text: ' + contentControlsWithTag.items[0].text);
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

*Additional information*
The [Word-Add-in-DocumentAssembly][contentControls.getByTag] sample has another example of using the getByTag method.


### getByTitle(title: string)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific title.
    var contentControlsWithTitle = context.document.contentControls.getByTitle('Enter Customer Address Here');

    // Queue a command to load the text property for all of content controls with a specific title.
    context.load(contentControlsWithTitle, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTitle.items.length === 0) {
            console.log("There isn't a content control with a title of 'Enter Customer Address Here' in this document.");
        } else {
            console.log("The first content control with the title of 'Enter Customer Address Here' has this text: " + contentControlsWithTitle.items[0].text);
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

*Additional information*

The [Word-Add-in-DocumentAssembly][contentControls.getByTitle] sample has another example of using the getByTitle method.

### getFirstOrnullObject()

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the first content control in the document.
	var contentControl = context.document.contentControls.getFirstOrNullObject();

    // Queue a command to load the text property for a content control.
    context.load(contentControl, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControl.isNullObject) {
            console.log('There are no content controls in this document.')
        } else {
            console.log('The first content control has been found in this document.');
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


### load(param: object)

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    // Queue a command to load the id property for all of the content controls.
    context.load(contentControls, 'id');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control.
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' +
                        '   ----- appearance: ' + contentControls.items[0].appearance +
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
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

The [Silly stories](https://aka.ms/sillystorywordaddin) add-in sample shows how the **load** method is used to load the content control collection with the **tag** and **title** properties.
