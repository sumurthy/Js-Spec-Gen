# Outline Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a container for Paragraph objects.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the Outline object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-id)|

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|pageContent|[PageContent](pagecontent.md)|Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-pageContent)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of Paragraph objects in the Outline. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-paragraphs)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[append(html: string)](#appendhtml-string)|void|Adds the specified HTML to the bottom of the Outline.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-append)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-load)|
|[prepend(html: string)](#prependhtml-string)|void|Adds the specified HTML to the top of the Outline.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-prepend)|

## Method Details


### append(html: string)
Adds the specified HTML to the bottom of the Outline.

#### Syntax
```js
outlineObject.append(html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|The HTML to append.|

#### Returns
void

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.activePage;

    // Get pageContents of the activePage. 
    var pageContents = activePage.getContents();

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.append("<p>new paragraph</p>");

                // Run the queued commands.
                return context.sync();
            }
        })
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```


### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### prepend(html: string)
Adds the specified HTML to the top of the Outline.

#### Syntax
```js
outlineObject.prepend(html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|The HTML to insert.|

#### Returns
void

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.activePage;

    // Get pageContents of the activePage. 
    var pageContents = activePage.getContents();

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to prepend a paragraph to the outline.
                outline.prepend("<p>new paragraph</p>");

                // Run the queued commands.
                return context.sync();
            }
        })
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```