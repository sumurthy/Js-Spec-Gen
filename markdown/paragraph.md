# Paragraph Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content, such as RichText, Image, or Table.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the Paragraph object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-id)|
|type|string|Gets the type of the Paragraph object. Read-only. Possible values are: RichText, Image, Other.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-type)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|Gets the Image object in the Paragraph. Returns null if ParagraphType is not Image. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-image)|
|outline|[Outline](outline.md)|Gets the Outline object that contains the Paragraph. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-outline)|
|parentParagraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the Paragraph. Returns null if the Paragraph is a direct child of an Outline. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraph)|
|richText|[RichText](richtext.md)|Gets the RichText object in the Paragraph. Returns null if ParagraphType is not RichText. Read-only Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-richText)|
|subParagraphs|[ParagraphCollection](paragraphcollection.md)|Gets the child Paragraph objects of the Paragraph. Applies only if ParagraphType is Table. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-subParagraphs)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-load)|

## Method Details


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
### Property access examples

**id and type**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.activePage.getContents();
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                console.log("Paragraph type: " + paragraph.type);
                console.log("Paragraph ID: " + paragraph.id);
            });
        })                
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }); 
    });
```