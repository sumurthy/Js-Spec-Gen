# RichText Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents a RichText object in a Paragraph.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the RichText object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-id)|
|text|string|Gets the text content of the RichText object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-text)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the RichText object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-paragraph)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)|void|Inserts HTML at the specified location in the RichText object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-insertHtml)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-load)|

## Method Details


### insertHtml(html: string, insertLocation: string)
Inserts HTML at the specified location in the RichText object.

#### Syntax
```js
richTextObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|The HTML to insert.|
|insertLocation|string|The location to insert the HTML.  Possible values are: Before, After|

#### Returns
void

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
