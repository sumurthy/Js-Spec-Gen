# Image Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|description|string|Gets or sets the description of the Image.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-description)|
|height|double|Gets or sets the height of the Image layout.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-height)|
|hyperlink|string|Gets or sets the hyperlink of the Image.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-hyperlink)|
|id|string|Gets the ID of the Image object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-id)|
|width|double|Gets or sets the width of the Image layout.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-width)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|pageContent|[PageContent](pagecontent.md)|Gets the PageContent object that contains the Image. Returns null if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-pageContent)|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the Image. Returns null if the Image is not a direct child of a Paragraph. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-paragraph)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[getBase64Image()](#getbase64image)|string|Gets the base64-encoded binary representation of the Image.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-getBase64Image)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-load)|

## Method Details


### getBase64Image()
Gets the base64-encoded binary representation of the Image.

#### Syntax
```js
imageObject.getBase64Image();
```

#### Parameters
None

#### Returns
string

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
