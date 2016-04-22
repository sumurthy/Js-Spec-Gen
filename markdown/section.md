# Section Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a section in a Word document.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|body|[Body](body.md)|Gets the body object of the section. This does not include the headerfooter and other section metadata. Read-only.|
|next|[Section](section.md)|Gets the next section. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getFooter(type: string)](#getfootertype-string)|[Body](body.md)|Gets one of the section's footers.|
|[getHeader(type: string)](#getheadertype-string)|[Body](body.md)|Gets one of the section's headers.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getFooter(type: string)
Gets one of the section's footers.

#### Syntax
```js
sectionObject.getFooter(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|string|Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'. Possible values are: `Primary` Returns the header or footer on all pages of a section, with the first page or odd pages excluded if they are different.,`FirstPage` Returns the header or footer on the first page of a section.,`EvenPages` Returns all headers or footers on even-numbered pages of a section.|

#### Returns
[Body](body.md)

### getHeader(type: string)
Gets one of the section's headers.

#### Syntax
```js
sectionObject.getHeader(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|string|Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'. Possible values are: `Primary` Returns the header or footer on all pages of a section, with the first page or odd pages excluded if they are different.,`FirstPage` Returns the header or footer on the first page of a section.,`EvenPages` Returns all headers or footers on even-numbered pages of a section.|

#### Returns
[Body](body.md)

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
