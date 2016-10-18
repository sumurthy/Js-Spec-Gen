# Page Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents the Page class.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|index|int|Index of the Page. Read-only.|1.1||
|isBackground|bool|Whether the page is a background page or not. Read-only.|1.1||
|name|string|Page name. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|shapes|[ShapeCollection](shapecollection.md)|Shapes in the Page. Read-only.|1.1||
|view|[PageView](pageview.md)|Returns the view of the page. Read-only.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[activate()](#activate)|void|Set the page as Active Page of the document.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### activate()
Set the page as Active Page of the document.

#### Syntax
```js
pageObject.activate();
```

#### Parameters
None

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
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
