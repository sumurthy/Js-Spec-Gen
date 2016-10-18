# Shape Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents the Shape class.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|name|string|Shape's name. Read-only.|1.1||
|text|string|Shape's Text. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|hyperlinks|[HyperlinkCollection](hyperlinkcollection.md)|Returns the Hyperlinks collection for a Shape object. Read-only.|1.1||
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|Returns the Shape's Data Section. Read-only.|1.1||
|view|[ShapeView](shapeview.md)|Returns the view of the shape. Read-only.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


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
