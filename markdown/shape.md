# Shape Object (JavaScript API for Visio)

_Visio Online_

Represents the Shape class.

## Properties

| Property	   | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|name|string|Shape's name. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-name)|
|text|string|Shape's Text. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-text)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|hyperlinks|[HyperlinkCollection](hyperlinkcollection.md)|Returns the Hyperlinks collection for a Shape object. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-hyperlinks)|
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|Returns the Shape's Data Section. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-shapeDataItems)|
|view|[ShapeView](shapeview.md)|Returns the view of the shape. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-view)|

## Methods

| Method		   | Return Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-load)|

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
