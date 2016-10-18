# ShapeDataItemCollection Object (JavaScript API for Visio)

_Visio Online_

Represents the ShapeDataItemCollection for a given Shape.

## Properties

| Property	   | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|items|[ShapeDataItem[]](shapedataitem.md)|A collection of shapeDataItem objects. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-items)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|[getCount()](#getcount)|int|Gets the number of Shape Data Items.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-getCount)|
|[getItem(key: string)](#getitemkey-string)|[ShapeDataItem](shapedataitem.md)|Gets the ShapeDataItem using its name.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-load)|

## Method Details


### getCount()
Gets the number of Shape Data Items.

#### Syntax
```js
shapeDataItemCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: string)
Gets the ShapeDataItem using its name.

#### Syntax
```js
shapeDataItemCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Key is the name of the ShapeDataItem to be retrieved.|

#### Returns
[ShapeDataItem](shapedataitem.md)

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
