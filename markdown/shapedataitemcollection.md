# ShapeDataItemCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents the ShapeDataItemCollection for a given Shape.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[ShapeDataItem[]](shapedataitem.md)|A collection of shapeDataItem objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of Shape Data Items.|1.1|
|[getItem(key: string)](#getitemkey-string)|[ShapeDataItem](shapedataitem.md)|Gets the ShapeDataItem using its name.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

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
