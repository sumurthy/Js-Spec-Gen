# ShapeCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents the Shape Collection.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Shape[]](shape.md)|A collection of shape objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of Shapes in the collection.|1.1|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Shape](shape.md)|Gets a Shape using its key (name or Index).|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getCount()
Gets the number of Shapes in the collection.

#### Syntax
```js
shapeCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: number or string)
Gets a Shape using its key (name or Index).

#### Syntax
```js
shapeCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|number or string|Key is the Name or Index of the shape to be retrieved.|

#### Returns
[Shape](shape.md)

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
