# Shapes Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Represents the no. of shapes. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(NameOrIndex: number or string)](#getitemnameorindex-number-or-string)|[Shape](shape.md)|Gets a Shape using its name or ID.|1.1|
|[item(NameOrIndex: number or string)](#itemnameorindex-number-or-string)|[Shape](shape.md)|Retrieves the Shape by Name Or Index. Mohan: can we just have over Index.|1.1|
|[itemFromID(ObjectID: long)](#itemfromidobjectid-long)|[Shape](shape.md)|Retrieves the Shape by ID.|1.1|
|[itemU(NameOrIndex: number or string)](#itemunameorindex-number-or-string)|[Shape](shape.md)|Retrieves the Shape by Name (Locale Independent) Or Index|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getItem(NameOrIndex: number or string)
Gets a Shape using its name or ID.

#### Syntax
```js
shapesObject.getItem(NameOrIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|NameOrIndex|number or string|Name or Index of the page to be retrieved.|

#### Returns
[Shape](shape.md)

### item(NameOrIndex: number or string)
Retrieves the Shape by Name Or Index. Mohan: can we just have over Index.

#### Syntax
```js
shapesObject.item(NameOrIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|NameOrIndex|number or string|Name or Index|

#### Returns
[Shape](shape.md)

### itemFromID(ObjectID: long)
Retrieves the Shape by ID.

#### Syntax
```js
shapesObject.itemFromID(ObjectID);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|ObjectID|long|ID|

#### Returns
[Shape](shape.md)

### itemU(NameOrIndex: number or string)
Retrieves the Shape by Name (Locale Independent) Or Index

#### Syntax
```js
shapesObject.itemU(NameOrIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|NameOrIndex|number or string|Name or Index|

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
