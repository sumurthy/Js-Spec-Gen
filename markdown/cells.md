# Cells Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Represents the no. of cells. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(NameOrIndex: number or string)](#getitemnameorindex-number-or-string)|[Cell](cell.md)|Gets a Cell using its name or Index.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getItem(NameOrIndex: number or string)
Gets a Cell using its name or Index.

#### Syntax
```js
cellsObject.getItem(NameOrIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|NameOrIndex|number or string|Name or Index of the page to be retrieved.|

#### Returns
[Cell](cell.md)

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
