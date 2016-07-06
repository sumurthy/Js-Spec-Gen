# Cell Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|formula|string|Cell's formula in locale specific syntax. Read-only.|1.1||
|formulaU|string|Cell's formula in locale independent syntax. Mohan: Good for editing scenario Read-only.|1.1||
|localName|string|Cell's name (Column Name) in locale specific syntax. Read-only.|1.1||
|name|string|Cell's name (Column Name) in locale independent syntax. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|column|[short](short.md)|Column index Read-only.|1.1||
|units|[short](short.md)|Indicates the unit of measure associated with a Cell object. Read-only.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[result(UnitsNameOrCode: VisUnitCodes or string)](#resultunitsnameorcode-visunitcodes-or-string)|[Double](double.md)|Gets the value of a ShapeSheet cell expressed as a double. Read-only.|1.1|
|[resultStr(UnitsNameOrCode: VisUnitCodes or string)](#resultstrunitsnameorcode-visunitcodes-or-string)|[String](string.md)|Gets the value of a ShapeSheet cell expressed as a string. Read-only.|1.1|

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

### result(UnitsNameOrCode: VisUnitCodes or string)
Gets the value of a ShapeSheet cell expressed as a double. Read-only.

#### Syntax
```js
cellObject.result(UnitsNameOrCode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|UnitsNameOrCode|VisUnitCodes or string|The units to use when retrieving the value.|

#### Returns
[Double](double.md)

### resultStr(UnitsNameOrCode: VisUnitCodes or string)
Gets the value of a ShapeSheet cell expressed as a string. Read-only.

#### Syntax
```js
cellObject.resultStr(UnitsNameOrCode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|UnitsNameOrCode|VisUnitCodes or string|The units to use when retrieving the value.|

#### Returns
[String](string.md)
