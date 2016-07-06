# Selection Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|shapes|[Shapes](shapes.md)|Gets the Shapes of the Selection Read-only.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[select(Shape: Shape, SelectAction: string)](#selectshape-shape-selectaction-string)|void|Selects or clears the selection of an object. Mohan: should this be Add.|1.1|

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

### select(Shape: Shape, SelectAction: string)
Selects or clears the selection of an object. Mohan: should this be Add.

#### Syntax
```js
selectionObject.select(Shape, SelectAction);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Shape|Shape|Shape|
|SelectAction|string|Type of Selection Possible values are: `visDeselect` 1,`visSelect` 2|

#### Returns
void
