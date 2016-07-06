# Page Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|index|int|Index of the Page. Read-only.|1.1||
|nameU|string|Page's locale independent name. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|iD|[long](long.md)|ID of the Page. Read-only.|1.1||
|pageRect|[BoundingBox](boundingbox.md)|Page's X,Y coordinates and Width & Height Read-only.|1.1||
|selection|[Selection](selection.md)|Represents the Selection Read-only.|1.1||
|shapes|[Shapes](shapes.md)|Shapes in the Page. Read-only.|1.1||
|viewRect|[BoundingBox](boundingbox.md)|Page's View Bounding Box|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[centerViewOnShape(Shape: Shape)](#centerviewonshapeshape-shape)|void|Pans the Visio drawing to place the specified shape in the center of the view. Mohan: Pass shapeID to minimize packet traffic|1.1|
|[isShapeInView(Shape: Shape)](#isshapeinviewshape-shape)|bool|To check if the shape is in view of the page or not|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### centerViewOnShape(Shape: Shape)
Pans the Visio drawing to place the specified shape in the center of the view. Mohan: Pass shapeID to minimize packet traffic

#### Syntax
```js
pageObject.centerViewOnShape(Shape);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Shape|Shape|Shape to be seen in the center|

#### Returns
void

### isShapeInView(Shape: Shape)
To check if the shape is in view of the page or not

#### Syntax
```js
pageObject.isShapeInView(Shape);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Shape|Shape|Shape to be checked|

#### Returns
bool

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
