# Shape Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|name|string|Shape's locale specific name. Read-only.|1.1||
|nameID|string|Returns synthetic name of shape (sheet.ID). Locale independent. Read-only.|1.1||
|text|string|Shape's Text. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|hyperlinks|[Hyperlinks](hyperlinks.md)|Returns the Hyperlinks collection for a Shape object. Read-only.|1.1||
|iD|[long](long.md)|Returns shape's ID Read-only.|1.1||
|shapeData|[Section](section.md)|Returns the Shape Data Section. Read-only.|1.1||
|shapeRect|[BoundingBox](boundingbox.md)|Shape's BoundingBox (x,y coordinates) and width & Height Read-only.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[addHighlight(Width: number, Color: String)](#addhighlightwidth-number-color-string)|void|Draws a rectangular highlight around the bounding box of the shape.|1.1|
|[addOverlay(OverlayID: String, OverlayType: string, Content: String, HorizontalPosition: string, VerticalPosition: string, OverlayWidth: number, OverlayHeight: number)](#addoverlayoverlayid-string-overlaytype-string-content-string-horizontalposition-string-verticalposition-string-overlaywidth-number-overlayheight-number)|void|Adds an Overlay on top of the Shape|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[removeHighlight()](#removehighlight)|void|Removes the highlight from the shape, if one exists.|1.1|
|[removeOverlay(OverlayID: String)](#removeoverlayoverlayid-string)|void|Removes particular Overlay or all Overlays on the Shape|1.1|

## Method Details


### addHighlight(Width: number, Color: String)
Draws a rectangular highlight around the bounding box of the shape.

#### Syntax
```js
shapeObject.addHighlight(Width, Color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Width|number|A positive integer that specifies the width of the highlight's stroke in pixels.|
|Color|String|A string that specifies the color of the highlight. It must have the form "#RRGGBB", where each letter represents a hexadecimal digit between 0 and F, and where RR is the red value between 0 and 0xFF (255), GG the green value between 0 and 0xFF (255), and BB is the blue value between 0 and 0xFF (255).|

#### Returns
void

### addOverlay(OverlayID: String, OverlayType: string, Content: String, HorizontalPosition: string, VerticalPosition: string, OverlayWidth: number, OverlayHeight: number)
Adds an Overlay on top of the Shape

#### Syntax
```js
shapeObject.addOverlay(OverlayID, OverlayType, Content, HorizontalPosition, VerticalPosition, OverlayWidth, OverlayHeight);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|OverlayID|String|A string that represents the identifier (ID) of the overlay|
|OverlayType|string|Text or Picture Possible values are: `text` 0,`image` 1|
|Content|String|HTML or location of the Picture|
|HorizontalPosition|string|Optional. Alignment - Left, Center, Right. Default = Left Possible values are: `left` 0,`center` 1,`right` 2|
|VerticalPosition|string|Optional. Alignment - Top, Middle, Bottom. Default = Top Possible values are: `top` 0,`middle` 1,`bottom` 2|
|OverlayWidth|number|Optional. A positive integer that specifies the width of the overlay, in pixels. Default = Shape's width|
|OverlayHeight|number|Optional. A positive integer that specifies the height of the overlay, in pixels. Default = Shape's height|

#### Returns
void

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

### removeHighlight()
Removes the highlight from the shape, if one exists.

#### Syntax
```js
shapeObject.removeHighlight();
```

#### Parameters
None

#### Returns
void

### removeOverlay(OverlayID: String)
Removes particular Overlay or all Overlays on the Shape

#### Syntax
```js
shapeObject.removeOverlay(OverlayID);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|OverlayID|String|Optional. A string that represents the identifier (ID) of the overlay. Default: Removes all overlays on the shape|

#### Returns
void
