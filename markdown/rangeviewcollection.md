# RangeViewCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a collection of worksheet objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[RangeView[]](rangeview.md)|A collection of rangeView objects. Read-only.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|Gets a RangeView Row via it's index. Zero-Indexed.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getItemAt(index: number)
Gets a RangeView Row via it's index. Zero-Indexed.

#### Syntax
```js
rangeViewCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index of the visible row.|

#### Returns
[RangeView](rangeview.md)

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
