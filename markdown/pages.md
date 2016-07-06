# Pages Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Represents the no. of pages. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(id: number or string)](#getitemid-number-or-string)|[Page](page.md)|Gets a page using its ID.|1.1|
|[item(Index: number)](#itemindex-number)|[Page](page.md)|Retrieves the Page by Index|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getItem(id: number or string)
Gets a page using its ID.

#### Syntax
```js
pagesObject.getItem(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|id|number or string|Name or ID of the page to be retrieved.|

#### Returns
[Page](page.md)

### item(Index: number)
Retrieves the Page by Index

#### Syntax
```js
pagesObject.item(Index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Index|number|Index of the Page. The first item in a Pages collection is item 1|

#### Returns
[Page](page.md)

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
