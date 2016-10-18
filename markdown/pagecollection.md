# PageCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a collection of Page objects that are part of the document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Page[]](page.md)|A collection of page objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of pages in the collection.|1.1|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Page](page.md)|Gets a page using its key (name or Id).|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getCount()
Gets the number of pages in the collection.

#### Syntax
```js
pageCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: number or string)
Gets a page using its key (name or Id).

#### Syntax
```js
pageCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|number or string|Key is the name or Id of the page to be retrieved.|

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
