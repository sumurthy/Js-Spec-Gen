# HyperlinkCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents the Hyperlink Collection.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Hyperlink[]](hyperlink.md)|A collection of hyperlink objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of hyperlinks.|1.1|
|[getItem(Key: number or string)](#getitemkey-number-or-string)|[Hyperlink](hyperlink.md)|Gets a Hyperlink using its key (name or Id).|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getCount()
Gets the number of hyperlinks.

#### Syntax
```js
hyperlinkCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(Key: number or string)
Gets a Hyperlink using its key (name or Id).

#### Syntax
```js
hyperlinkCollectionObject.getItem(Key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Key|number or string|Key is the name or index of the Hyperlink to be retrieved.|

#### Returns
[Hyperlink](hyperlink.md)

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
