# Hyperlink Object (JavaScript API for Visio)

_Visio 2016, Visio for iPad, Visio for Mac_

Dispatch Ids

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Gets the address for a shape's Hyperlink object, the address to which the hyperlink navigates. Read-only.|1.1||
|description|string|Gets the description of a hyperlink. Read-only.|1.1||
|name|string|Hyperlinks's locale specific name. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

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
