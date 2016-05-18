# CustomProperties resource type

Provides methods for accessing item-specific custom properties in an Outlook add-in.

The `CustomProperties` object represents custom properties that are specific to a particular item and specific to a mail add-in for Outlook. For example, there might be a need for a mail add-in to save some data that is specific to the current email message that activated the add-in. If the user revisits the same message in the future and activates the mail add-in again, the add-in will be able to retrieve the data that had been saved as custom properties. 	 
 	 
Because Outlook for Mac doesn’t cache custom properties, if the user’s network goes down, mail add-ins cannot access their custom properties. 	 
Office.initialize = function () { 	 
// Checks for the DOM to load using the jQuery ready function. 	 
$(document).ready(function () { 	 
// After the DOM is loaded, add-in-specific code can run. 	 
var mailbox = Office.context.mailbox; 	 
mailbox.item.loadCustomPropertiesAsync(customPropsCallback); 	 
}); 	 
} 	 
function customPropsCallback(asyncResult) { 	 
var customProps = asyncResult.value; 	 
var myProp = customProps.get("myProp"); 	 
 	 
customProps.set("otherProp", "value"); 	 
customProps.saveAsync(saveCallback); 	 
} 	 
 	 
function saveCallback(asyncResult) { 	 
} 	 
##### Example 
 	 

```js 	 
The following example shows how to use the [`loadCustomPropertiesAsync`]{@link Office.context.mailbox.Item#loadCustomPropertiesAsync} method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the [`saveAsync`]{@link CustomProperties#saveAsync} method to save these properties back to the server. After loading the custom properties, the example uses the [`get`]{@link CustomProperties#get} method to read the custom property `myProp`, the [`set`]{@link CustomProperties#set} method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties. 	 
Office.initialize = function () { 	 
// Checks for the DOM to load using the jQuery ready function. 	 
$(document).ready(function () { 	 
// After the DOM is loaded, add-in-specific code can run. 	 
var mailbox = Office.context.mailbox; 	 
mailbox.item.loadCustomPropertiesAsync(customPropsCallback); 	 
}); 	 
} 	 
function customPropsCallback(asyncResult) { 	 
var customProps = asyncResult.value; 	 
var myProp = customProps.get("myProp"); 	 
 	 
customProps.set("otherProp", "value"); 	 
customProps.saveAsync(saveCallback); 	 
} 	 
 	 
function saveCallback(asyncResult) { 	 
} 	 
```


*	Namespace: *CustomProperties*
*	Minimum requirement set/version: *1.0*
*	Minimum permission level: *ReadItem*
*	Modes supported: *Read, Compose*



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [get](get)     | String | Returns the value of the specified custom property. | 1.0|  
| [remove](remove)     |  | Removes the specified property from the custom property collection. | 1.0|  
| [saveAsync](saveasync)     |  | Saves item-specific custom properties to the server. | 1.0|  
| [set](set)     |  | Sets the specified property to the specified value. | 1.0|  
>| [%name%](%link%)     | %type% | %description% | %req%|

