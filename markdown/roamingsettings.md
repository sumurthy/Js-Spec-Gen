# RoamingSettings resource type

Provides methods for accessing custom settings for an Outlook add-in.

The settings created by using the methods of the `RoamingSettings` object are saved per add-in and per user. That is, they are available only to the add-in that created them, and only from the user's mail box in which they are saved. 	 
 	 
> While the Outlook Add-in API limits access to these settings to only the add-in that created them, these settings should not be considered secure storage. They can be accessed by Exchange Web Services or Extended MAPI. They should not be used to store sensitive information such as user credentials or security tokens. 	 
 	 
The name of a setting is a String, while the value can be a String, Number, Boolean, null, Object, or Array. 	 
 	 
The `RoamingSettings` object is accessible via the [roamingSettings]{@linkcode Office.context#roamingSettings} property in the `Office.context` namespace. 	 
var value = Office.context.roamingSettings.get('myKey'); 	 
// Update the value of the 'myKey' setting 	 
Office.context.roamingSettings.set('myKey', 'Hello World!'); 	 
// Persist the change 	 
Office.context.roamingSettings.saveAsync(); 	 
##### Example 
 	 
// Get the current value of the 'myKey' setting 	 

```js 	 
var value = Office.context.roamingSettings.get('myKey'); 	 
// Update the value of the 'myKey' setting 	 
Office.context.roamingSettings.set('myKey', 'Hello World!'); 	 
// Persist the change 	 
Office.context.roamingSettings.saveAsync(); 	 
```


*	Namespace: *RoamingSettings*
*	Minimum requirement set/version: *1.0*
*	Minimum permission level: *Restricted*
*	Modes supported: *Read, Compose*



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [get](get)     | (String|Number|Boolean|Object|Array) | Retrieves the specified setting. | 1.0|  
| [remove](remove)     |  | Removes the specified setting. | 1.0|  
| [saveAsync](saveasync)     |  | Saves the settings. | 1.0|  
| [set](set)     |  | Sets or creates the specified setting. | 1.0|  
>| [%name%](%link%)     | %type% | %description% | %req%|

