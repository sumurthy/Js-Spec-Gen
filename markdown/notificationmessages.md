# NotificationMessages resource type

Provides methods for accessing notification messages in an Outlook add-in.

The `NotificationMessages` object is returned as the [notificationMessages]{@linkcode Office.context.mailbox.item#notificationMessages} property of an item.

*	Namespace: *NotificationMessages*
*	Minimum requirement set/version: *1.3*
*	Minimum permission level: *ReadItem*
*	Modes supported: *Read, Compose*



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [addAsync](addasync)     |  | Adds a notification to an item. | 1.3|  
| [getAllAsync](getallasync)     |  | Returns all keys and messages for an item. | 1.3|  
| [removeAsync](removeasync)     |  | Removes a notification message for an item. | 1.3|  
| [replaceAsync](replaceasync)     |  | Replaces a notification message that has a given key with another message. | 1.3|  
>| [%name%](%link%)     | %type% | %description% | %req%|

