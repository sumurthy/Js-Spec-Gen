# item Object (JavaScript API for Outlook)

_Applies to: Outlook Online_
_Note: This API is in preview_

Provides methods and properties for accessing a message or appointment in an Outlook add-in.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|attachments|AttachmentDetails[]|Gets an array of attachments for the item.|
|bcc|Recipients|Gets or sets the recipients on the Bcc (blind carbon copy) line of a message.|
|body|Body|Gets an object that provides methods for manipulating the body of an item.|
|cc|EmailAddressDetails[] Recipients|Gets or sets the Cc (carbon copy) recipients of a message.|
|conversationId|String|Gets an identifier for the email conversation that contains a particular message.|
|dateTimeCreated|Date|Gets the date and time that an item was created.|
|dateTimeModified|Date|Gets the date and time that an item was last modified.|
|end|Date Time|Gets or sets the date and time that the appointment is to end.|
|from|EmailAddressDetails|Gets the email address of the sender of a message.|
|internetMessageId|String|Gets the Internet message identifier for an email message.|
|itemClass|String|Gets the Exchange Web Services item class of the selected item.|
|itemId|String|Gets the Exchange Web Services item identifier for the current item.|
|itemType|Office.MailboxEnums.ItemType|Gets the type of item that an instance represents.|
|location|String Location|Gets or sets the location of an appointment.|
|normalizedSubject|String|Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`).|
|notificationMessages|NotificationMessages|Gets the notification messages for an item.|
|optionalAttendees|EmailAddressDetails[] Recipients|Gets or sets a list of email addresses for optional attendees.|
|organizer|EmailAddressDetails|Gets the email address of the meeting organizer for a specified meeting.|
|requiredAttendees|EmailAddressDetails[] Recipients|Gets or sets a list of email addresses for required attendees.|
|resources|EmailAddressDetails|Gets the resources required for an appointment.|
|sender|EmailAddressDetails|Gets the email address of the sender of an email message.|
|start|Date Time|Gets or sets the date and time that the appointment is to begin.|
|subject|String Subject|Gets or sets the description that appears in the subject field of an item.|
|to|EmailAddressDetails[] Recipients|Gets or sets the recipients of an email message.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[](#)|[](.md)|Adds a file to a message or appointment as an attachment.|
|[](#)|[](.md)|Adds an Exchange item, such as a message, as an attachment to the message or appointment.|
|[](#)|[](.md)|Closes the current item that is being composed.|
|[](#)|[](.md)|Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.|
|[](#)|[](.md)|Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.|
|[](#)|[Entities](entities.md)||
|[](#)|[Array<(String|Contact|MeetingSuggestion|PhoneNumber|TaskSuggestion)>](array<(string|contact|meetingsuggestion|phonenumber|tasksuggestion)>.md)|Gets an array of all the entities of the specified entity type found in the selected item.|
|[](#)|[Array<(String|Contact|MeetingSuggestion|PhoneNumber|TaskSuggestion)>](array<(string|contact|meetingsuggestion|phonenumber|tasksuggestion)>.md)|Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.|
|[](#)|[Object](object.md)|Returns string values in the selected item that match the regular expressions defined in the manifest XML file.|
|[](#)|[String[]](string[].md)|Returns string values in the selected item that match the named regular expression defined in the manifest XML file.|
|[](#)|[String](string.md)|Asynchronously returns selected data from the subject or body of a message.|
|[](#)|[](.md)|Asynchronously loads custom properties for this add-in on the selected item.|
|[](#)|[](.md)|Removes an attachment from a message or appointment.|
|[](#)|[](.md)|Asynchronously saves an item.|
|[](#)|[](.md)|Asynchronously inserts data into the body or subject of a message.|

## Method Details


### 
Adds a file to a message or appointment as an attachment.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|uri|String|The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.|
|attachmentName|String|The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.|
|options|Object|An object literal that contains one or more of the following properties. For more information on defining and using the `options` parameter, see {@tutorial options}.|
|options.asyncContext|Object|Developers can provide any object they wish to access in the callback method.|
|callback|function|On success, the attachment identifier will be provided in the `asyncResult.value` property.|

#### Returns
[](.md)

### 
Adds an Exchange item, such as a message, as an attachment to the message or appointment.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|itemId|String|The Exchange identifier of the item to attach. The maximum length is 100 characters.|
|attachmentName|String|The subject of the item to be attached. The maximum length is 255 characters.|
|options|Object|An object literal that contains one or more of the following properties. For more information on defining and using the `options` parameter, see {@tutorial options}.|
|options.asyncContext|Object|Developers can provide any object they wish to access in the callback method.|
|callback|function|On success, the attachment identifier will be provided in the `asyncResult.value` property.|

#### Returns
[](.md)

### 
Closes the current item that is being composed.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|

#### Returns
[](.md)

### 
Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|formData|String | Object|| Object} A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.|
|formData.htmlBody|String|A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.|
|formData.attachments|Object[]|An array of JSON objects that are either file or item attachments.|
|formData.attachments.type|String|Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.|
|formData.attachments.name|String|A string that contains the name of the attachment, up to 255 characters in length.|
|formData.attachments.url|String|Only used if `type` is set to `file`. The URI of the location for the file.|
|formData.attachments.itemId|String|Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.|
|formData.callback|function||

#### Returns
[](.md)

### 
Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|formData|String|A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.|
|formData|Object|An object that contains body or attachment data and a callback function. The object is defined as follows:|

#### Returns
[](.md)

### 


#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|

#### Returns
[Entities](entities.md)

### 
Gets an array of all the entities of the specified entity type found in the selected item.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|entityType|Office.MailboxEnums.EntityType|One of the EntityType enumeration values.|

#### Returns
[Array<(String|Contact|MeetingSuggestion|PhoneNumber|TaskSuggestion)>](array<(string|contact|meetingsuggestion|phonenumber|tasksuggestion)>.md)

### 
Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|String|The name of the `ItemHasKnownEntity` rule element that defines the filter to match.|

#### Returns
[Array<(String|Contact|MeetingSuggestion|PhoneNumber|TaskSuggestion)>](array<(string|contact|meetingsuggestion|phonenumber|tasksuggestion)>.md)

### 
Returns string values in the selected item that match the regular expressions defined in the manifest XML file.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|

#### Returns
[Object](object.md)

### 
Returns string values in the selected item that match the named regular expression defined in the manifest XML file.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|String|The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.|

#### Returns
[String[]](string[].md)

### 
Asynchronously returns selected data from the subject or body of a message.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|coercionType|Office.CoercionType|Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.|
|options|Object|An object literal that contains one or more of the following properties. For more information on defining and using the `options` parameter, see {@tutorial options}.|
|options.asyncContext|Object|Developers can provide any object they wish to access in the callback method.|
|callback|function|To access the selected data from the callback method, call `asyncResult.value.data`. To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.|

#### Returns
[String](string.md)

### 
Asynchronously loads custom properties for this add-in on the selected item.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|callback|function|The custom properties are provided as a [CustomProperties]{@linkcode CustomProperties} object in the `asyncResult.value` property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.|
|userContext|Object|Developers can provide any object they wish to access in the callback function. This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|

#### Returns
[](.md)

### 
Removes an attachment from a message or appointment.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|attachmentId|String|The identifier of the attachment to remove. The maximum length of the string is 100 characters.|
|options|Object|An object literal that contains one or more of the following properties. For more information on defining and using the `options` parameter, see {@tutorial options}.|
|options.asyncContext|Object|Developers can provide any object they wish to access in the callback method.|
|callback|function|If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.|

#### Returns
[](.md)

### 
Asynchronously saves an item.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|options|Object|An object literal that contains one or more of the following properties. For more information on defining and using the `options` parameter, see {@tutorial options}.|
|options.asyncContext|Object|Developers can provide any object they wish to access in the callback method.|
|callback|function|On success, the item identifier is provided in the `asyncResult.value` property.|

#### Returns
[](.md)

### 
Asynchronously inserts data into the body or subject of a message.

#### Syntax
```js

```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|data|String|The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.|
|options|Object|An object literal that contains one or more of the following properties. For more information on defining and using the `options` parameter, see {@tutorial options}.|
|options.asyncContext|Object|Developers can provide any object they wish to access in the callback method.|
|options.coercionType|Office.CoercionType|If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.|
|callback|function||

#### Returns
[](.md)
