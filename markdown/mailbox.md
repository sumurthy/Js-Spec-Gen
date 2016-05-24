# mailbox resource type

##### Namespace: *Office.context.mailbox*

Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.

Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.

|Requirement| Value|
|:----------|:-----|
|Minimum requirement set/version: | *1.0*|
|Minimum permission level |*Restricted* |
|Modes supported | *Read, Compose*|


### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
|ewsUrl      | String | Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. | 1.0  %readonly%|  



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [convertToLocalClientTime](converttolocalclienttime)     | [LocalClientTime](localclienttime.md) | Gets a dictionary containing time information in local client time.  | 1.0|  
| [convertToUtcClientTime](converttoutcclienttime)     | Date | Gets a Date object from a dictionary containing time information.  | 1.0|  
| [displayAppointmentForm](displayappointmentform)     |  | Displays an existing calendar appointment.  | 1.0|  
| [displayMessageForm](displaymessageform)     |  | Displays an existing message.  | 1.0|  
| [displayNewAppointmentForm](displaynewappointmentform)     |  | Displays a form for creating a new calendar appointment.  | 1.0|  
| [getCallbackTokenAsync](getcallbacktokenasync)     |  | Gets a string that contains a token used to get an attachment or item from an Exchange Server.  | 1.0|  
| [getUserIdentityTokenAsync](getuseridentitytokenasync)     |  | Gets a token identifying the user and the Office Add-in.  | 1.0|  
| [makeEwsRequestAsync](makeewsrequestasync)     |  | Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.  | 1.0|  

