# mailbox resource type

Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.

Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.

*	Namespace: *Office.context.mailbox*
*	Minimum requirement set/version: *1.0*
*	Minimum permission level: *Restricted*
*	Modes supported: *Read, Compose*


### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
|ewsUrl      | String | Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. | 1.0 |  
>|%name%      | %type% | %description% | %req% |



## Methods

| Method	   | Return Type    | Description | Requirements|
|:-------------|:---------------|:------------|:----|
| [convertToLocalClientTime](converttolocalclienttime)     | %dtype% | Gets a dictionary containing time information in local client time. | 1.0|  
| [convertToUtcClientTime](converttoutcclienttime)     | %dtype% | Gets a Date object from a dictionary containing time information. | 1.0|  
| [displayAppointmentForm](displayappointmentform)     | %dtype% | Displays an existing calendar appointment. | 1.0|  
| [displayMessageForm](displaymessageform)     | %dtype% | Displays an existing message. | 1.0|  
| [displayNewAppointmentForm](displaynewappointmentform)     | %dtype% | Displays a form for creating a new calendar appointment. | 1.0|  
| [getCallbackTokenAsync](getcallbacktokenasync)     | %dtype% | Gets a string that contains a token used to get an attachment or item from an Exchange Server. | 1.0|  
| [getUserIdentityTokenAsync](getuseridentitytokenasync)     | %dtype% | Gets a token identifying the user and the Office Add-in. | 1.0|  
| [makeEwsRequestAsync](makeewsrequestasync)     | %dtype% | Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox. | 1.0|  
>| [%name%](%link%)     | %dtype% | %description% | %req%|

