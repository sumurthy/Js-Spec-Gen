# context resource type

Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.

The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [documentation on MSDN](https://msdn.microsoft.com/EN-US/library/office/fp161104.aspx).

*	Namespace: *Office.context*
*	Minimum requirement set/version: *1.0*
*	Minimum permission level: **
*	Modes supported: *Read, Compose*


### Properties

| Property	   | Type	| Description| Requirements|
|:-------------|:-------|:-----------|:------------|
|displayLanguage      | String | Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application. | 1.0 |  
|officeTheme      | Object | Provides access to the properties for Office theme colors. | 1.3 |  
|roamingSettings      | RoamingSettings | Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox. | 1.0 |  
>|%name%      | %type% | %description% | %req% |


