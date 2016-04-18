@+header
%displayname%
%description%
@-header
@+prerequisites
### Prerequisites
One of the following scopes are required to execute this API:
@-prerequisites
@+httprequest
### HTTP Request
%httprequest%
@-httprequest
@+requestparameter
### Request parameter
In the request URL, provide following query parameters with values.

| Parameter	   | Type	| Description|
|:-------------|:-------|:-----------|
>r|%name%        | %type% | %description% |

@-requestparameter
@+optionalheader
### Optional request headers

| Name         | Value      |
|:-------------|:-----------|
>r| %name%   | %value% |

@-optionalheader
@+requestbody
### Request body
%requestbodymessage%

| Property	   | Type	| Description|
|:-------------|:-------|:-----------|
>r|%name%        | %type% | %description% |

@-requestbody
@+response
### Request body
%responsemessage%
@-response
@+example
### Example
##### Request
Here is an example of the request.
```http
%requesturl%
%reqheaders%

%reqbody%
```
##### Response
Here is an example of the response. 
%truncatenote%
%responseexample%
```http
HTTP/1.1 %status%
%rspheaders%

%rspbody%
```
@-example