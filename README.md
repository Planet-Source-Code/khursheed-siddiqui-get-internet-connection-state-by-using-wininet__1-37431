<div align="center">

## Get Internet Connection State by Using Wininet


</div>

### Description

Ever need to find out if you are connected to the internet through your application and need to know what type of connection do you have. Well in this article will use InternetGetConnectedState Function in Library Wininet.dll
 
### More Info
 
Boolean and Connection Type


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Khursheed\_Siddiqui](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/khursheed-siddiqui.md)
**Level**          |Intermediate
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/khursheed-siddiqui-get-internet-connection-state-by-using-wininet__1-37431/archive/master.zip)

### API Declarations

```
Private Declare Function InternetGetConnectedState Lib "Wininet" _
 (ByRef dwflags As Long, ByVal dwreserved As Long) As Long
```


### Source Code

```
Option Explicit
Private Declare Function InternetGetConnectedState Lib "Wininet" _
 (ByRef dwflags As Long, ByVal dwreserved As Long) As Long
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
Private Const INTERNET_CONNECTION_LAN As Long = &H2
Private Const INTERNET_CONNECTION_MODEM As Long = &H1
Private Const INTERNET_CONNECTION_PROXY As Long = &H4
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_RAS_INSTALLED As Long = &H10
'local variable(s) to hold property value(s)
Private mvarGetConnectionType As String 'local copy
Public Property Get GetConnectionType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GetConnectionType
 GetConnectionType = mvarGetConnectionType
End Property
Public Function GetInternetConnectedState() As Boolean
 Dim dwflags As Long 'Returns which Connection type
 Dim RetCode As Boolean 'If is connected
 'dwreserved needs to be set to 0&
 RetCode = InternetGetConnectedState(dwflags, 0&)
 Select Case RetCode
 Case dwflags And INTERNET_CONNECTION_CONFIGURED
 mvarGetConnectionType = "Local system has a valid connection to the Internet, but it might or might not be currently connected."
 Case dwflags And INTERNET_CONNECTION_LAN
 mvarGetConnectionType = "Local system uses a local area network to connect to the Internet."
 Case dwflags And INTERNET_CONNECTION_MODEM
 mvarGetConnectionType = "Local system uses a modem to connect to the Internet."
 Case dwflags And INTERNET_CONNECTION_PROXY
 mvarGetConnectionType = "Local system uses a proxy server to connect to the Internet."
 Case dwflags And INTERNET_CONNECTION_OFFLINE
 mvarGetConnectionType = "Local system is in offline mode."
 Case dwflags And INTERNET_RAS_INSTALLED
 mvarGetConnectionType = "Local system has RAS installed."
 End Select
 GetInternetConnectedState = RetCode
End Function
Useage:
Private Sub Form_Load()
 Dim test As New Class1
 Me.AutoRedraw = True
 Print test.GetInternetConnectedState
 Print test.GetConnectionType
End Sub
```

