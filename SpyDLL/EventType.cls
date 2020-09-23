VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Property Get SHCNE_RENAMEITEM() As Long
     SHCNE_RENAMEITEM = &H1
End Property

Public Property Get SHCNE_CREATE() As Long
    SHCNE_CREATE = &H2
End Property

Public Property Get SHCNE_DELETE() As Long
      SHCNE_DELETE = &H4
End Property

Public Property Get SHCNE_MKDIR() As Long
    SHCNE_MKDIR = &H8
End Property

Public Property Get SHCNE_RMDIR() As Long
    SHCNE_RMDIR = &H10
End Property

Public Property Get SHCNE_MEDIAINSERTED() As Long
    SHCNE_MEDIAINSERTED = &H20
End Property

Public Property Get SHCNE_MEDIAREMOVED() As Long
     SHCNE_MEDIAREMOVED = &H40
End Property
      
Public Property Get SHCNE_DRIVEREMOVED() As Long
    SHCNE_DRIVEREMOVED = &H80
End Property

Public Property Get SHCNE_DRIVEADD() As Long
    SHCNE_DRIVEADD = &H100
End Property
              
Public Property Get SHCNE_NETSHARE() As Long
    SHCNE_NETSHARE = &H200
End Property
           
Public Property Get SHCNE_NETUNSHARE() As Long
    SHCNE_NETUNSHARE = &H400
End Property
          
Public Property Get SHCNE_ATTRIBUTES() As Long
    SHCNE_ATTRIBUTES = &H800
End Property
       
Public Property Get SHCNE_UPDATEDIR() As Long
    SHCNE_UPDATEDIR = &H1000
End Property
         
Public Property Get SHCNE_UPDATEITEM() As Long
    SHCNE_UPDATEITEM = &H2000
End Property

Public Property Get SHCNE_SERVERDISCONNECT() As Long
    SHCNE_SERVERDISCONNECT = &H4000
End Property

Public Property Get SHCNE_UPDATEIMAGE() As Long
    SHCNE_UPDATEIMAGE = &H8000&
End Property
  
Public Property Get SHCNE_DRIVEADDGUI() As Long
    SHCNE_DRIVEADDGUI = &H10000
End Property
                             
Public Property Get SHCNE_RENAMEFOLDER() As Long
     SHCNE_RENAMEFOLDER = &H20000
End Property
   
Public Property Get SHCNE_FREESPACE() As Long
    SHCNE_FREESPACE = &H40000
End Property
   

Public Property Get SHCNE_ASSOCCHANGED() As Long
    SHCNE_ASSOCCHANGED = &H8000000
End Property
 
Public Property Get SHCNE_DISKEVENTS() As Long
    SHCNE_DISKEVENTS = &H2381F
End Property
  
Public Property Get SHCNE_INTERRUPT() As Long
    SHCNE_INTERRUPT = &H80000000
End Property

Public Property Get SHCNE_ALLEVENTS() As Long
     SHCNE_ALLEVENTS = &H7FFFFFFF
End Property

Public Property Get SHCNE_GLOBALEVENTS() As Long
     SHCNE_GLOBALEVENTS = &HC0581E0
End Property
 


