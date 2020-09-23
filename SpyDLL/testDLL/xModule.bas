Attribute VB_Name = "Module1"
'Private x As New OO7.Spy
Private x As New Spy
    
Public Function WndProc3(ByVal hwnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
Select Case uMsg
       Case x.SHNOTIFY
            MsgBox "Activity in Folder"
       Case x.NCDESTROY
            MsgBox "bye"
            
End Select
WndProc3 = x.NotifyWindows(hwnd, uMsg, wParam, lParam)

End Function


