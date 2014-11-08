Attribute VB_Name = "modMisc"
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Public Function FileExists(FileName As String) As Boolean
    
    On Error GoTo Hell

    ' Make sure the file name isn't empty
    If LenB(FileName) Then
    
        ' Check the Dir$ return value
        If LenB(Dir$(FileName)) Then
            
            ' The file exists
            FileExists = True
            
        End If
        
    End If
    
Hell:
End Function
