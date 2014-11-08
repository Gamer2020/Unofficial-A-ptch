VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":000C
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrShowUpdate 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   120
      Top             =   1080
   End
   Begin VB.Frame fraCopyright 
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   5775
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2010 HackMew"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1920
         TabIndex        =   16
         Top             =   195
         Width           =   2055
      End
   End
   Begin VB.Frame fraWorkingMode 
      Caption         =   "Working Mode"
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Tag             =   "1"
      Top             =   1560
      Width           =   5775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   14
         Top             =   240
         Width           =   5475
         Begin VB.OptionButton optWorkingMode 
            Caption         =   "Get patch info"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Tag             =   "4"
            Top             =   840
            Width           =   5175
         End
         Begin VB.OptionButton optWorkingMode 
            Caption         =   "Create a new APS patch"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Tag             =   "3"
            Top             =   480
            Width           =   5175
         End
         Begin VB.OptionButton optWorkingMode 
            Caption         =   "Apply an APS patch to a file"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Tag             =   "2"
            Top             =   120
            Value           =   -1  'True
            Width           =   5175
         End
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4560
      TabIndex        =   9
      Tag             =   "8"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   8
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   5520
      TabIndex        =   4
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4080
      Width           =   3855
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patch"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Tag             =   "7"
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modified File"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Tag             =   "6"
      Top             =   3720
      Width           =   885
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Original File"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Tag             =   "5"
      Top             =   3360
      Width           =   825
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      HelpContextID   =   100
      Begin VB.Menu mnuRun 
         Caption         =   "Run"
         Enabled         =   0   'False
         HelpContextID   =   101
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         HelpContextID   =   102
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      HelpContextID   =   103
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         HelpContextID   =   104
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiveUpdate 
         Caption         =   "Live Update"
         HelpContextID   =   10000
         Begin VB.Menu mnuCheckNow 
            Caption         =   "Check Now..."
            Enabled         =   0   'False
            HelpContextID   =   10024
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuAutomaticallyCheck 
            Caption         =   "Automatically Check"
            Checked         =   -1  'True
            HelpContextID   =   10025
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Copyright © 2010 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Const sMyName As String = "frmMain"
Private Const sAutoUpdateFile As String = "\autoupdate.dat"

' File Constants
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3&
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_BEGIN = 0&

' SendMessage Constants
Private Const WM_SETTEXT = &HC&

' SendMessage API
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

' Internet API
Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef lpSFlags As Long, ByVal dwReserved As Long) As Long

' File handling APIs
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

' Memory APIs
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDest As Any, ByRef pSource As Any, ByVal lLength As Long)
Private Declare Sub RtlFillMemory Lib "kernel32" (ByRef pDest As Any, ByVal lLength As Long, ByVal lFillByte As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32" (ByRef pDest As Any, ByVal lLength As Long)

Private Const lSignature As Long = &H31535041 ' "APS1"

Private Const lChunkBytes = 65536 ' 64k
Private Const lChunkInt = lChunkBytes \ 2&
Private Const lChunkLong = lChunkBytes \ 4&

Private Const lIntegerOffset = 65536
Private Const lMaxInteger = (lIntegerOffset \ 2&) - 1&

Private lCrcTable() As Long

Private bByteArray1() As Byte
Private bByteArray2() As Byte
Private bByteArray3() As Byte

Private lLongArray1() As Long
Private lLongArray2() As Long
Private lLongArray3() As Long

Private lByteAddress1 As Long
Private lByteAddress2 As Long
Private lByteAddress3 As Long

Private lDataPointer1 As Long
Private lDataPointer2 As Long
Private lDataPointer3 As Long

Private lLongAddress1 As Long
Private lLongAddress2 As Long
Private lLongAddress3 As Long

Private Type tPatchData
    lOffset As Long
    iCrc16_1 As Integer
    iCrc16_2 As Integer
End Type

Private Function Max(ByRef First As Long, ByRef Second As Long) As Long
    
    If First >= Second Then
        Max = First
    Else
        Max = Second
    End If
    
End Function

Private Sub CrcTableInit()
Const Poly = &H1021&
Dim i As Long
Dim j As Long
Dim lCrc As Long
Dim lTemp As Long
    
    For i = LBound(lCrcTable) To UBound(lCrcTable)
        
        lTemp = i * &H100&
        lCrc = 0&
        
        For j = 0& To 7&
            
            If ((lCrc Xor lTemp) And &H8000&) Then
                lCrc = ((lCrc * 2&) Xor Poly) And &HFFFF&
            Else
                lCrc = (lCrc * 2&) And &HFFFF&
            End If
            
            lTemp = (lTemp * 2&) And &HFFFF&
            
        Next j

        lCrcTable(i) = lCrc
            
    Next i
    
End Sub

'Private Function Crc16(ByRef Data() As Byte, Optional ByRef Size As Long = -1) As Long
'Dim i As Long
'
'    If Size = -1& Then
'        Size = UBound(Data) + 1&
'    End If
'
'    Crc16 = &HFFFF&
'
'    For i = 0& To Size - 1&
'        Crc16 = (((Crc16 * &H100&) Xor lCrcTable(((Crc16 \ &H100&) Xor Data(i))))) And &HFFFF&
'    Next i
'
'    If Crc16 > lMaxInteger Then
'        Crc16 = Crc16 - lIntegerOffset
'    End If
'
'End Function

Private Function Crc16(ByRef Data() As Byte) As Long
Dim i As Long
    
    Crc16 = &HFFFF&
    
    For i = 0& To &HFFFF&
        Crc16 = (((Crc16 * &H100&) Xor lCrcTable(((Crc16 \ &H100&) Xor Data(i))))) And &HFFFF&
    Next i
    
    If Crc16 > lMaxInteger Then
        Crc16 = Crc16 - lIntegerOffset
    End If
    
End Function

'Private Sub CreatePatch(ByRef PatchFile As String, ByRef Original As String, ByRef Modified As String)
'Const sThis As String = "CreatePatch"
'Dim i As Long
'Dim j As Long
'Dim k As Long
'Dim iPatchFile As Integer
'Dim iOriginal As Integer
'Dim iModified As Integer
'Dim iCrc16_1 As Integer
'Dim iCrc16_2 As Integer
'Dim lPatchSize As Long
'Dim lTemp As Long
'Dim bNull() As Byte
'Dim bTempArray1() As Byte
'Dim bTempArray2() As Byte
'Dim lCount As Long
'Dim lIncrement As Long
'
'    On Error Resume Next
'
'    ' Kill path file if it exists
'    If FileExists(PatchFile) Then
'        Kill PatchFile
'    End If
'
'    On Error GoTo LocalHandler
'
'    ' Store the addresses to the Long arrays data
'    RtlMoveMemory lLongAddress1, ByVal ArrPtr(lLongArray1), 4&
'    RtlMoveMemory lLongAddress2, ByVal ArrPtr(lLongArray2), 4&
'    RtlMoveMemory lLongAddress3, ByVal ArrPtr(lLongArray3), 4&
'
'    ' Change the pointers of the Long arrays so that they point to the Byte arrays data
'    RtlMoveMemory ByVal lLongAddress1 + 12&, lDataPointer1, 4&
'    RtlMoveMemory ByVal lLongAddress2 + 12&, lDataPointer2, 4&
'    RtlMoveMemory ByVal lLongAddress3 + 12&, lDataPointer3, 4&
'
'    ' Initialize temp arrays
'    ReDim iTempArray1(lChunkInt - 1&)
'    ReDim iTempArray2(lChunkInt - 1&)
'
'    ' Initialize null array
'    ReDim bNull(0)
'
'    ' Get the next free file number
'    iPatchFile = FreeFile
'
'    ' Open the patch file
'    Open PatchFile For Binary As #iPatchFile
'
'        ' Patch structure:
'        ' Signature (APS1) - 4 bytes
'        ' File Size XOR - 4 bytes
'        ' ---
'        ' Patch Address - 4 bytes
'        ' Patch Size - 2 bytes
'        ' Patch CRC16 (1) - 2 bytes
'        ' Patch CRC16 (2) - 2 bytes
'        ' Bytes to Patch (XOR) - up to 65535 bytes
'        ' ----
'        ' Patch Address - 4 bytes
'        ' Patch Size - 2 bytes
'        ' ...
'
'        Screen.MousePointer = vbHourglass
'
'        ' Write the signature
'        Put #iPatchFile, 1&, lSignature
'
'        ' Get the next free file number
'        iOriginal = FreeFile
'
'        ' Open the original file
'        Open Original For Binary As #iOriginal
'
'            ' Get the next free file number
'            iModified = FreeFile
'
'            ' Open the modified file
'            Open Modified For Binary As iModified
'
'                ' Write the XORed file size
'                Put #iPatchFile, , LOF(iOriginal) Xor LOF(iModified)
'
'                ' Loop through the biggest file in chunks
'                For i = 0& To Max(LOF(iOriginal), LOF(iModified)) \ (lChunkBytes - 1&)
'
'                    ' Get the original and modified chunks
'                    Get #iOriginal, , bByteArray1
'                    Get #iModified, , bByteArray2
'
'                    ' Check if there are some differences
'                    If InStrB(bByteArray1, bByteArray2) <> 1& Then
'
'                        ' XOR the original and the modified chunks
'                        For j = 0& To lChunkLong - 1&
'                            lLongArray3(j) = lLongArray1(j) Xor lLongArray2(j)
'                        Next j
'
'                        ' Get the zero-based index of the first unchanged byte in the XORed chunk
'                        lTemp = InStrB(bByteArray3, bNull) - 1&
'
'                        ' Check if there are some unchanged bytes
'                        If lTemp < (lChunkBytes - 1&) Then
'
'                            ' Decrease the index by 1, unless it's zero
'                            If lTemp <> 0& Then
'                                lTemp = lTemp - 1&
'                            End If
'
'                            ' Loop through the XORed bytes
'                            For j = lTemp To ((lChunkBytes - 2&) + lTemp)
'
'                                ' Check if the byte was changed
'                                If bByteArray1(j - lTemp) <> bByteArray2(j - lTemp) Then
'
'                                    k = j - lTemp
'
'                                    ' Increment change counter
'                                    lCount = lCount + 1&
'
'                                    ' Write the change address
'                                    Put #iPatchFile, , ((lChunkBytes - 1&) * i) + k
'
'                                    ' Reset patch size
'                                    lPatchSize = 0&
'
'                                    ' Calculate patch size
'                                    Do While bByteArray3(j - lTemp) <> 0
'                                        lPatchSize = lPatchSize + 1&
'                                        j = j + 1&
'                                    Loop
'
'                                    ' Fill another temp array with the XORed data
'                                    ReDim bTempArray1(lPatchSize - 1&)
'
'                                    ' Calculate the first CRC16
'                                    RtlMoveMemory bTempArray1(0), bByteArray1(k), lPatchSize
'                                    iCrc16_1 = CInt(Crc16(bTempArray1))
'
'                                    ' Calculate the second CRC16
'                                    RtlMoveMemory bTempArray1(0), bByteArray2(k), lPatchSize
'                                    iCrc16_2 = CInt(Crc16(bTempArray1))
'
'                                    RtlMoveMemory bTempArray1(0), bByteArray3(k), lPatchSize
'
'                                    ' Check if the patch size is 4 bytes or more
'                                    If lPatchSize > 3& Then
'
'                                        ' Reset temp array
'                                        ReDim bTempArray2(lPatchSize - 1&)
'
'                                        ' Fill it with the first byte of the byte temp array
'                                        RtlFillMemory bTempArray2(0), lPatchSize, bTempArray1(0)
'
'                                        ' Adjust the patch size if needed to fit into an Integer
'                                        If lPatchSize > lMaxInteger Then
'                                            lPatchSize = lPatchSize - lIntegerOffset
'                                        End If
'
'                                        ' Check if the patch data can be compressed
'                                        If InStrB(bTempArray1, bTempArray2) <> 1& Then
'                                            ' No compression
'                                            ' Write the patch data along with the size and the checksum
'                                            Put #iPatchFile, , CInt(lPatchSize)
'                                            Put #iPatchFile, , iCrc16_1
'                                            Put #iPatchFile, , iCrc16_2
'                                            Put #iPatchFile, , bTempArray1
'                                        Else
'                                            ' Compression
'                                            ' Write the compressed data along with the size and checksum
'                                            Put #iPatchFile, , 0
'                                            Put #iPatchFile, , CInt(lPatchSize)
'                                            Put #iPatchFile, , iCrc16_1
'                                            Put #iPatchFile, , iCrc16_2
'                                            Put #iPatchFile, , bTempArray1(0)
'                                        End If
'
'                                    Else
'                                        ' Could have been compressed, but it wouldn't have made sense
'                                        ' Write data as usual
'                                        Put #iPatchFile, , CInt(lPatchSize)
'                                        Put #iPatchFile, , iCrc16_1
'                                        Put #iPatchFile, , iCrc16_2
'                                        Put #iPatchFile, , bTempArray1
'                                    End If
'
'                                End If
'
'                            Next j
'
'                        ' All bytes are changed
'                        ElseIf lTemp = (lChunkBytes - 1&) Then
'
'                            ' Increment the change counter
'                            lCount = lCount + 1&
'
'                            ' Write the patch address
'                            Put #iPatchFile, , (lTemp * i)
'
'                            ' Set the patch size
'                            lPatchSize = lTemp
'
'                            ReDim bTempArray1(lPatchSize - 1&)
'
'                            RtlMoveMemory bTempArray1(0), bByteArray1(0), lPatchSize
'                            iCrc16_1 = CInt(Crc16(bTempArray1))
'
'                            RtlMoveMemory bTempArray1(0), bByteArray2(0), lPatchSize
'                            iCrc16_2 = CInt(Crc16(bTempArray1))
'
'                            RtlMoveMemory bTempArray1(0), bByteArray3(0), lPatchSize
'
'                            ReDim bTempArray2(lPatchSize - 1&)
'                            RtlFillMemory bTempArray2(0), lPatchSize, bTempArray1(0)
'
'                            If lPatchSize > lMaxInteger Then
'                                lPatchSize = lPatchSize - lIntegerOffset
'                            End If
'
'                            If InStrB(bTempArray1, bTempArray2) <> 1& Then
'                                Put #iPatchFile, , CInt(lPatchSize)
'                                Put #iPatchFile, , iCrc16_1
'                                Put #iPatchFile, , iCrc16_2
'                                Put #iPatchFile, , bTempArray1
'                            Else
'                                Put #iPatchFile, , 0
'                                Put #iPatchFile, , CInt(lPatchSize)
'                                Put #iPatchFile, , iCrc16_1
'                                Put #iPatchFile, , iCrc16_2
'                                Put #iPatchFile, , bTempArray1(0)
'                            End If
'
'                        End If
'
'                    End If
'
'                Next i
'
'                Close #iModified
'
'            Close #iOriginal
'
'        Close #iPatchFile
'
'    ' Clear Long array data pointers
'    RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
'    RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
'    RtlMoveMemory ByVal lLongAddress3 + 12&, 0&, 4&
'
'    Screen.MousePointer = vbDefault
'    MsgBox LoadString(ID_PATCHEDSUCCESSFULLY), vbInformation
'    Exit Sub
'
'LocalHandler:
'
'    ' Clear Long array data pointers
'    RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
'    RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
'    RtlMoveMemory ByVal lLongAddress3 + 12&, 0&, 4&
'
'    Screen.MousePointer = vbDefault
'
'    Select Case GlobalHandler(sThis, sMyName)
'        Case vbRetry
'            Resume
'        Case vbAbort
'            Quit
'        Case Else
'            Resume Next
'    End Select
'
'End Sub

Private Sub CreatePatch(ByRef PatchFile As String, ByRef Original As String, ByRef Modified As String)
Const sThis As String = "CreatePatch"
Dim i As Long
Dim j As Long
Dim iPatchFile As Integer
Dim iOriginal As Integer
Dim iModified As Integer
Dim iCrc16_1 As Integer
Dim iCrc16_2 As Integer
Dim tPatch As tPatchData
                    
    On Error GoTo LocalHandler
    
    ' Kill the patch file
    DeleteFileW (StrPtr(PatchFile))
                    
    ' Store the addresses to the Long arrays data
    RtlMoveMemory lLongAddress1, ByVal ArrPtr(lLongArray1), 4&
    RtlMoveMemory lLongAddress2, ByVal ArrPtr(lLongArray2), 4&
    RtlMoveMemory lLongAddress3, ByVal ArrPtr(lLongArray3), 4&
    
    ' Change the pointers of the Long arrays so that they point to the Byte arrays data
    RtlMoveMemory ByVal lLongAddress1 + 12&, lDataPointer1, 4&
    RtlMoveMemory ByVal lLongAddress2 + 12&, lDataPointer2, 4&
    RtlMoveMemory ByVal lLongAddress3 + 12&, lDataPointer3, 4&
    
    ' Get the next free file number
    iPatchFile = FreeFile
                
    ' Open the patch file
    Open PatchFile For Binary As #iPatchFile
        
        ' Patch structure:
        ' Signature (APS1) - 4 bytes
        ' File Size XOR - 4 bytes
        ' ---
        ' Patch Address - 4 bytes
        ' Patch CRC16 (1) - 2 bytes
        ' Patch CRC16 (2) - 2 bytes
        ' Bytes to Patch (XOR) - 64 KB
        ' ----
        ' Patch Address - 4 bytes
        ' Patch CRC16 (1) - 2 bytes
        ' ...
        
        Screen.MousePointer = vbHourglass
                        
        ' Write the signature
        Put #iPatchFile, 1&, lSignature
                        
        ' Get the next free file number
        iOriginal = FreeFile
            
        ' Open the original file
        Open Original For Binary As #iOriginal
                     
            ' Get the next free file number
            iModified = FreeFile
                            
            ' Open the modified file
            Open Modified For Binary As iModified
                                
                ' Write the respective file sizes
                Put #iPatchFile, , LOF(iOriginal)
                Put #iPatchFile, , LOF(iModified)
                                
                ' Loop through the biggest file in chunks of 64 KB
                For i = 0& To Max(LOF(iOriginal), LOF(iModified)) \ lChunkBytes
                                    
                    ' Get the original and modified chunks
                    Get #iOriginal, , bByteArray1
                    Get #iModified, , bByteArray2
                                    
                    ' Check if there are some differences
                    If InStrB(bByteArray1, bByteArray2) <> 1& Then
                                        
                        ' XOR the original and the modified chunks
                        For j = 0& To lChunkLong - 1&
                            lLongArray3(j) = lLongArray1(j) Xor lLongArray2(j)
                        Next j
                        
                        tPatch.lOffset = lChunkBytes * i
                        tPatch.iCrc16_1 = CInt(Crc16(bByteArray1))
                        tPatch.iCrc16_2 = CInt(Crc16(bByteArray2))
                                
                        Put #iPatchFile, , tPatch
                        Put #iPatchFile, , lLongArray3
                        
                    End If
                                
                Next i
                                
                Close #iModified
                
            Close #iOriginal
            
        Close #iPatchFile
                    
    ' Clear Long array data pointers
    RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
    RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
    RtlMoveMemory ByVal lLongAddress3 + 12&, 0&, 4&
                    
    Screen.MousePointer = vbDefault
    MsgBox LoadString(ID_PATCHEDSUCCESSFULLY), vbInformation
    Exit Sub
    
LocalHandler:

    ' Clear Long array data pointers
    RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
    RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
    RtlMoveMemory ByVal lLongAddress3 + 12&, 0&, 4&
    
    Screen.MousePointer = vbDefault

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
                    
End Sub

Private Sub TruncateFile(ByRef FileName As String, ByVal Size As Long)
Dim hFile As Long
    
    ' Get a handle for the file
    hFile = CreateFileW(StrPtr(FileName), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    
    ' Seek the location to then new size
    SetFilePointer hFile, Size, 0&, FILE_BEGIN
    
    ' Set the end of file there and close the file
    SetEndOfFile hFile
    CloseHandle hFile

End Sub

'Private Sub ApplyPatch(ByRef PatchFile As String, ByRef Original As String)
'Const sThis As String = "ApplyPatch"
'Const lIntegerOffset As Long = 65536
'Const lMaxInteger As Long = (lIntegerOffset \ 2&) - 1&
'Dim i As Long
'Dim j As Long
'Dim iPatchFile As Integer
'Dim iOriginal As Integer
'Dim iModified As Integer
'Dim lOriginalSize As Long
'Dim lFileSize As Long
'Dim lPatchSize As Long
'Dim iPatchCrc16_1 As Integer
'Dim iPatchCrc16_2 As Integer
'Dim lCrc16 As Long
'Dim bTemp As Byte
'Dim iTemp As Integer
'Dim lTemp As Long
'Dim lAddress As Long
'Dim lBytesLeft As Long
'Dim lTotalBytes As Long
'Dim bTempArray1() As Byte
'Dim lCounter As Long
'
'    On Error GoTo LocalHandler
'
'    RtlMoveMemory lLongAddress1, ByVal ArrPtr(lLongArray1), 4&
'    RtlMoveMemory lLongAddress2, ByVal ArrPtr(lLongArray2), 4&
'
'    RtlMoveMemory ByVal lLongAddress1 + 12&, lDataPointer1, 4&
'    RtlMoveMemory ByVal lLongAddress2 + 12&, lDataPointer2, 4&
'
'    iPatchFile = FreeFile
'
'    Open PatchFile For Binary As #iPatchFile
'
'        Get #iPatchFile, , lTemp
'
'        If lTemp = lSignature Then ' "APS1"
'
'            iOriginal = FreeFile
'            Screen.MousePointer = vbHourglass
'
'            lTotalBytes = LOF(iPatchFile)
'            lBytesLeft = lTotalBytes
'
'            Open Original For Binary As #iOriginal
'
'                Get #iPatchFile, , lFileSize
'
'                lOriginalSize = LOF(iOriginal)
'                lBytesLeft = lBytesLeft - 8&
'
'                If lBytesLeft > 2048& Then
'                    lblStatus.Visible = True
'                    lblStatus.Caption = "[" & (lTotalBytes - lBytesLeft) & " / " & lTotalBytes & "]"
'                End If
'
'                Do While lBytesLeft > 0&
'
'                    lCounter = lCounter + 1&
'                    lblStatus.Visible = True
'
'                    If (lCounter Mod 2048& = 0&) Then
'                        lblStatus.Caption = "[" & (lTotalBytes - lBytesLeft) & " / " & lTotalBytes & "]"
'                        MyDoEvents
'                    End If
'
'                    Get #iPatchFile, , lAddress
'                    Get #iPatchFile, , iTemp
'
'                    If iTemp <> 0 Then
'
'                        lPatchSize = iTemp
'                        Get #iPatchFile, , iPatchCrc16_1
'                        Get #iPatchFile, , iPatchCrc16_2
'
'                        If lPatchSize < 0& Then
'                            lPatchSize = lPatchSize + lIntegerOffset
'                        End If
'
'                        RtlZeroMemory bByteArray1(0), lChunkBytes - 1&
'                        RtlZeroMemory bByteArray2(0), lChunkBytes - 1&
'
'                        RtlMoveMemory ByVal lByteAddress1 + 16&, lPatchSize, 4&
'                        RtlMoveMemory ByVal lByteAddress2 + 16&, lPatchSize, 4&
'
'                        Get #iOriginal, lAddress + 1&, bByteArray1
'                        Get #iPatchFile, , bByteArray2
'
'                        RtlMoveMemory ByVal lByteAddress1 + 16&, lChunkBytes - 1&, 4&
'                        RtlMoveMemory ByVal lByteAddress2 + 16&, lChunkBytes - 1&, 4&
'
'                        lBytesLeft = lBytesLeft - 10& - lPatchSize
'
'                    Else
'
'                        Get #iPatchFile, , iTemp
'
'                        lPatchSize = iTemp
'                        Get #iPatchFile, , iPatchCrc16_1
'                        Get #iPatchFile, , iPatchCrc16_2
'                        Get #iPatchFile, , bTemp
'
'                        If lPatchSize < 0& Then
'                            lPatchSize = lPatchSize + lIntegerOffset
'                        End If
'
'                        RtlZeroMemory bByteArray1(0), lChunkBytes - 1&
'                        RtlZeroMemory bByteArray2(0), lChunkBytes - 1&
'
'                        RtlMoveMemory ByVal lByteAddress1 + 16&, lPatchSize, 4&
'                        RtlMoveMemory ByVal lByteAddress2 + 16&, lPatchSize, 4&
'
'                        Get #iOriginal, lAddress + 1&, bByteArray1
'                        RtlFillMemory bByteArray2(0), lPatchSize, bTemp
'
'                        RtlMoveMemory ByVal lByteAddress1 + 16&, lChunkBytes - 1&, 4&
'                        RtlMoveMemory ByVal lByteAddress2 + 16&, lChunkBytes - 1&, 4&
'
'                        lBytesLeft = lBytesLeft - 13&
'
'                    End If
'
'                    lCrc16 = Crc16(bByteArray1, lPatchSize)
'
'                    For i = 0 To lPatchSize \ 4&
'                        lLongArray1(i) = lLongArray1(i) Xor lLongArray2(i)
'                    Next i
'
'                    If lCrc16 = iPatchCrc16_1 Then
'
'                        RtlMoveMemory ByVal lByteAddress1 + 16&, lPatchSize, 4&
'                        Put #iOriginal, lAddress + 1&, bByteArray1
'                        RtlMoveMemory ByVal lByteAddress1 + 16&, lChunkBytes - 1&, 4&
'
'                    ElseIf lCrc16 = iPatchCrc16_2 Then
'
'                        RtlMoveMemory ByVal lByteAddress1 + 16&, lPatchSize, 4&
'                        Put #iOriginal, lAddress + 1&, bByteArray1
'                        RtlMoveMemory ByVal lByteAddress1 + 16&, lChunkBytes - 1&, 4&
'
'                    Else
'
'                        RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
'                        RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
'                        Screen.MousePointer = vbDefault
'                        lblStatus.Visible = False
'                        MsgBox LoadString(ID_FILENOTVALID), vbExclamation
'                        Exit Sub
'
'                    End If
'
'                Loop
'
'                If (lOriginalSize Xor lFileSize) > lOriginalSize Then
'                    TruncateFile Original, lOriginalSize Xor lFileSize
'                ElseIf (lOriginalSize Xor lFileSize) = 0& Then
'                    TruncateFile Original, 0&
'                End If
'
'            Close #iOriginal
'
'            Screen.MousePointer = vbDefault
'            lblStatus.Visible = False
'
'        Else
'            Screen.MousePointer = vbDefault
'            MsgBox LoadString(ID_PATCHNOTVALID), vbExclamation
'        End If
'
'    Close #iPatchFile
'
'    MsgBox LoadString(ID_PATCHINGCOMPLETED), vbInformation
'
'    RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
'    RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
'
'    Exit Sub
'
'LocalHandler:
'
'    RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
'    RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
'    Screen.MousePointer = vbDefault
'
'    Select Case GlobalHandler(sThis, sMyName)
'        Case vbRetry
'            Resume
'        Case vbAbort
'            Quit
'        Case Else
'            Resume Next
'    End Select
'
'End Sub

Private Sub ApplyPatch(ByRef PatchFile As String, ByRef Original As String)
Const sThis As String = "ApplyPatch"
Dim i As Long
Dim j As Long
Dim iPatchFile As Integer
Dim iOriginal As Integer
Dim lOriginalSize As Long
Dim lFileSize1 As Long
Dim lFileSize2 As Long
Dim iCrc16 As Integer
Dim lTemp As Long
Dim lBytesLeft As Long
Dim tPatch As tPatchData
Dim fIsOriginal As Boolean
Dim fIsModified As Boolean

    On Error GoTo LocalHandler

    RtlMoveMemory lLongAddress1, ByVal ArrPtr(lLongArray1), 4&
    RtlMoveMemory lLongAddress2, ByVal ArrPtr(lLongArray2), 4&

    RtlMoveMemory ByVal lLongAddress1 + 12&, lDataPointer1, 4&
    RtlMoveMemory ByVal lLongAddress2 + 12&, lDataPointer2, 4&

    iPatchFile = FreeFile

    Open PatchFile For Binary As #iPatchFile

        Get #iPatchFile, , lTemp

        If lTemp = lSignature Then ' "APS1"

            Screen.MousePointer = vbHourglass
            
            iOriginal = FreeFile
            
            Get #iPatchFile, , lFileSize1
            Get #iPatchFile, , lFileSize2

            lBytesLeft = LOF(iPatchFile) - 12&

            Open Original For Binary As #iOriginal

                lOriginalSize = LOF(iOriginal)

                Do While lBytesLeft > 0&
                    
                    Get #iPatchFile, , tPatch
                    Get #iOriginal, tPatch.lOffset + 1&, bByteArray1
                    Get #iPatchFile, , bByteArray2

                    lBytesLeft = lBytesLeft - lChunkBytes - 8&
                    iCrc16 = CInt(Crc16(bByteArray1))

                    For i = 0& To lChunkLong - 1&
                        lLongArray1(i) = lLongArray1(i) Xor lLongArray2(i)
                    Next i

                    If iCrc16 = tPatch.iCrc16_1 Then
                        
                        If fIsModified Then
                   
                            If fIsOriginal Then
                                
                                RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
                                RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
                                Screen.MousePointer = vbDefault
                                MsgBox LoadString(ID_FILENOTVALID), vbExclamation
                                Exit Sub
                              
                              End If
                            
                      
                            
                        End If
                        
                        Put #iOriginal, tPatch.lOffset + 1&, bByteArray1
                        fIsOriginal = True

                    ElseIf iCrc16 = tPatch.iCrc16_2 Then
                        
                        If fIsOriginal Then
               
                            If fIsModified Then
                            
                                RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
                                RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
                                Screen.MousePointer = vbDefault
                                MsgBox LoadString(ID_FILENOTVALID), vbExclamation
                                Exit Sub
                            
                            End If

                        End If
                        
                        Put #iOriginal, tPatch.lOffset + 1&, bByteArray1
                        fIsModified = True

                    Else

                        RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
                        RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
                        Screen.MousePointer = vbDefault
                        MsgBox LoadString(ID_FILENOTVALID), vbExclamation
                        Exit Sub

                    End If

                Loop
                
                If fIsOriginal Then
                    TruncateFile Original, lFileSize1
                Else
                    TruncateFile Original, lFileSize2
                End If

            Close #iOriginal
            Screen.MousePointer = vbDefault

        Else
            Screen.MousePointer = vbDefault
            MsgBox LoadString(ID_PATCHNOTVALID), vbExclamation
        End If

    Close #iPatchFile

    MsgBox LoadString(ID_PATCHINGCOMPLETED), vbInformation

    RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
    RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&

    Exit Sub

LocalHandler:

    RtlMoveMemory ByVal lLongAddress1 + 12&, 0&, 4&
    RtlMoveMemory ByVal lLongAddress2 + 12&, 0&, 4&
    Screen.MousePointer = vbDefault

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select

End Sub

Private Function PadHex$(ByVal lNumber As Long, ByVal lSize As Long)
    PadHex$ = RightB$("0000000" & Hex$(lNumber), lSize * 2&)
End Function

Private Sub GetPatchInfo(ByRef PatchFile As String)
Const sThis As String = "GetPatchInfo"
Dim cString As cStringBuilder
Const sHexPrefix As String = "0x"
Const sDotTab As String = "." & vbTab
Const lIntegerOffset As Long = 65536
Dim iPatchFile As Integer
Dim bTemp As Byte
Dim iTemp As Integer
Dim lTemp As Long
Dim lAddress As Long
Dim lBytesLeft As Long
Dim lPatchSize As Long
Dim lCounter As Long
    
    On Error GoTo LocalHandler
    
    iPatchFile = FreeFile

    Open PatchFile For Binary As #iPatchFile
        
        Get #iPatchFile, , lTemp
            
        If lTemp = lSignature Then ' "APS1"
            
            lAddress = 12&
            lBytesLeft = LOF(iPatchFile) - 12&
            
            Screen.MousePointer = vbHourglass
            Load frmPatchInfo
            
            Set cString = New cStringBuilder
                
            Do While lBytesLeft > 0&
                
                Get #iPatchFile, lAddress + 1&, lTemp
                lCounter = lCounter + 1&
                cString.Append lCounter & sDotTab & sHexPrefix & PadHex$(lTemp, 8&) & vbNewLine
                
                lAddress = lAddress + lChunkBytes + 8&
                lBytesLeft = lBytesLeft - lChunkBytes - 8&
            
            Loop
            
            SendMessageA frmPatchInfo.txtPatchData.hWnd, WM_SETTEXT, 0&, ByVal cString.ToString
            frmPatchInfo.Show , Me
            Screen.MousePointer = vbDefault
            
            cString.Clear
            Set cString = Nothing
            
        Else
            Screen.MousePointer = vbDefault
            MsgBox LoadString(ID_PATCHNOTVALID), vbExclamation
        End If
                   
    Close iPatchFile
    Exit Sub
    
LocalHandler:
    
    Screen.MousePointer = vbDefault
    
    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
Const sAPS As String = "Alternate Patching System (*.aps)|*.aps|"
Dim sFile As String

    If Index <> 2 Then
        sFile = ShowOpen(Me.hWnd, , , "GameBoy Advance (*.gba; *.agb; *.bin)|*.gba;*.agb;*.bin|GameBoy Color (*.gb; *.gbc; *.bin)|*.gb;*.gbc;*.bin|")
    Else
        
        If optWorkingMode(1).Value = False Then
            sFile = ShowOpen(Me.hWnd, , , sAPS)
        Else
            sFile = ShowSave(Me.hWnd, , , sAPS, , ".aps")
        End If
        
    End If
    
    If LenB(sFile) <> 0& Then
        txtFile(Index).Text = sFile
    End If

End Sub

Private Sub cmdRun_Click()
Const sThis As String = "cmdRun_Click"

    On Error GoTo LocalHandler
    
    cmdRun.Enabled = False
    mnuRun.Enabled = False

    Select Case True
        
        ' Apply a patch
        Case optWorkingMode(0).Value
            
            ApplyPatch txtFile(2).Text, txtFile(0).Text
            
        ' Create a patch
        Case optWorkingMode(1).Value
        
            CreatePatch txtFile(2).Text, txtFile(0).Text, txtFile(1).Text
        
        ' Get patch info
        Case optWorkingMode(2).Value
        
            GetPatchInfo txtFile(2).Text
            
    End Select
    
    cmdRun.Enabled = True
    mnuRun.Enabled = True
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select

End Sub

Private Sub UpdateCheck()
Const sThis As String = "UpdateCheck"
Dim iFileNum As Integer
Dim bUpdate As Byte

    On Error GoTo LocalHandler

    ' Check if the AutoUpdate file exits
    If FileExists(App.Path & sAutoUpdateFile) Then
        
        ' Get the next free number
        iFileNum = FreeFile
        
        ' Open the file
        Open App.Path & sAutoUpdateFile For Binary As #iFileNum
        
            ' Get the value
            Get #iFileNum, , bUpdate
            
        Close #iFileNum
        
        ' Set the AutoCheck accordingly
        mnuAutomaticallyCheck.Checked = CBool(bUpdate)
     
    Else
        ' No file, defaulting to True
        mnuAutomaticallyCheck.Checked = True
    End If
    
    ' If the AutoCheck is enabled
    If mnuAutomaticallyCheck.Checked Then
        
        ' Load the form unattended
        Load frmUpdate
        frmUpdate.IsUnattended = True
       
        ' Increment step
        frmUpdate.NextStep
        
    End If
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Sub

Private Sub Form_Load()
Const sThis As String = "Form_Load"

    On Error GoTo LocalHandler
    
    ReDim lCrcTable(255&)
    CrcTableInit
        
'    ReDim bByteArray1(lChunkBytes - 2&)
'    ReDim bByteArray2(lChunkBytes - 2&)
    ReDim bByteArray1(lChunkBytes - 1&)
    ReDim bByteArray2(lChunkBytes - 1&)
    ReDim bByteArray3(lChunkBytes - 1&)
    ReDim lLongArray1(lChunkLong - 1&)
    ReDim lLongArray2(lChunkLong - 1&)
    ReDim lLongArray3(lChunkLong - 1&)
    
    RtlMoveMemory lByteAddress1, ByVal ArrPtr(bByteArray1), 4&
    RtlMoveMemory lByteAddress2, ByVal ArrPtr(bByteArray2), 4&
    RtlMoveMemory lByteAddress3, ByVal ArrPtr(bByteArray3), 4&
    RtlMoveMemory lDataPointer1, ByVal lByteAddress1 + 12&, 4&
    RtlMoveMemory lDataPointer2, ByVal lByteAddress2 + 12&, 4&
    RtlMoveMemory lDataPointer3, ByVal lByteAddress3 + 12&, 4&
    
    ' Set the form's icon
    SetIcon Me.hWnd, "AAA"
    
    ' Localize the form
    Localize Me
    
    ' Set the caption and the copyright
    Me.Caption = App.Title & " - " & App.ProductName
    lblCopyright.Caption = App.LegalCopyright
    
    ' Look for updates
    UpdateCheck
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select

End Sub

Private Sub ValidateFiles()
    
    ' Assume it's disabled
    cmdRun.Enabled = False
    mnuRun.Enabled = False
    
    Select Case True
        
        ' Apply a patch
        Case optWorkingMode(0).Value
        
            ' Make sure the original file exists
            If FileExists(txtFile(0).Text) Then
                
                ' Check if the patch file textbox is empty
                If LenB(txtFile(2).Text) <> 0& Then
                    cmdRun.Enabled = True
                    mnuRun.Enabled = True
                End If
                
            End If
        
        ' Create a patch
        Case optWorkingMode(1).Value
        
            ' Make sure the original file exists
            If FileExists(txtFile(0).Text) Then

                ' Make sure the modified file exists too
                If FileExists(txtFile(1).Text) Then
                    
                    ' Check if the two files are the same one
                    If txtFile(0).Text <> txtFile(1).Text Then

                        ' Check if the patch file textbox is empty
                        If LenB(txtFile(2).Text) <> 0& Then
                            cmdRun.Enabled = True
                            mnuRun.Enabled = True
                        End If
                    
                    End If
                    
                End If
                
            End If
        
        ' Get patch info
        Case optWorkingMode(2).Value
        
            ' Make sure the patch file exists
            If FileExists(txtFile(2).Text) Then
                cmdRun.Enabled = True
                mnuRun.Enabled = True
            End If
        
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const sThis As String = "Form_Unload"
    
    On Error GoTo LocalHandler
    
    ' Free the memory associated with the form
    Set frmMain = Nothing
    
    ' Ensure all the forms but this one are unloaded
    Do While Forms.Count > 1&
        Unload Forms(Forms.Count - 1&)
    Loop
    
    If m_hMod Then
        ' Free the previously loaded library
        FreeLibrary m_hMod
    End If
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show , frmMain
End Sub

Private Sub mnuAutomaticallyCheck_Click()
Const sThis = "mnuAutomaticallyCheck_Click"
Dim iFileNum As Integer
    
    On Error GoTo LocalHandler
    
    ' Get the next free number
    iFileNum = FreeFile
    
    ' Toggle the Checked property
    mnuAutomaticallyCheck.Checked = Not mnuAutomaticallyCheck.Checked
        
    ' Open the AutoUpdate file
    Open App.Path & sAutoUpdateFile For Binary As #iFileNum
    
        ' Write 0 or 1 accorgindly
        Put #iFileNum, , CByte(-CInt(mnuAutomaticallyCheck.Checked))
        
    Close #iFileNum
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Sub

Private Sub mnuCheckNow_Click()
    frmUpdate.Show , Me
End Sub

Private Sub mnuRun_Click()
    cmdRun_Click
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Function IsOpen(ByRef FormName As String) As Boolean
Const sThis As String = "IsOpen"
Dim i As Long
    
    On Error GoTo LocalHandler
    
    ' Loop through the forms
    For i = 0 To Forms.Count - 1&
        
        ' Check if we got a match
        If InStrB(Forms(i).Name, FormName) Then
            
            ' Form is open, exit
            IsOpen = True
            Exit Function
            
        End If
        
    Next i
    Exit Function
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Function

Private Sub mnuLiveUpdate_Click()
    
    ' Enable the CheckNow menu only if the computer
    ' is connected to the Internet and the update form isn't open
    mnuCheckNow.Enabled = CBool(InternetGetConnectedState(0&, 0&)) And (IsOpen("frmUpdate") = False)
    
End Sub

Private Sub optWorkingMode_Click(Index As Integer)

    Select Case Index
    
        ' Apply patch
        Case 0
        
            cmdBrowse(0).Enabled = True
            cmdBrowse(1).Enabled = False
            lblFile(0).Enabled = True
            lblFile(1).Enabled = False
            txtFile(0).Enabled = True
            txtFile(1).Enabled = False
        
        ' Create patch
        Case 1
               
            cmdBrowse(0).Enabled = True
            cmdBrowse(1).Enabled = True
            lblFile(0).Enabled = True
            lblFile(1).Enabled = True
            txtFile(0).Enabled = True
            txtFile(1).Enabled = True
        
        ' Get patch info
        Case 2
        
            cmdBrowse(0).Enabled = False
            cmdBrowse(1).Enabled = False
            lblFile(0).Enabled = False
            lblFile(1).Enabled = False
            txtFile(0).Enabled = False
            txtFile(1).Enabled = False
    
    End Select
    
    ValidateFiles

End Sub

Private Sub tmrShowUpdate_Timer()
Const sThis As String = "tmrShowUpdate_Timer"

    On Error GoTo LocalHandler
    
    ' Show the update form
    frmUpdate.IsUnattended = False
    frmUpdate.Show , Me
    
    ' Disable the timer
    tmrShowUpdate.Enabled = False
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Sub

Private Sub txtFile_Change(Index As Integer)
    ValidateFiles
End Sub
