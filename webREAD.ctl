VERSION 5.00
Begin VB.UserControl webREAD 
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   1695
   ToolboxBitmap   =   "webREAD.ctx":0000
   Begin VB.Image picICON 
      Height          =   240
      Left            =   0
      Picture         =   "webREAD.ctx":0312
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "webREAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstrFileFromURL As String
Private fNAME As String
Private f() As Byte
Event DownloadComplete()
Event DownloadProgress(bytesREAD As Long, bytesMAX As Long)
Event DownloadError(errNUM As Long, errSTR As String)
Public Property Let OutputFile(fSTR As String)
    fNAME = fSTR
End Property
Public Property Get OutputFile() As String
    OutputFile = fNAME
End Property
Public Property Let FileURL(ByVal NewString As String)
    On Error GoTo ErrorFileURL
    If fNAME <> "" Then
        mstrFileFromURL = NewString
        If (Ambient.UserMode = True) And (NewString <> "") Then
            AsyncRead NewString, vbAsyncTypeByteArray, "FileURL", vbAsyncReadForceUpdate
        End If
    End If
    Exit Property
ErrorFileURL:
    RaiseEvent DownloadError(Err.Number, Err.Description)
End Property
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error Resume Next
    Select Case AsyncProp.PropertyName
        Case "FileURL"
            If AsyncProp.bytesMAX = 0 Then
                RaiseEvent DownloadError(-1, "Invalid URL.  Please check the URL and try again.")
            Else
                f = AsyncProp.Value
                Open fNAME For Binary Access Write As #1
                '------------------------------------------------------------
                ' Write it out
                '------------------------------------------------------------
                Put #1, , f
                '------------------------------------------------------------
                ' Close it
                '------------------------------------------------------------
                Close #1
            End If
            RaiseEvent DownloadComplete
        Case Else
    End Select
End Sub
Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    RaiseEvent DownloadProgress(AsyncProp.bytesREAD, AsyncProp.bytesMAX)
End Sub
Private Sub UserControl_Initialize()
    On Error Resume Next
    With UserControl
        .Height = .picICON.Height
        .Width = .picICON.Width
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    fNAME = PropBag.ReadProperty("OutputFile", "")
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    With UserControl
        .Height = .picICON.Height
        .Width = .picICON.Width
    End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "OutputFile", fNAME, ""
End Sub



