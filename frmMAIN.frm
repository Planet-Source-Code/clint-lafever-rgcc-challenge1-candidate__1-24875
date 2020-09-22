VERSION 5.00
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hit Count"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin ProgramHitCount.webREAD webFILE 
      Left            =   120
      Top             =   480
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Label lblCOUNT 
      Alignment       =   2  'Center
      Caption         =   "This program has been run X time(s) by all the users in the world."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' For a compiled version of this control, you can visit http://lafever.iscool.net
'------------------------------------------------------------
Option Explicit
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Sub Form_Load()
    On Error Resume Next
    With Me.webFILE
        .OutputFile = App.Path & "\hitcount.ini"
        .FileURL = LoadResString(101)
    End With
End Sub
Private Sub webFILE_DownloadComplete()
    On Error Resume Next
    Dim x As Long
    x = GetPrivateProfileInt("HITCOUNT", "COUNT", 1, Me.webFILE.OutputFile)
    Me.lblCOUNT.Caption = "This program has been run " & x & " time(s) by all the users in the world."
    Kill Me.webFILE.OutputFile
End Sub
Private Sub webFILE_DownloadError(errNUM As Long, errSTR As String)
    On Error Resume Next
    MsgBox errNUM & ":" & errSTR
End Sub
Private Sub webFILE_DownloadProgress(bytesREAD As Long, bytesMAX As Long)
    On Error Resume Next
    Me.lblCOUNT.Caption = "Accessing server..."
End Sub
'The logic of this code is to use a custom control that will access an ASP page on a webserver.
'That ASP page is coded to connect to a database on the server and query/update a hitcount
'table and then format the results to an INI file format.  The user control then reads the resulting
'page and saves it to the local hard drive as an .INI file.  Next it uses the GetPrivateProfileInt API
'to read the hit count value from the INI and then display the value.
