VERSION 5.00
Begin VB.Form frmRepInt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replicate interface"
   ClientHeight    =   7428
   ClientLeft      =   2880
   ClientTop       =   2112
   ClientWidth     =   6192
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7428
   ScaleWidth      =   6192
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAsObjectUnknownClasses 
      Caption         =   "Use ""As Object"" for external Classes"
      Enabled         =   0   'False
      Height          =   300
      Left            =   552
      TabIndex        =   7
      Top             =   6264
      Width           =   4980
   End
   Begin VB.ComboBox cboEnums 
      Enabled         =   0   'False
      Height          =   336
      ItemData        =   "frmRepInt.frx":0000
      Left            =   552
      List            =   "frmRepInt.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5808
      Width           =   5460
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   444
      Left            =   4512
      TabIndex        =   4
      Top             =   6744
      Width           =   1308
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Save *.ctl file"
      Height          =   444
      Left            =   504
      TabIndex        =   3
      Top             =   6744
      Width           =   3180
   End
   Begin VB.OptionButton optReplicate 
      Caption         =   "Just replicate the interface"
      Height          =   564
      Left            =   552
      TabIndex        =   2
      Top             =   4896
      Width           =   5000
   End
   Begin VB.OptionButton optEncapsulate 
      Caption         =   "Encapsulate the original control in the UserControl"
      Height          =   564
      Left            =   552
      TabIndex        =   1
      Top             =   4320
      Width           =   5000
   End
   Begin RepInt.Rep Rep1 
      Height          =   3732
      Left            =   216
      TabIndex        =   0
      Top             =   240
      Width           =   5772
      _ExtentX        =   10181
      _ExtentY        =   6583
   End
   Begin VB.Label lblEnums 
      Caption         =   "Enums:"
      Enabled         =   0   'False
      Height          =   420
      Left            =   552
      TabIndex        =   5
      Top             =   5520
      Width           =   684
   End
End
Attribute VB_Name = "frmRepInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    Dim iDlg As New CDlg
    
    If (Not optEncapsulate.Value) And (Not optReplicate.Value) Then
        MsgBox "You need to select whteher to include code for the original control inside the UserControl or not", vbExclamation
        Exit Sub
    End If
    
    If Rep1.GetControlTypeName = "" Then Exit Sub
    Rep1.Encapsulate = optEncapsulate.Value
    Rep1.EnumsTreatment = cboEnums.ListIndex
    Rep1.AsObjectUnknownClasses = (chkAsObjectUnknownClasses.Value = 1)
    
    iDlg.Filter = "UserControl files (*.ctl)|*.ctl"
    iDlg.FilterIndex = 0
    iDlg.FileName = "My" & Rep1.GetControlTypeName & ".ctl"
    iDlg.ShowSave
    If Not iDlg.Canceled Then
        If FileExists(iDlg.FileName) Then
            If MsgBox("File already exists, overwrite?", vbYesNo Or vbQuestion) = vbNo Then
                MsgBox "Canceled.", vbInformation
                Exit Sub
            Else
                Kill iDlg.FileName
                If FileExists(iDlg.FileName) Then
                    MsgBox "File is locked, cannot be replaced.", vbCritical
                    Exit Sub
                End If
            End If
        End If
        SaveTextFile iDlg.FileName, Rep1.GetText
    End If
End Sub

Private Sub SaveTextFile(nPath As String, nText As String)
    Dim iFreeFile
    
    If nText = "" Then
        MsgBox "Canceled.", vbInformation
        Exit Sub
    End If
    
    If FileExists(nPath) Then
        Kill nPath
        If FileExists(nPath) Then
            MsgBox "File is locked, cannot be replaced.", vbCritical
            Exit Sub
        End If
    End If
    
    'On Error Resume Next
    iFreeFile = FreeFile
    Open nPath For Output As #iFreeFile
    Print #iFreeFile, nText
    Close #iFreeFile
End Sub

Private Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

    On Error Resume Next
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum
    FileExists = (Err.Number = 0) Or (Err.Number = 70) Or (Err.Number = 55)
    Close intFileNum
    Err.Clear
End Function

Private Sub Form_Load()
    If Not InIDE Then
        MsgBox "This project needs to be run in the IDE (uncompiled).", vbCritical
        End
    End If
    cboEnums.ListIndex = 0
End Sub

Private Function InIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number Then
        InIDE = True
    End If
End Function

Private Sub optEncapsulate_Click()
    lblEnums.Enabled = False
    cboEnums.Enabled = False
    chkAsObjectUnknownClasses.Enabled = False
End Sub

Private Sub optReplicate_Click()
    lblEnums.Enabled = True
    cboEnums.Enabled = True
    chkAsObjectUnknownClasses.Enabled = True
End Sub
