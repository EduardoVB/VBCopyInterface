VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties persistance"
   ClientHeight    =   8976
   ClientLeft      =   2916
   ClientTop       =   2148
   ClientWidth     =   4488
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8976
   ScaleWidth      =   4488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   1176
      TabIndex        =   3
      Top             =   8304
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   2784
      TabIndex        =   2
      Top             =   8304
      Width           =   1260
   End
   Begin VB.ListBox lstProperties 
      Height          =   7176
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   816
      Width           =   4260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please uncheck what properties are not persistent (not available at design time)."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   216
      TabIndex        =   0
      Top             =   216
      Width           =   4068
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ControlTypeName As String
Public PressedOK As Boolean

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    Dim c As Long
    
    For c = 0 To lstProperties.ListCount - 1
        SaveSetting App.Title, "Control_" & ControlTypeName, lstProperties.List(c), CLng(lstProperties.Selected(c))
    Next
    
    PressedOK = True
    Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = 1
        Hide
    End If
End Sub

Public Sub ShowForm()
    Dim c As Long
    
    For c = 0 To lstProperties.ListCount - 1
        lstProperties.Selected(c) = CBool(Val(GetSetting(App.Title, "Control_" & ControlTypeName, lstProperties.List(c), "-1")))
    Next
    Show vbModal
End Sub

Public Function GetSelectionByPropertyName(nPropName As String) As Boolean
    Dim c As Long
    
    For c = 0 To lstProperties.ListCount - 1
        If lstProperties.List(c) = nPropName Then
            GetSelectionByPropertyName = lstProperties.Selected(c)
            Exit Function
        End If
    Next
End Function
    
