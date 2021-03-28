VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public ReturnTypeName As String
Public ReturnTypeName2 As String
Public HasGet As Boolean
Public HasLet As Boolean
Public HasSet As Boolean
Public HelpString As String
Public MemberFlags As Long
Public MemberId As Long
Public ReturnTypeObject As Boolean
Public ReturnTypeLong As Boolean
Public ReturnTypeIsNumeric As Boolean
Public Persistent As Boolean
Private mParameters() As cParameter

Private Sub Class_Initialize()
    ReDim mParameters(0)
End Sub

Public Sub AddParameter(ByVal nName As String, ByVal nTypeName As String, ByVal nTypeName2 As String, ByVal nTypeLong As Boolean, ByVal nTypeObject As Boolean, ByVal nIsByVal As Boolean, ByVal nOptional As Boolean, nDefaultValue As String, nLibraryName As String)
    Dim iParam As cParameter
    
    If nLibraryName = "VB" Then
        If (nName = "Index") And (nTypeName = "Variant") Then nOptional = True
    End If
    
    ReDim Preserve mParameters(UBound(mParameters) + 1)
    Set iParam = New cParameter
    iParam.Name = nName
    iParam.TypeName = nTypeName
    iParam.TypeName2 = nTypeName2
    iParam.TypeLong = nTypeLong
    iParam.TypeObject = nTypeObject
    iParam.IsByVal = nIsByVal
    iParam.IsOptional = nOptional
    iParam.DefaultValue = nDefaultValue
    Set mParameters(UBound(mParameters)) = iParam
End Sub

Public Property Get Parameters(nIndex As Long) As cParameter
    Set Parameters = mParameters(nIndex)
End Property

Public Property Get ParamCount() As Long
    ParamCount = UBound(mParameters)
End Property

Public Function ParamExists(ByVal nName As String) As Boolean
    Dim c As Long
    
    nName = LCase$(nName)
    For c = 1 To UBound(mParameters)
        If LCase$(mParameters(c).Name) = nName Then
            ParamExists = True
            Exit Function
        End If
    Next
End Function

Public Function Clone() As cMember
    Dim c As Long
    Dim iPar As cParameter
    
    Set Clone = New cMember
    
    Clone.Name = Name
    Clone.ReturnTypeName = ReturnTypeName
    Clone.ReturnTypeName2 = ReturnTypeName2
    Clone.HasGet = HasGet
    Clone.HasLet = HasLet
    Clone.HasSet = HasSet
    Clone.HelpString = HelpString
    Clone.MemberFlags = MemberFlags
    Clone.MemberId = MemberId
    Clone.ReturnTypeObject = ReturnTypeObject
    Clone.ReturnTypeLong = ReturnTypeLong
    Clone.ReturnTypeIsNumeric = ReturnTypeIsNumeric
    Clone.Persistent = Persistent
    
    For c = 1 To ParamCount
        Set iPar = mParameters(c)
        Clone.AddParameter iPar.Name, iPar.TypeName, iPar.TypeName2, iPar.TypeLong, iPar.TypeObject, iPar.IsByVal, iPar.IsOptional, iPar.DefaultValue, ""
    Next
End Function

Public Function GetParam(ByVal nName As String) As cParameter
    Dim c As Long
    
    nName = LCase$(nName)
    For c = 1 To UBound(mParameters)
        If LCase$(mParameters(c).Name) = nName Then
            Set GetParam = mParameters(c)
            Exit Function
        End If
    Next
End Function

Public Function GetParamIndex(ByVal nName As String) As Long
    Dim c As Long
    
    nName = LCase$(nName)
    For c = 1 To UBound(mParameters)
        If LCase$(mParameters(c).Name) = nName Then
            GetParamIndex = c
            Exit Function
        End If
    Next
End Function


