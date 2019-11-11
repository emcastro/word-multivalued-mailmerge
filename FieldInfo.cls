VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Range As Range
Public Values As Variant 'String(*)
Public GroupName As String

Public Function InitGroupName() As String
    Dim items
    items = Split(Name, MailMergeEx.Tag_GroupDelimiter, 2)
    If UBound(items) = 0 Then
        GroupName = ""
    Else
        GroupName = items(0)
    End If
End Function

