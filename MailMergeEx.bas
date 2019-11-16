Attribute VB_Name = "MailMergeEx"
Option Explicit

Dim EventClassModule As New MailMergeExEventHandler

Public Tag_StartOfDocument As String
Public Tag_FieldStart As String
Public Tag_FieldEnd As String
Public Tag_FieldNameDelimiter As String
Public Tag_ValueDelimiter As String
Public Tag_GroupDelimiter As String

Sub AutoOpen()
    '  Tags for tests
    'Tag_StartOfDocument = "++++sof++++"
    'Tag_FieldStart = "++++start++++"
    'Tag_FieldEnd = "++++end++++"
    'Tag_FieldNameDelimiter = "@@@@"
        
    '  Tags for prod
    Tag_StartOfDocument = Chr(1)  ' SOH
    Tag_FieldStart = Chr(2) ' STX
    Tag_FieldEnd = Chr(3)   ' ETX
    Tag_FieldNameDelimiter = Chr(26) 'SUB
    
    Tag_ValueDelimiter = Chr(10)
    Tag_GroupDelimiter = "_"
    
    Register_Event_Handler
End Sub

Sub Register_Event_Handler()
    Set EventClassModule.App = Word.Application
End Sub


Public Function GroupName(str As String) As String
    Dim Items
    Items = Split(str, MailMergeEx.Tag_GroupDelimiter, 2)
    If UBound(Items) = 0 Then
        GroupName = ""
    Else
        GroupName = Items(0)
    End If
End Function

Public Sub AddToGroup(Groups As Collection, Key As String, Item As Object)
    Dim GroupItems As Collection
    On Error Resume Next
    Set GroupItems = Groups(Key)
    If GroupItems Is Nothing Then
        Set GroupItems = New Collection
        Groups.Add GroupItems, Key
    End If
    Err.Clear
    GroupItems.Add Item
End Sub

