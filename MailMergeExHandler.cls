VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailMergeExHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Doc As Document

Dim Tag_StartOfDocument As String
Dim Tag_FieldStart As String
Dim Tag_FieldEnd As String
Dim Tag_FieldNameDelimiter As String

Sub Class_Initialize()
    Tag_StartOfDocument = MailMergeEx.Tag_StartOfDocument
    Tag_FieldStart = MailMergeEx.Tag_FieldStart
    Tag_FieldEnd = MailMergeEx.Tag_FieldEnd
    Tag_FieldNameDelimiter = MailMergeEx.Tag_FieldNameDelimiter
End Sub

Sub InstallTags()
    AddStartOfDocumentTag
   
    Dim f As field
    
    For Each f In Doc.Fields
        AddFieldTags f
    Next
End Sub

Sub RemoveTags()
    RemoveStartOfDocumentTag
    
    Dim f As field
    
    For Each f In Doc.Fields
        RemoveFieldTags f
    Next
End Sub

Sub AddFieldTags(f As field)
    Dim fieldCommand As String
    fieldCommand = Trim(f.Code.Text)

    If f.Type = wdFieldMergeField And fieldCommand Like "MERGEFIELD *" Then
        Dim Name As String
        Dim r As Range
        
        Name = Trim(Split(fieldCommand, " ", 2)(1))
        Set r = FieldRange(f)
        ' insert text in reverse order
        r.InsertBefore Tag_FieldNameDelimiter
        r.InsertBefore Name
        r.InsertBefore Tag_FieldStart
        r.InsertAfter Tag_FieldEnd
    End If
End Sub

Sub RemoveFieldTags(f As field)
    Dim fieldCommand As String
    fieldCommand = Trim(f.Code.Text)

    If f.Type = wdFieldMergeField And fieldCommand Like "MERGEFIELD *" Then
        Dim Name As String
        Dim r As Range
        
        Name = Trim(Split(fieldCommand, " ", 2)(1))
        Set r = FieldRange(f)
        ' remove text in reverse order
        SafeRemoveBefore r, Tag_FieldNameDelimiter
        SafeRemoveBefore r, Name
        SafeRemoveBefore r, Tag_FieldStart
        SafeRemoveAfter r, Tag_FieldEnd
    End If
End Sub

Sub AddStartOfDocumentTag()
    StartOfDocumentRange.InsertAfter Tag_StartOfDocument
End Sub

Sub RemoveStartOfDocumentTag()
    SafeRemoveAfter StartOfDocumentRange, Tag_StartOfDocument
End Sub

' -----------------------------------
' SafeRemove means remove that checks that the text to remove does
' exist before removing it.

Sub SafeRemoveBefore(location As Range, str)
    Dim r As Range
    Set r = Doc.Range(location.Start - Len(str), location.Start)
    SafeRemoveAt r, str
End Sub

Sub SafeRemoveAfter(location As Range, str)
    Dim r As Range
    Set r = Doc.Range(location.End, location.End + Len(str))
    SafeRemoveAt r, str
End Sub

Sub SafeRemoveAt(location As Range, str)
    If location.Text <> str Then
        MsgBox location.Text & " not found as " & location.Start
    Else
        location.Delete
    End If
End Sub

Function StartOfDocumentRange() As Range
    Set StartOfDocumentRange = Doc.Range(Doc.Range.Start, Doc.Range.Start)
End Function

Function FieldRange(f As field) As Range
    If f.Result.Text <> "" Then
        Set FieldRange = Doc.Range(f.Code.Start - 1, f.Result.End + 1)
    Else
        Set FieldRange = Doc.Range(f.Code.Start - 1, f.Result.End)
    End If
End Function

Function FieldByGroups() As Collection
    Dim groups As New Collection
    
    Dim data As Object
    For Each data In Doc.MailMerge.DataSource.DataFields
        Dim group As String
        Dim groupItems As Collection
        
        group = MailMergeEx.GroupName(data.Name)
        
        MailMergeEx.AddToGroup groups, group, data
    Next
  
    Set FieldByGroups = groups
End Function
