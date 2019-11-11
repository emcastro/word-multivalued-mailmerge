VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailMergeExResultHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Doc As Document

Public Tag_StartOfDocument As String
Public Tag_FieldStart As String
Public Tag_FieldEnd As String
Public Tag_FieldNameDelimiter As String
Public Tag_NewLineSubstitute As String

Sub Class_Initialize()
    Tag_StartOfDocument = MailMergeEx.Tag_StartOfDocument
    Tag_FieldStart = MailMergeEx.Tag_FieldStart
    Tag_FieldEnd = MailMergeEx.Tag_FieldEnd
    Tag_FieldNameDelimiter = MailMergeEx.Tag_FieldNameDelimiter
    Tag_ValueDelimiter = MailMergeEx.Tag_ValueDelimiter
    Tag_NewLineSubstitute = MailMergeEx.Tag_NewLineSubstitute
End Sub

Public Sub CleanMerged()
    
    ' Foreach sub-document
    Dim subdoc As Range
    Set subdoc = SubDocAfter(StartOfDocumentRange)
    Do Until subdoc Is Nothing
        
        Dim analysis As Collection
        Set analysis = AnalyzeSubDoc(subdoc)
        
        Dim fieldGroup As Collection
        Dim field As FieldInfo
        
        For Each fieldGroup In analysis
            If fieldGroup.Count = 1 Then
                ' Simple substitution
                Set field = fieldGroup(1)
                field.Range.Text = field.Values(0)
            Else
                ' Complex substitution
                Dim r As Range
                Set r = Doc.Range(fieldGroup(1).Range.Start, fieldGroup(fieldGroup.Count).Range.End)
                Set r = Doc.Range(r.Paragraphs.First.Range.Start, r.Paragraphs.Last.Range.End)
                                
                If Asc(Right(r, 1)) <> 13 Then ' Special case when not standard paragraph (i.e. table)
                    r.End = r.End + 1
                End If
                
                r.Copy
                r.Collapse
                ' Replace in reverse order
                ReplaceFieldWithValues fieldGroup, UBound(fieldGroup(1).Values), r
                
                Dim i As Integer
                For i = UBound(fieldGroup(1).Values) - 1 To LBound(fieldGroup(1).Values) Step -1
                    r.Paste
                    ReplaceFieldWithValues fieldGroup, i, r
                Next
            End If
        Next
                
        Set subdoc = SubDocAfter(subdoc)
    Loop
    
    RemoveStartOfDocumentTags
    
End Sub

Sub RemoveStartOfDocumentTags()
    Doc.Range.Find.Execute Tag_StartOfDocument, ReplaceWith:=""
End Sub

Sub ReplaceFieldWithValues(fieldGroup As Collection, idx As Integer, r As Range)
    Dim f As FieldInfo
    For Each f In fieldGroup
        Dim r2 As Range
        Set r2 = r.Duplicate
        r2.Find.Execute Tag_FieldStart & f.Name & Tag_FieldNameDelimiter & "*" & Tag_FieldEnd, _
                MatchWildcards:=True
        r2.Text = f.Values(idx)
    Next
End Sub

Function AnalyzeSubDoc(subdoc As Range) As Collection
    
    Dim r As Range
    Dim field As Range
    
    Dim items As New Collection
    
    Set r = subdoc.Duplicate
    
    ' While finding field
    Do While r.Find.Execute(Tag_FieldStart & "*" & Tag_FieldEnd, MatchWildcards:=True)
        ' that are in the current document
        If r.Start > subdoc.End Then
            Exit Do
        End If
        
        items.Add AnalyseField(r)
    Loop
    
    ' Groupe by non-empty GroupName
    Dim groupedItems As New Collection
    
    Dim item As FieldInfo
    Dim currentGroup As Collection
    Dim currentGroupName As String
    currentGroupName = ""
    For Each item In items
        If item.GroupName <> currentGroupName Or item.GroupName = "" Then
            Set currentGroup = New Collection
            groupedItems.Add currentGroup
            currentGroupName = item.GroupName
        End If
        currentGroup.Add item
    Next
            
    Set AnalyzeSubDoc = groupedItems
    
End Function

Function AnalyseField(field As Range) As FieldInfo
    Dim item As New FieldInfo
    
    Set item.Range = field.Duplicate
    
    ' Cut start and end tags
    Dim r As Range
    Set r = field.Duplicate
    r.Start = r.Start + Len(Tag_FieldStart)
    r.End = r.End - Len(Tag_FieldEnd)
        
    ' Find name delimiter
    Dim nameDelimiterRange As Range
    Set nameDelimiterRange = r.Duplicate
    nameDelimiterRange.Find.Execute Tag_FieldNameDelimiter

    ' Extract name
    Dim nameRange As Range
    Set nameRange = Doc.Range(r.Start, nameDelimiterRange.Start)
        
    item.Name = nameRange.Text
    
    ' Extract values
    Dim valueRange As Range
    Dim valueStrings As Variant
    Set valueRange = Doc.Range(nameDelimiterRange.End, r.End)
    valueStrings = Split(valueRange.Text, Tag_ValueDelimiter)
    
    Dim i As Integer
    For i = LBound(valueStrings) To UBound(valueStrings)
        valueStrings(i) = EscapeFieldValue(valueStrings(i))
    Next
    
    item.Values = valueStrings
    
    item.InitGroupName
        
    Set AnalyseField = item
End Function

Function EscapeFieldValue(ByVal str As String) As String
    EscapeFieldValue = Replace(str, Tag_NewLineSubstitute, Chr(13))
End Function

Function SubDocAfter(location As Range)
    Dim sod As Range
    Dim nextSod As Range
    Dim found As Boolean
    
    ' If at the end of the document
    If location.End >= Doc.Range.End Then
        Set SubDocAfter = Nothing
        Exit Function
    End If
    
    ' Searching start tag
    Set sod = Doc.Range(location.End, Doc.Range.End)
    found = sod.Find.Execute(Tag_StartOfDocument)
    If Not found Then
        Set SubDocAfter = Nothing
        Exit Function
    End If
    
    ' Searching end tag
    Set nextSod = Doc.Range(sod.End, Doc.Range.End)
    found = nextSod.Find.Execute(Tag_StartOfDocument)
    
    If found Then ' Standard case
        Set SubDocAfter = Doc.Range(sod.Start, nextSod.Start)
    Else ' When at end of document
        Set SubDocAfter = Doc.Range(sod.Start, Doc.Range.End)
    End If
        
End Function

' ----------------

Function StartOfDocumentRange() As Range
    Set StartOfDocumentRange = Doc.Range(Doc.Range.Start, Doc.Range.Start)
End Function
