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

Sub Class_Initialize()
    Tag_StartOfDocument = MailMergeEx.Tag_StartOfDocument
    Tag_FieldStart = MailMergeEx.Tag_FieldStart
    Tag_FieldEnd = MailMergeEx.Tag_FieldEnd
    Tag_FieldNameDelimiter = MailMergeEx.Tag_FieldNameDelimiter
    Tag_ValueDelimiter = MailMergeEx.Tag_ValueDelimiter
End Sub

Public Sub PopulateMerged(DocData As Collection) ' by doc index of (Collection by field name of (Collection of string values))
    
    Dim docIdx As Integer
    docIdx = 1
    
    ' Foreach sub-document
    Dim Subdoc As Range
    Set Subdoc = SubDocAfter(StartOfDocumentRange)
    Do Until Subdoc Is Nothing
        Dim DocInfo As Collection
        Set DocInfo = DocData(docIdx)
        
        Dim Analysis As Collection
        Set Analysis = AnalyzeSubDoc(Subdoc, DocData(docIdx))
        
        Dim FieldGroup As Collection
        Dim Field As FieldInfo
        
        For Each FieldGroup In Analysis
            ' Complex substitution
            Dim r As Range
            Set r = Doc.Range(FieldGroup(1).Range.Start, FieldGroup(FieldGroup.Count).Range.End)
            Set r = Doc.Range(r.Paragraphs.First.Range.Start, r.Paragraphs.Last.Range.End)
                            
            If Asc(Right(r, 1)) <> 13 Then ' Special case when not standard paragraph (i.e. table)
                r.End = r.End + 1
            End If
            
            r.Copy
            r.Collapse
            ' Replace in reverse order
            ReplaceFieldWithValues FieldGroup, FieldGroup(1).Values.Count, r
            
            Dim i As Integer
            For i = FieldGroup(1).Values.Count - 1 To 1 Step -1
                r.Paste
                ReplaceFieldWithValues FieldGroup, i, r
            Next
        Next
                
        Set Subdoc = SubDocAfter(Subdoc)
        docIdx = docIdx + 1
    Loop
    
    RemoveStartOfDocumentTags
    
End Sub

Sub RemoveStartOfDocumentTags()
    Doc.Range.Find.Execute Tag_StartOfDocument, ReplaceWith:="", Replace:=wdReplaceAll
End Sub

Sub ReplaceFieldWithValues(FieldGroup As Collection, idx As Integer, r As Range)
    Dim f As FieldInfo
    For Each f In FieldGroup
        Dim r2 As Range
        Set r2 = r.Duplicate
        r2.Find.Execute Tag_FieldStart & f.Name & Tag_FieldNameDelimiter & "*" & Tag_FieldEnd, _
                MatchWildcards:=True
        r2.Text = f.Values(idx)
    Next
End Sub

Function AnalyzeSubDoc(Subdoc As Range, DocInfo As Collection) As Collection
    
    Dim r As Range
    Dim Field As Range
    
    Dim Items As New Collection
    
    Set r = Subdoc.Duplicate
    
    ' While finding field
    Do While r.Find.Execute(Tag_FieldStart & "*" & Tag_FieldEnd, MatchWildcards:=True)
        ' that are in the current document
        If r.Start > Subdoc.End Then
            Exit Do
        End If
        
        Items.Add AnalyseField(r, DocInfo)
    Loop
    
    ' Group by non-empty GroupName
    Dim GroupedItems As New Collection
    
    Dim Item As FieldInfo
    Dim CurrentGroup As Collection
    Dim CurrentGroupName As String
    CurrentGroupName = ""
    For Each Item In Items
        If Item.GroupName <> CurrentGroupName Or Item.GroupName = "" Then
            Set CurrentGroup = New Collection
            GroupedItems.Add CurrentGroup
            CurrentGroupName = Item.GroupName
        End If
        CurrentGroup.Add Item
    Next
            
    Set AnalyzeSubDoc = GroupedItems
    
End Function

Function AnalyseField(Field As Range, DocInfo As Collection) As FieldInfo
    Dim Item As New FieldInfo
    
    Set Item.Range = Field.Duplicate
    
    ' Cut start and end tags
    Dim r As Range
    Set r = Field.Duplicate
    r.Start = r.Start + Len(Tag_FieldStart)
    r.End = r.End - Len(Tag_FieldEnd)
        
    ' Find name delimiter
    Dim NameDelimiterRange As Range
    Set NameDelimiterRange = r.Duplicate
    NameDelimiterRange.Find.Execute Tag_FieldNameDelimiter

    ' Extract name
    Dim NameRange As Range
    Set NameRange = Doc.Range(r.Start, NameDelimiterRange.Start)
        
    Item.Name = NameRange.Text
    
    Item.GroupName = MailMergeEx.GroupName(Item.Name)
    
    Set Item.Values = DocInfo(Item.GroupName)(Item.Name)
        
    Set AnalyseField = Item
End Function

Function SubDocAfter(Location As Range)
    Dim Sod As Range
    Dim NextSod As Range
    Dim Found As Boolean
    
    ' If at the end of the document
    If Location.End >= Doc.Range.End Then
        Set SubDocAfter = Nothing
        Exit Function
    End If
    
    ' Searching start tag
    Set Sod = Doc.Range(Location.End, Doc.Range.End)
    Found = Sod.Find.Execute(Tag_StartOfDocument)
    If Not Found Then
        Set SubDocAfter = Nothing
        Exit Function
    End If
    
    ' Searching end tag
    Set NextSod = Doc.Range(Sod.End, Doc.Range.End)
    Found = NextSod.Find.Execute(Tag_StartOfDocument)
    
    If Found Then ' Standard case
        Set SubDocAfter = Doc.Range(Sod.Start, NextSod.Start)
    Else ' When at end of document
        Set SubDocAfter = Doc.Range(Sod.Start, Doc.Range.End)
    End If
        
End Function

' ----------------

Function StartOfDocumentRange() As Range
    Set StartOfDocumentRange = Doc.Range(Doc.Range.Start, Doc.Range.Start)
End Function
