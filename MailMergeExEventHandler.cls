VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailMergeExEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private GroupBy_Fields As Collection ' by group name of Collection
Private PreviousGroupBy As Collection ' by field name
Private DocData As Collection ' by doc index of (Collection by field name of (Collection of string values))
Private CurrentDoc As Collection ' by field name of (Collection of string values)

Dim handler As New MailMergeExHandler

Private Sub App_MailMergeBeforeMerge(ByVal Doc As Document, ByVal StartRecord As Long, ByVal EndRecord As Long, Cancel As Boolean)
        
    Set handler.Doc = Doc
    
    ' Call with error handling
    On Error GoTo No_GroupBy_Fields
    Set GroupBy_Fields = handler.FieldByGroups("")
    Err.Clear
    
    Set PreviousGroupBy = New Collection
    Dim data As MailMergeDataField
    For Each data In GroupBy_Fields
        PreviousGroupBy.Add "", data.Name
    Next
    
    Set DocData = New Collection
    
    handler.InstallTags
    Exit Sub
    
No_GroupBy_Fields:
    MsgBox "Aucun champs de regroupement"
        
End Sub

Private Sub App_MailMergeBeforeRecordMerge(ByVal Doc As Document, Cancel As Boolean)
    
    Dim handler As New MailMergeExHandler
    Set handler.Doc = Doc
      
    Dim current As New Collection
    Dim different As Boolean
    
    ' Compare each group-by field s(those in group "") to the one of previous document
    ' to check if the current document is in the same group or not.
    Dim data As MailMergeDataField
    For Each data In GroupBy_Fields
        current.Add data.Value, data.Name
        If data.Value <> PreviousGroupBy(data.Name) Then
            different = True
        End If
    Next
        
    If different Then
        ' Init all fields of the current document
        Set CurrentDoc = New Collection
        
        For Each data In Doc.MailMerge.DataSource.DataFields
            Dim valueSet As Collection
            Set valueSet = New Collection
            valueSet.Add data.Value
            CurrentDoc.Add valueSet, data.Name
        Next
        
        DocData.Add CurrentDoc
    Else
        Cancel = True
        ' Add values to other groups
        
        For Each data In Doc.MailMerge.DataSource.DataFields
            CurrentDoc(data.Name).Add data.Value
        Next
    End If
  
    Set PreviousGroupBy = current
End Sub

Private Sub App_MailMergeAfterMerge(ByVal Doc As Document, ByVal DocResult As Document)
    
    Set handler.Doc = Doc
    
    handler.RemoveTags
    
    
    Dim resultHandler As New MailMergeExResultHandler
    Set resultHandler.Doc = DocResult
    
    'resultHandler.PopulateMerged
    
End Sub



