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

Dim Handler As New MailMergeExHandler

Private Sub App_MailMergeBeforeMerge(ByVal Doc As Document, ByVal StartRecord As Long, ByVal EndRecord As Long, Cancel As Boolean)
        
    Set Handler.Doc = Doc
    
    ' Call with error handling
    On Error GoTo No_GroupBy_Fields
    Set GroupBy_Fields = Handler.FieldByGroups("")
    Err.Clear
    
    Set PreviousGroupBy = New Collection
    Dim Data As MailMergeDataField
    For Each Data In GroupBy_Fields
        PreviousGroupBy.Add "", Data.Name
    Next
    
    Set DocData = New Collection
    
    Handler.InstallTags
    Exit Sub
    
No_GroupBy_Fields:
    MsgBox "Aucun champs de regroupement"
        
End Sub

Private Sub App_MailMergeBeforeRecordMerge(ByVal Doc As Document, Cancel As Boolean)
    
    Dim Handler As New MailMergeExHandler
    Set Handler.Doc = Doc
      
    Dim Current As New Collection
    Dim Different As Boolean
    
    ' Compare each group-by field s(those in group "") to the one of previous document
    ' to check if the current document is in the same group or not.
    Dim Data As MailMergeDataField
    For Each Data In GroupBy_Fields
        Current.Add Data.Value, Data.Name
        If Data.Value <> PreviousGroupBy(Data.Name) Then
            Different = True
        End If
    Next
        
    If Different Then
        ' Init all fields of the current document
        Set CurrentDoc = New Collection
        
        For Each Data In Doc.MailMerge.DataSource.DataFields
            Dim valueSet As Collection
            Set valueSet = New Collection
            valueSet.Add Data.Value
            CurrentDoc.Add valueSet, Data.Name
        Next
        
        DocData.Add CurrentDoc
    Else
        Cancel = True
        ' Add values to other groups
        
        For Each Data In Doc.MailMerge.DataSource.DataFields
            CurrentDoc(Data.Name).Add Data.Value
        Next
    End If
  
    Set PreviousGroupBy = Current
End Sub

Private Sub App_MailMergeAfterMerge(ByVal Doc As Document, ByVal DocResult As Document)
    
    Set Handler.Doc = Doc
    
    Handler.RemoveTags
    
    
    Dim ResultHandler As New MailMergeExResultHandler
    Set ResultHandler.Doc = DocResult
    
    ResultHandler.PopulateMerged DocData
    
End Sub



