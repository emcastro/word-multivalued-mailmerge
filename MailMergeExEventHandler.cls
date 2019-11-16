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

Private FieldByGroups As Collection ' by group name of Collection
Private PreviousGroupBy As Collection ' by field name
Private DocData As Collection ' by doc index of (Collection by field name of (Collection of string values))
Private CurrentDoc As Collection ' by field name of (Collection of string values)

Dim Handler As New MailMergeExHandler

Private Sub App_MailMergeBeforeMerge(ByVal Doc As Document, ByVal StartRecord As Long, ByVal EndRecord As Long, Cancel As Boolean)
        
    Set Handler.Doc = Doc
    Set FieldByGroups = Handler.FieldByGroups
    
    ' Call with error handling
    On Error GoTo No_MainGroup
    FieldByGroups.Item "" ' Test if main group exists
    Err.Clear
    
    ' Init mail merge data athering
    Set PreviousGroupBy = New Collection
    Dim Data As MailMergeDataField
    For Each Data In FieldByGroups("")
        PreviousGroupBy.Add "", Data.Name
    Next
    
    Set DocData = New Collection
    
    ' Add tags to the mail merge document
    Handler.InstallTags
    Exit Sub
    
No_MainGroup:
    MsgBox "Aucun champs de regroupement"
        
End Sub

Private Sub App_MailMergeBeforeRecordMerge(ByVal Doc As Document, Cancel As Boolean)
    
    Dim Handler As New MailMergeExHandler
    Set Handler.Doc = Doc
      
    Dim Current As New Collection
    Dim Different As Boolean
    Dim GroupName As String
    Dim Fields As Collection
    
    ' Compare each group-by field s(those in group "") to the one of previous document
    ' to check if the current document is in the same group or not.
    Dim Data As MailMergeDataField
    For Each Data In FieldByGroups("")
        Current.Add Data.Value, Data.Name
        If Data.Value <> PreviousGroupBy(Data.Name) Then
            Different = True
        End If
    Next
        
    If Different Then
        ' Init groups and fields of the current document
        Set CurrentDoc = New Collection
        
        For Each Fields In FieldByGroups
            Dim Group As Collection
            Set Group = New Collection
        
            For Each Data In Fields
                GroupName = MailMergeEx.GroupName(Data.Name) ' the safe for each iteration
                
                Dim Values As Collection
                Set Values = New Collection
                Values.Add Data.Value
                Group.Add Values, Data.Name
            Next
            CurrentDoc.Add Group, GroupName
        Next
        
        DocData.Add CurrentDoc
    Else
        Cancel = True
    End If
    
    ' Add values to groups
    For Each Fields In FieldByGroups
        For Each Data In Fields
            GroupName = MailMergeEx.GroupName(Data.Name)
            CurrentDoc(GroupName)(Data.Name).Add Data.Value
        Next
    Next
  
    Set PreviousGroupBy = Current
End Sub

Private Sub App_MailMergeAfterMerge(ByVal Doc As Document, ByVal DocResult As Document)
    
    Set Handler.Doc = Doc
    
    ' Clean up the mail merge document
    Handler.RemoveTags
    
    ' Analyse collected mail merge data
    Dim DocInfo As Collection
    For Each DocInfo In DocData
        Dim Group As Collection
        For Each Group In DocInfo
            MailMergeEx.RemoveDuplicates Group
        Next
    Next
    
    Dim ResultHandler As New MailMergeExResultHandler
    Set ResultHandler.Doc = DocResult
    
    'ResultHandler.PopulateMerged DocData
    
End Sub



