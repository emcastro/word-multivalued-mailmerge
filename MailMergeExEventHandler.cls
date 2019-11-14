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

Private Sub App_MailMergeBeforeMerge(ByVal Doc As Document, ByVal StartRecord As Long, ByVal EndRecord As Long, Cancel As Boolean)
        
    Dim handler As New MailMergeExHandler
    Set handler.Doc = Doc
    
    handler.InstallTags
End Sub

Private Sub App_MailMergeBeforeRecordMerge(ByVal Doc As Document, Cancel As Boolean)
  
  MsgBox "ooo"
  
  Cancel = True
  
End Sub

Private Sub App_MailMergeAfterMerge(ByVal Doc As Document, ByVal DocResult As Document)
    
    Dim handler As New MailMergeExHandler
    Set handler.Doc = Doc
    
    handler.RemoveTags
    
    
    Dim resultHandler As New MailMergeExResultHandler
    Set resultHandler.Doc = DocResult
    
    'resultHandler.CleanMerged
    
End Sub



