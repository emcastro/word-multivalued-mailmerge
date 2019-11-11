Attribute VB_Name = "MailMergeEx"
Option Explicit

Dim EventClassModule As New MailMergeExEventHandler

Public Tag_StartOfDocument As String
Public Tag_FieldStart As String
Public Tag_FieldEnd As String
Public Tag_FieldNameDelimiter As String
Public Tag_ValueDelimiter As String
Public Tag_GroupDelimiter As String
Public Tag_NewLineSubstitute As String

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
    
    Tag_NewLineSubstitute = "\n"
    Tag_ValueDelimiter = Chr(10)
    Tag_GroupDelimiter = "_"
    
    Register_Event_Handler
End Sub

Sub Register_Event_Handler()
    Set EventClassModule.App = Word.Application
End Sub


