Attribute VB_Name = "CCS_Edit"
Sub ChangeControlEdit()
Application.ScreenUpdating = False
With ActiveWindow.View.RevisionsFilter
        .Markup = wdRevisionsMarkupAll
        .View = wdRevisionsViewFinal
End With
ActiveDocument.TrackRevisions = True

Dim SearchedWord1 As String
Dim SearchedWord2 As String
Dim SearchedWord3 As String
Dim CancelOrNot As Integer
Dim DocRange As Range
Dim DocRange2 As Range
Dim DocRange3 As Range
Dim ParaText As Range
Dim ParaText2 As Range
Dim WasFound As Boolean

SearchedWord1 = "(Internal): "
SearchedWord2 = "(Public): "
SearchedWord3 = "not to be posted"
CancelOrNot = -1
WasFound = False

Do While CancelOrNot = -1
CancelOrNot = MsgBox("Searched keywords include: '" & SearchedWord1 & "', '" & SearchedWord2 & "', '" & SearchedWord3 & "'", _
vbOKCancel, "Press OK to continue...")
'Ok = 1, Cancel = 2
If CancelOrNot = 2 Then Exit Sub
Loop

'find and remove the paragraph containing search word 1 '(Internal): '
Set DocRange = ActiveDocument.Range
With DocRange.Find
.ClearFormatting
.Text = SearchedWord1
.Forward = True
.Wrap = wdFindStop
.Format = False
.MatchCase = False
.MatchWholeWord = True
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False

Do While .Execute
WasFound = True
Set ParaText = DocRange.Paragraphs(1).Range
ParaText.Delete
DocRange.SetRange DocRange.Paragraphs(1).Range.End, _
ActiveDocument.Range.End
Loop
End With

'find and remove search word 2 '(Public): '
Set DocRange2 = ActiveDocument.Range
With DocRange2.Find
  .ClearFormatting
  .Text = SearchedWord2
  .Replacement.Text = ""
  .Forward = True
  .Wrap = wdFindContinue
  .Format = False
  .MatchCase = False
  .MatchWholeWord = True
  .MatchWildcards = False
  .MatchSoundsLike = False
  .MatchAllWordForms = False
  .Execute Replace:=wdReplaceAll
  WasFound = True
End With

'find and remove the paragraph containing search word 3 'not to be posted'
Set DocRange3 = ActiveDocument.Range
With DocRange3.Find
.ClearFormatting
.Text = SearchedWord3
.Forward = True
.Wrap = wdFindStop
.Format = False
.MatchCase = False
.MatchWholeWord = True
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False

Do While .Execute
WasFound = True
Set ParaText2 = DocRange3.Paragraphs(1).Range

ParaText2.Select
Selection.MoveUp Unit:=wdParagraph, Count:=2, Extend:=wdExtend
Selection.Delete
Selection.MoveDown Unit:=wdParagraph, Count:=5, Extend:=wdExtend
Selection.Delete

DocRange3.SetRange DocRange3.Paragraphs(1).Range.End, _
ActiveDocument.Range.End
Loop

End With

'Alert user on end of script
If Not WasFound Then
  MsgBox "'" & SearchedWord1 & "' or '" & SearchedWord2 & "' was not found in the document.", _
vbExclamation + vbOKOnly, "Words not found"
Else
  MsgBox "Please inspect and reject any unwanted changes, then click 'Accept All Changes' in the 'Review' tab", _
vbExclamation + vbOKOnly, "Accept/Reject changes"
End If

Application.ScreenUpdating = True
End Sub


