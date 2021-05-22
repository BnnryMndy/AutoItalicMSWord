Attribute VB_Name = "NewMacros"

Sub auto_Ctrl_i()
Attribute auto_Ctrl_i.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.auto_Ctrl_i"
ActiveDocument.Content.Select
Do
  With Selection.Find
      .Text = "([A-Za-z])"
      .MatchWildcards = True
      .Forward = True
      .Execute
  End With
 
  If Selection.Find.Found Then
        Selection.Range.Font.Italic = True
  Else
      Exit Do
  End If
Loop

End Sub


