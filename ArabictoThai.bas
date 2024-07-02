Attribute VB_Name = "ArabictoThai"
Sub arabic2thai()
  For i = 0 To 9
    With Selection.Find
      .Text = Chr(48 + i)
      .Replacement.Text = Chr(240 + i)
      .Wrap = wdFindContinue
    End With
  Selection.Find.Execute Replace:=wdReplaceAll
  Next
End Sub

Sub thai2arabic()
  For i = 0 To 9
    With Selection.Find
      .Text = Chr(240 + i)
      .Replacement.Text = Chr(48 + i)
      .Wrap = wdFindContinue
    End With
  Selection.Find.Execute Replace:=wdReplaceAll
  Next
End Sub
