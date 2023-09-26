Sub MrgCll()
Dim Cll As Range, Temp As String
On Error Resume Next
If Selection.MergeCells = False Then
For Each Cll In Selection
If Cll <> "" Then Temp = Temp + Cll.Text + " "
Next Cll
Selection.Merge
Selection.Value = Left(Temp, Len(Temp) - 1)
Else
Selection.UnMerge
End If
Selection.HorizontalAlignment = xlCenter
Selection.VerticalAlignment = xlCenter
End Sub
