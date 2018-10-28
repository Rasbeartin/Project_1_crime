Attribute VB_Name = "Module2"
Sub Deltxtleft()
Dim rng As Range
Dim sizeleft As Integer
Dim sizeright As Integer
Dim UserRange As Range, DefaultRange As String, UserRange2 As Range, DefaultRange2 As String
Dim ReplaceRange As String

sizeleft = InputBox("How many characters do you want to delete from the left?")
DefaultRange = Selection.Address

sizeright = InputBox("How many characters do you want to delete from the right?")
DefaultRange2 = Selection.Address

Set UserRange = Application.InputBox _
    (Prompt:="Range to search: ", _
    Title:="Search", _
    Default:=DefaultRange, _
    Type:=8)

Set UserRange2 = Application.InputBox _
    (Prompt:="Range to search: ", _
    Title:="Search", _
    Default:=DefaultRange2, _
    Type:=8)
    
UserRange.Select

    For Each UserRange In UserRange.Cells
        UserRange.Value = Right(UserRange, Len(UserRange) - sizeleft)
    Next
        Set UserRange = Nothing
        On Error GoTo 0
    
    For Each UserRange In UserRange2.Cells
        UserRange.Value = Left(UserRange, Len(UserRange) - sizeright)
    Next
        Set UserRange = Nothing
        On Error GoTo 0
        
End Sub


