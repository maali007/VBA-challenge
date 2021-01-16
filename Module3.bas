Attribute VB_Name = "Module3"
Sub Clear_Summary()

'Loop through all worksheets
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
ws_num = ThisWorkbook.Worksheets.Count

Dim x As Integer

For x = 1 To ws_num
ThisWorkbook.Worksheets(x).Activate
    

Range("J:M").Clear

starting_ws.Activate 'activate the worksheet that was originally active

Next x

End Sub
