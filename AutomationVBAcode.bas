Attribute VB_Name = "Module1"
Sub Sending20Keys()
Attribute Sending20Keys.VB_ProcData.VB_Invoke_Func = "N\n14"

    Dim ws As Worksheet
    Dim rowNum As Long
    Dim i As Long
    
    Dim val1 As String, val2 As String, val3 As String, val4 As String, val5 As String

    Set ws = ActiveSheet
    
    ' Start from the currently selected row
    rowNum = Selection.Row

    ' Activate Chrome once
    AppActivate "Google Chrome"
    Application.Wait Now + TimeValue("0:00:02")
    
    ' ===== Loop through 20 rows =====
    For i = rowNum To rowNum + 19
        
        ' Read values from columns C, E, F, G, H
        val1 = ws.Cells(i, "C").Value
        val2 = ws.Cells(i, "E").Value
        val3 = ws.Cells(i, "F").Value
        val4 = ws.Cells(i, "G").Value
        val5 = ws.Cells(i, "H").Value

        ' Stop when entire row is empty
        If val1 = "" And val2 = "" And val3 = "" And val4 = "" And val5 = "" Then Exit For

        ' ===== Send values and TABs =====
        Application.SendKeys val1, True
        Application.SendKeys "{TAB}", True
        
        Application.SendKeys val2, True
        Application.SendKeys "{TAB}", True
        
        Application.SendKeys val3, True
        Application.SendKeys "{TAB}", True
        
        Application.SendKeys val4, True
        Application.SendKeys "{TAB}", True
        
        Application.SendKeys val5, True
        Application.SendKeys "{TAB}", True

        ' Small delay to prevent skipping
        Application.Wait Now + TimeValue("0:00:01")
    
    Next i


End Sub

