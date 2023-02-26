Attribute VB_Name = "Filtering_and_saving"
Sub Filtering_Saving()
Attribute Filtering_Saving.VB_ProcData.VB_Invoke_Func = " \n14"
  
    
    ' declaring variables type
    Dim final_path As String
    
    ' filtering and saving a file
    Sheets("invoices").Select
    
    For Each User In Range("user_list")
        
        If User = "" Then Exit For ' exit loop if cell does not contain any string, otherwise continue
        final_path = Range("path_filtered_data") & "\" & User & ".xlsx"
        ActiveSheet.ListObjects("invoices").Range.AutoFilter Field:=5, Criteria1:=User
        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Workbooks.Add
        ActiveSheet.Paste
        Range("A1").Select
        ActiveWorkbook.SaveAs Filename:=final_path
        ActiveWindow.Close
        ActiveSheet.ListObjects("invoices").Range.AutoFilter Field:=5
        Range("A1").Select
        
        Next User
    
    Sheets("main").Select
    Range("A1").Select
    
End Sub

