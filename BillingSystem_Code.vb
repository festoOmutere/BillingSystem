/*Code written and tested by Festo Omutere on 14/04/2023*/

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
'Clear all form fields
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox3.Value = ""
    Me.TextBox4.Value = ""
    Me.TextBox5.Value = ""
    Me.TextBox6.Value = ""
    Me.TextBox7.Value = ""
    Me.TextBox8.Value = ""
    Me.TextBox9.Value = ""
    Me.TextBox10.Value = ""
    Me.TextBox11.Value = ""
    Me.TextBox12.Value = ""
    Me.TextBox13.Value = ""
    Me.TextBox14.Value = ""
    Me.TextBox15.Value = ""
    'Clear the listbox
    ListBox1.Clear
End Sub

Private Sub CommandButton2_Click()
'Confirm exit message
Dim iExit As VbMsgBoxResult
iExit = MsgBox("Confirm if you want to exit", vbQuestion + vbYesNo, "Customer Billing System")
'if yes option is selected, the program is exited
If iExit = vbYes Then
Unload Me
End If
End Sub

Private Sub CommandButton3_Click()
Dim wks As Worksheet ' set wks as worksheet
Dim AddNew As Range  ' Set AddNew are a range
Set wks = Sheet2     ' select sheet1 as the worksheet
Set AddNew = wks.Range("A65326").End(xlUp).Offset(1, 0) 'This line is a command to add a new record/input data in rows in the workbook after when button is executed
'List of data input in the workbook
AddNew.Offset(0, 0).Value = TextBox1.Text * 12
AddNew.Offset(0, 1).Value = TextBox2.Text * 12
AddNew.Offset(0, 2).Value = TextBox3.Text * 14
AddNew.Offset(0, 3).Value = TextBox4.Text * 15
AddNew.Offset(0, 4).Value = TextBox5.Text * 16
AddNew.Offset(0, 5).Value = TextBox6.Text * 16
AddNew.Offset(0, 6).Value = TextBox7.Text * 18
AddNew.Offset(0, 7).Value = TextBox8.Text * 19
AddNew.Offset(0, 8).Value = TextBox9.Text * 9
AddNew.Offset(0, 9).Value = TextBox10.Text * 9
AddNew.Offset(0, 10).Value = TextBox11.Text * 13
AddNew.Offset(0, 11).Value = TextBox12.Text * 13
AddNew.Offset(0, 12).Value = TextBox13.Text
AddNew.Offset(0, 13).Value = TextBox14.Text
AddNew.Offset(0, 14).Value = TextBox15.Text


End Sub

Private Sub CommandButton4_Click()
'Button to execute calculations of the bill
'Set blank input textbox to zero value to avoid errors in calculations
Dim txts
For Each txts In Frame1.Controls
  If TypeOf txts Is msforms.TextBox Then
     If txts.Text = "" Then
        txts.Text = "0"
          End If
          End If
    Next txts
  'Calculate the cost of selected items
Dim Cake(15) As Double

Cake(0) = TextBox1.Text * 12
Cake(1) = TextBox2.Text * 12
Cake(2) = TextBox3.Text * 14
Cake(3) = TextBox4.Text * 15
Cake(4) = TextBox5.Text * 16
Cake(5) = TextBox6.Text * 16
Cake(6) = TextBox7.Text * 18
Cake(7) = TextBox8.Text * 19
Cake(8) = TextBox9.Text * 9
Cake(9) = TextBox10.Text * 9
Cake(10) = TextBox11.Text * 13
Cake(11) = TextBox12.Text * 13
'Calculate subtotal
Cake(12) = Cake(0) + Cake(1) + Cake(2) + Cake(3) + Cake(4) + Cake(5) + Cake(6) + Cake(7) + Cake(8) + Cake(9) + Cake(10) + Cake(11)
'calculate the tax amount
Cake(13) = Cake(12) * 0.15
TextBox14.Text = "KES." & Cake(13)
TextBox13.Text = "KES." & Cake(12)
'Calculate the grand total amount
 TextBox15.Text = "KES." & (Cake(12) + Cake(13))
 
 Dim wks As Worksheet
Dim AddNew As Range
Set wks = Sheet1
Set AddNew = wks.Range("B2").End(xlUp).Offset(1, 0)
AddNew.Offset(0, 0).Value = TextBox1.Text * 12
AddNew.Offset(0, 1).Value = TextBox2.Text * 12
AddNew.Offset(0, 2).Value = TextBox3.Text * 14
AddNew.Offset(0, 3).Value = TextBox4.Text * 15
AddNew.Offset(0, 4).Value = TextBox5.Text * 16
AddNew.Offset(0, 5).Value = TextBox6.Text * 16
AddNew.Offset(0, 6).Value = TextBox7.Text * 18
AddNew.Offset(0, 7).Value = TextBox8.Text * 19
AddNew.Offset(0, 8).Value = TextBox9.Text * 9
AddNew.Offset(0, 9).Value = TextBox10.Text * 9
AddNew.Offset(0, 10).Value = TextBox11.Text * 13
AddNew.Offset(0, 11).Value = TextBox12.Text * 13
AddNew.Offset(0, 12).Value = TextBox13.Text
AddNew.Offset(0, 13).Value = TextBox14.Text
AddNew.Offset(0, 14).Value = TextBox15.Text

Dim rng As Range
Dim tbl As ListObject
Dim rngRow As Range
Set tbl = wks.ListObjects("Table2")
Set rng = tbl.Range
rng.AutoFilter Field:=1, Criteria1:="<>"
rng.AutoFilter Field:=2, Criteria1:="<>"
'Set rng = rng.Offset(1).Resize(rng.Rows.Count - 1)
Set rng = rng.SpecialCells(xlCellTypeVisible)
'Fill list box
For Each rngRow In rng.Rows
    ListBox1.ColumnCount = 2
    ListBox1.AddItem rngRow.Cells(1, 1) & vbTab & rngRow.Cells(1, 2)
    
Next rngRow

End Sub


Private Sub Reset_Click()

End Sub

Private Sub UserForm_Initialize()
Application.WindowState = xlMaximized

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Application.Visible = True
Unload Me
End Sub
