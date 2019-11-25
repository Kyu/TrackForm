VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheckoutForm 
   Caption         =   "Checkout Form"
   ClientHeight    =   6255
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7560
   OleObjectBlob   =   "CheckoutForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CheckoutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim selected_ID As String
Dim cell As Range
Dim update As Boolean
Dim warn_food As Boolean
Dim warn_other As Boolean

Function Total_Food() As Integer

' Current food
Dim FoodC As Integer
' Food entered
Dim FoodU As Integer

' Current food items on file
FoodC = Val(FoodC_Label.Caption)
' Food items to be entered
FoodU = Val(Food_TextBox.Text)

Total_Food = FoodC + FoodU

End Function

Function Total_Hygiene() As Integer

Dim HygC As Integer
Dim HygU As Integer

' Current hygiene items on file
HygC = Val(HygieneC_Label.Caption)
' Hygine items to be entered
HygU = Val(Hygiene_TextBox.Text)

Total_Hygiene = HygC + HygU

End Function

Function Total_Baby() As Integer

Dim BabyC As Integer
Dim BabyU As Integer

' Current Baby items on file
BabyC = Val(BabyC_Label.Caption)
' Baby items to be entered
BabyU = Val(Baby_TextBox.Text)

Total_Baby = BabyC + BabyU

End Function

Function Total_Other() As Integer

Dim OtherC As Integer
Dim OtherU As Integer

' Current other items on file
OtherC = Val(OtherC_Label.Caption)
' Baby items to be entered
OtherU = Val(Other_TextBox.Text)

Total_Other = OtherC + OtherU

End Function

Function Total_Total() As Integer

Dim TotalC As Integer
Dim TotalU As Integer

' Current total items on file
TotalC = Val(TotalC_Label.Caption)
' Total total to be updated
TotalU = Val(Total_TextBox.Text)

Total_Total = TotalC + TotalU

End Function

Function Total_Non_Food() As Integer
Total_Non_Food = Total_Hygiene + Total_Baby + Total_Other
End Function

Sub Display_Total()

Dim FoodU As Integer
Dim HygieneU As Integer
Dim BabyU As Integer
Dim OtherU As Integer

Dim TotalU As Integer

' Sum values entered in texboxes, then display
FoodU = Val(Food_TextBox.Text)
HygieneU = Val(Hygiene_TextBox.Text)
BabyU = Val(Baby_TextBox.Text)
OtherU = Val(Other_TextBox.Text)

TotalU = FoodU + HygieneU + BabyU + OtherU

Total_TextBox.Value = TotalU

End Sub

Sub Display_Current_Stats()
    Dim food As Integer
    Dim hygiene As Integer
    Dim baby As Integer
    Dim other As Integer
    
    Dim total As Integer
    
    ' Get Row Values of ofsetting cell N to the right
    food = cell.Offset(0, 1).Value
    hygiene = cell.Offset(0, 2).Value
    baby = cell.Offset(0, 3).Value
    other = cell.Offset(0, 4).Value
    
    total = cell.Offset(0, 5).Value
    
    ' Display the cell values if they exist, otherwise display 0
    FoodC_Label.Caption = IIf(IsNull(food), "0", food)
    HygieneC_Label.Caption = IIf(IsNull(hygiene), "0", hygiene)
    BabyC_Label.Caption = IIf(IsNull(baby), "0", baby)
    OtherC_Label.Caption = IIf(IsNull(other), "0", other)
    
    TotalC_Label.Caption = IIf(IsNull(total), "0", total)
End Sub

Sub Warn_Over(code As Integer)
' 1 is food, 2 is other
' Display warning label
' Turn fonts red & bold

Warning_Frame.Visible = True
Warning_Label.Visible = True

If code = 1 Then
    FoodU_Label.FontBold = True
    FoodU_Label.ForeColor = RGB(255, 0, 0)
    
    warn_food = True
ElseIf code = 2 Then
    HygieneU_Label.ForeColor = RGB(255, 0, 0)
    HygieneU_Label.FontBold = True
    
    BabyU_Label.ForeColor = RGB(255, 0, 0)
    BabyU_Label.FontBold = True
    
    OtherU_Label.ForeColor = RGB(255, 0, 0)
    OtherU_Label.FontBold = True
    
    warn_other = True
End If

End Sub

Sub End_Warning(code As Integer)
' code = 1 for food,
' 2 for other

' Disable food warning signs
If code = 1 Then
    FoodU_Label.ForeColor = RGB(0, 0, 0)
    FoodU_Label.FontBold = False
    
    warn_food = False
End If

' Disable non food warning signs
If code = 2 Then
    HygieneU_Label.ForeColor = RGB(0, 0, 0)
    HygieneU_Label.FontBold = False
    
    BabyU_Label.ForeColor = RGB(0, 0, 0)
    BabyU_Label.FontBold = False
    
    OtherU_Label.ForeColor = RGB(0, 0, 0)
    OtherU_Label.FontBold = False
    
    warn_other = False
End If

' If both warnings are false, then deactivate warning label
If Not (warn_food Or warn_other) Then
    
    Warning_Frame.Visible = False
    Warning_Label.Visible = False
End If

End Sub

Private Sub Exit_Button_Click()

' Create an exit confirmation box
Dim ExitConf As VbMsgBoxResult
ExitConf = MsgBox("Confirm if you want to exit", vbQuestion + vbYesNo, "Checkout Form")

' If Yes, then unload form
If ExitConf = vbYes Then
    Unload Me
End If

End Sub

Private Sub Food_TextBox_Change()

If update Then
    Dim tot_food As Integer
    
    ' Get total food and update label
    tot_food = Total_Food()
    FoodU_Label.Caption = tot_food
    Call Display_Total
    
    ' If gone over limit, Warn
    ' If not gone over limit but warning is active, disable.
    If tot_food > 15 Then
        Warn_Over (1)
    ElseIf warn_food = True Then
        End_Warning (1)
    End If
End If

End Sub

Private Sub Hygiene_TextBox_Change()

' If Hygiene_TextBox.Value + Baby + Other > 10, warn
If update Then
    Dim tot_hyg As Integer
    
    ' Get total food and update label
    tot_hyg = Total_Hygiene()
    HygieneU_Label.Caption = tot_hyg
    Call Display_Total
    
    ' If gone over limit
    If Total_Non_Food() > 10 Then
        Warn_Over (2)
    ElseIf warn_other = True Then
        End_Warning (2)
    End If
End If

End Sub

Private Sub Baby_TextBox_Change()
      
If update Then
    Dim tot_baby As Integer
    
    ' Get total food and update label
    tot_baby = Total_Baby()
    BabyU_Label.Caption = tot_baby
    Call Display_Total
    
    ' If gone over limit
    If Total_Non_Food() > 10 Then
        Warn_Over (2)
    ElseIf warn_other = True Then
        End_Warning (2)
    End If
End If

End Sub

Private Sub Other_TextBox_Change()

If update Then
    Dim tot_other As Integer
    
    ' Get total food and update label
    tot_other = Total_Other()
    OtherU_Label.Caption = tot_other
    Call Display_Total
    
    ' If gone over limit
    If Total_Non_Food() > 10 Then
        Warn_Over (2)
    ElseIf warn_other = True Then
        End_Warning (2)
    End If
End If
End Sub

Private Sub Search_Button_Click()

' Search only in first worksheet "Student Checkout"
Worksheets(1).Activate

' Search query is the content of the WSU ID Box
Query = WSUID_TextBox.Text

' If Query isn't "ID" (Don't want to select first row)
If Query <> "ID" Then
    ' for Update button validation
    selected_ID = Query
    
    ' Allow form to update
    update = True
    
    ' Select the entire of Column A
    Columns("A:A").Select
    
    ' Look for all cells that match the query, starting from the first cell
    ' TODO Duplicates
    Set cell = Selection.Find(What:=Query, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
    ' If cell containing ID isn't found, display 0 for everything
    If cell Is Nothing Then
        FoodC_Label.Caption = "0"
        HygieneC_Label.Caption = "0"
        BabyC_Label.Caption = "0"
        OtherC_Label.Caption = "0"
        TotalC_Label.Caption = "0"
    ' Otherwise, retrieve data from cell
    Else
        Call Display_Current_Stats
    End If
    
    FoodU_Label.Caption = ""
    HygieneU_Label.Caption = ""
    BabyU_Label.Caption = ""
    OtherU_Label.Caption = ""
    Warning_Label.Visible = False
End If

End Sub

Private Sub Total_TextBox_Change()

End Sub

Private Sub Update_Button_Click()

If update Then
    Dim food As String
    Dim hygiene As String
    Dim baby As String
    Dim other As String
    
    Dim FoodU As Integer
    Dim HygieneU As Integer
    Dim BabyU As Integer
    Dim OtherU As Integer
    
    ' Get the amounts
    food = Food_TextBox.Text
    hygiene = Hygiene_TextBox.Text
    baby = Baby_TextBox.Text
    other = Other_TextBox.Text
    
    FoodU = IIf(IsNull(food), 0, Val(food))
    HygieneU = IIf(IsNull(hygiene), 0, Val(hygiene))
    BabyU = IIf(IsNull(baby), 0, Val(baby))
    OtherU = IIf(IsNull(other), 0, Val(other))
    
    If cell Is Nothing Then
        Dim empty_cell As Range
        
        ' Find first empty cell in row, then set the cell to that variable'
        ' Set empty_cell = Range("A1").End(xlDown).Offset(1, 0).Select
        ' next_empty = Range("A2:A" & Rows.Count).Cells.SpecialCells(xlCellTypeBlanks).Cells
        
        Set cell = Selection.Find(What:="", LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        cell.Value = WSUID_TextBox.Value
    End If
    
    ' Update cells if given a new value
    cell.Offset(0, 1).Value = IIf(FoodU, cell.Offset(0, 1).Value + FoodU, cell.Offset(0, 1).Value)
    cell.Offset(0, 2).Value = IIf(HygieneU, cell.Offset(0, 2).Value + HygieneU, cell.Offset(0, 2).Value)
    cell.Offset(0, 3).Value = IIf(BabyU, cell.Offset(0, 3).Value + BabyU, cell.Offset(0, 3).Value)
    cell.Offset(0, 4).Value = IIf(OtherU, cell.Offset(0, 4).Value + OtherU, cell.Offset(0, 4).Value)

    ' Reset everything here
    Food_TextBox.Value = ""
    Hygiene_TextBox.Value = ""
    Baby_TextBox.Value = ""
    Other_TextBox.Value = ""
    
    FoodU_Label.Caption = ""
    HygieneU_Label.Caption = ""
    BabyU_Label.Caption = ""
    OtherU_Label.Caption = ""
    
    Call Display_Current_Stats
End If

End Sub

Private Sub WSUID_TextBox_Change()

' Disallow updating if current content of TextBox isn't the same as previous content
' Updating is allowed after Search is clicked

update = Not (WSUID_TextBox.Text <> selected_ID)
' If Warning is active when ID Box changes,
' Warning Label is still displayed when user ID changes, and that's okay
    
End Sub

