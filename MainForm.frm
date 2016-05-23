VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Game of Life Controls"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'
' Jon Conway's Game of Life
'
' Coded by Richard Kelley
'
'   May 2012
'


Dim Gen_Count As Integer
Dim Grid As Variant
Dim Grid2 As Variant
Dim Speed As Integer
Dim StopNow As Boolean
Dim Model As Integer
Dim Cont As Boolean




Private Sub ContinueBox_Click()

    If ContinueBox.Value = True Then
        
        Cont = True
    
    ElseIf ContinueBox.Value = False Then
    
        Cont = False
    
    End If
    
End Sub

Private Sub Exit_Button_Click()

    Me.Hide
    
End Sub

Private Sub Models_Dropdown_Change()

    Sheet1.Cells.ClearContents
    Gen_Count = 0
    Me.Display_Gen.Caption = Gen_Count
    Me.Repaint

    If Models_Dropdown.Value = "Glider" Then
        
        Model = 1
    
    ElseIf Models_Dropdown.Value = "Tumbler" Then
    
        Model = 2
    
    ElseIf Models_Dropdown.Value = "Shooter" Then
    
        Model = 3
    
    Else
    
        Model = 0
        
    End If
    
End Sub

Private Sub Next_Button_Click()

        GameOfLife
        Gen_Count = Gen_Count + 1
        Me.Display_Gen.Caption = Gen_Count
        Me.Repaint
    
End Sub


Private Sub Reset_Button_Click()

    Sheet1.Cells.ClearContents
    Gen_Count = 0
    Me.Display_Gen.Caption = Gen_Count
    Me.Repaint
    
End Sub




Private Sub Speed_DropDown_Change()

        If Speed_DropDown.Value = "Slow" Then
        
            Speed = 0
        
        ElseIf Speed_DropDown.Value = "Medium" Then
        
            Speed = 1
            
        Else
            Speed = 2
            
        End If

End Sub

Private Sub Start_Button_Click()

If Cont = False Then
    Gen = 100
Else
    Gen = 1000000000

End If

If StopNow = True Then
    Start_Button.Caption = "Stop"
    StopNow = False
ElseIf StopNow = False Then
    Start_Button.Caption = "Start"
    StopNow = True
End If
    
    For i = 1 To Gen
    DoEvents
    
        If StopNow Then Exit For
                       
        If Speed = 0 Then
        
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 2
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
        
        ElseIf Speed = 1 Then
        
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 1
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
            
        Else
            'Fast is regular speed
        End If
        
        GameOfLife
        Gen_Count = Gen_Count + 1
        Me.Display_Gen.Caption = Gen_Count
        Me.Repaint
        
    Next i
        
    Start_Button.Caption = "Start"
    StopNow = True
End Sub

Private Sub UserForm_Initialize()

    Sheet1.Cells.ClearContents
    GridSize = 50
    Gen_Count = 0
    StopNow = True
    Cont = False
    
    ActiveSheet.Cells(1, 1).Select
    Grid_Right_Corner = ActiveCell.Offset(GridSize - 1, GridSize - 1).Address
    ActiveSheet.Range("A1:" & Grid_Right_Corner).Value = 0
    ActiveSheet.Range("A1:" & Grid_Right_Corner).ColumnWidth = 2
    ActiveSheet.Range("A1:" & Grid_Right_Corner).RowHeight = 12
    ActiveCell.Range("A1:" & Grid_Right_Corner).BorderAround ColorIndex:=3, Weight:=xlThick
    
    Speed_DropDown.AddItem "Slow"
    Speed_DropDown.AddItem "Medium"
    Speed_DropDown.AddItem "Fast"
    Speed_DropDown.Value = "Fast"
    
    Models_Dropdown.AddItem "None"
    Models_Dropdown.AddItem "Glider"
    Models_Dropdown.AddItem "Tumbler"
    Models_Dropdown.AddItem "Shooter"
    Models_Dropdown.Value = "None"
    
End Sub


Private Sub GameOfLife()

    Dim Neighbors As Integer
    GridSize = 50
    
    ActiveSheet.Cells(1, 1).Select
    Grid_Right_Corner = ActiveCell.Offset(GridSize - 1, GridSize - 1).Address
    
    ' Load Current display
    Grid = Range("A1:" & Grid_Right_Corner)
    
    ' Load Models
    If Gen_Count = 0 Then
    
        Load_Model (Model)
        
    End If
            
    ' Load up working grid
    Grid2 = Grid
        
    Range("A1:" & Grid_Right_Corner) = Grid2
    
    For i = 1 To GridSize
    
        For j = 1 To GridSize
        
            iminus1 = i - 1
            iplus1 = i + 1
            jminus1 = j - 1
            jplus1 = j + 1

            If i - 1 <= 0 Then
                iminus1 = GridSize
            End If
    
            If i + 1 >= GridSize Then
                 iplus1 = 1
            End If
  
            If j - 1 <= 0 Then
                jminus1 = GridSize
            End If
    
            If j + 1 >= GridSize Then
                jplus1 = 1
            End If
            
            Neighbors = Grid(iminus1, jminus1) + Grid(i, jminus1) + Grid(iplus1, jminus1) + Grid(iminus1, j) + Grid(iplus1, j) + Grid(iminus1, jplus1) + Grid(i, jplus1) + Grid(iplus1, jplus1)
        
            If Grid(i, j) = 1 Then
            
                If Neighbors <= 1 Or Neighbors >= 4 Then
                
                    Grid2(i, j) = 0
                    'MsgBox ("Grid" & i & " " & j & " N = " & Neighbors)
                Else
                
                    Grid2(i, j) = 1
                    'MsgBox ("Grid" & i & " " & j & " N = " & Neighbors)
                    
                End If
                    
            Else
            
                If Neighbors = 3 Then
                
                    Grid2(i, j) = 1
                    
                End If
            
            End If
                       
        Next j
    
    Next i
    
    Range("A1:" & Grid_Right_Corner) = Grid2
    Grid = Null
    Grid2 = Null
    
End Sub

Private Sub Load_Model(Model_Choice As Integer)

    If Model_Choice = 1 Then
        
        ' Small Glider
        Grid(11, 10) = 1
        Grid(12, 11) = 1
        Grid(12, 12) = 1
        Grid(11, 12) = 1
        Grid(10, 12) = 1
        
    ElseIf Model_Choice = 2 Then
    
        'Tumbler
        Grid(10, 13) = 1
        Grid(10, 14) = 1
        Grid(10, 15) = 1
        Grid(11, 10) = 1
        Grid(11, 11) = 1
        Grid(11, 15) = 1
        Grid(12, 10) = 1
        Grid(12, 11) = 1
        Grid(12, 12) = 1
        Grid(12, 13) = 1
        Grid(12, 14) = 1
        Grid(14, 10) = 1
        Grid(14, 11) = 1
        Grid(14, 12) = 1
        Grid(14, 13) = 1
        Grid(14, 14) = 1
        Grid(15, 10) = 1
        Grid(15, 11) = 1
        Grid(15, 15) = 1
        Grid(16, 13) = 1
        Grid(16, 14) = 1
        Grid(16, 15) = 1
        
        
    ElseIf Model_Choice = 3 Then
    
        'Shooter
        Grid(10, 12) = 1
        Grid(10, 13) = 1
        Grid(11, 12) = 1
        Grid(11, 13) = 1
        Grid(18, 13) = 1
        Grid(18, 14) = 1
        Grid(19, 12) = 1
        Grid(19, 14) = 1
        Grid(20, 12) = 1
        Grid(20, 13) = 1
        Grid(26, 14) = 1
        Grid(26, 15) = 1
        Grid(26, 16) = 1
        Grid(27, 14) = 1
        Grid(28, 15) = 1
        Grid(32, 11) = 1
        Grid(32, 12) = 1
        Grid(33, 10) = 1
        Grid(33, 12) = 1
        Grid(34, 10) = 1
        Grid(34, 11) = 1
        Grid(34, 22) = 1
        Grid(34, 23) = 1
        Grid(35, 22) = 1
        Grid(35, 24) = 1
        Grid(36, 22) = 1
        Grid(44, 10) = 1
        Grid(44, 11) = 1
        Grid(45, 10) = 1
        Grid(45, 11) = 1
        Grid(45, 17) = 1
        Grid(45, 18) = 1
        Grid(45, 19) = 1
        Grid(46, 17) = 1
        Grid(47, 18) = 1
    
    Else
        ' Do nothing
    End If
       

End Sub



