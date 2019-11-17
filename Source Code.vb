'Created By Kodbear
'Declaration of Global Varibles
Public affLesCount As Integer 'Total Count of needed Lesser Potions needed
Public affGreatCount As Integer 'Total Count of needed Greater Potions needed
Public affSupCount As Integer 'Total Count of needed Superior Potions needed
Public arcLesCount As Integer 'Total Count of needed Lesser Potions needed
Public arcGreatCount As Integer 'Total Count of needed Greater Potions needed
Public arcSupCount As Integer 'Total Count of needed Superior Potions needed
Public affLes As Integer 'Current amount of Affinity Lesser Potions
Public affGreat As Integer 'Current amount of Affinity Greater Potions
Public affSup As Integer 'Current amount of Affinity Superior Potions
Public arcLes As Integer 'Current amount of Arcane Lesser Potions
Public arcGreat As Integer 'Current amount of Arcane Greater Potions
Public arcSup As Integer 'Current amount of Arcane Superior Potions



Sub assignVals()
Range("I25").Value = affLesCount
Range("J25").Value = affGreatCount
Range("K25").Value = affSupCount
Range("I26").Value = arcLesCount
Range("J26").Value = arcGreatCount
Range("K26").Value = arcSupCount

End Sub

Sub affLesCalc(champRarity As String, curAsc As Integer, wantedAsc As Integer)

Dim count As Integer
count = curAsc

If champRarity = "Uncommon" Then
     If wantedAsc >= 5 Then
        MsgBox "Sorry I don't have that information yet!"
     Else
        Do While count < wantedAsc
            Select Case count
                Case "0"
                    affLesCount = affLesCount + 2
                Case "1"
                    affLesCount = affLesCount + 2
                Case "2"
                    affLesCount = affLesCount + 3
                Case "3"
                    affLesCount = affLesCount + 3
            End Select
            count = count + 1
        Loop
     End If
     
ElseIf champRarity = "Rare" Then
    If curAsc > 2 Then
        affLesCount = 0
    Else
        Do While count < wantedAsc
            Select Case count
                Case "0"
                    affLesCount = affLesCount + 4
                Case "1"
                    affLesCount = affLesCount + 6
            End Select
            count = count + 1
        Loop
    End If
         
ElseIf champRarity = "Epic" Then
    affLesCount = 0
ElseIf champRarity = "Legendary" Then
    affLesCount = 0
Else
    
End If

End Sub
Sub affGreatCalc(champRarity As String, curAsc As Integer, wantedAsc As Integer)

Dim count As Integer
count = curAsc

If champRarity = "Uncommon" Then
     affGreatCount = 0
ElseIf champRarity = "Rare" Then
    Do While count < wantedAsc
            Select Case count
                Case "0"
                    
                Case "1"
                    
                Case "2"
                    affGreatCount = affGreatCount + 2
                Case "3"
                    affGreatCount = affGreatCount + 2
                Case "4"
                    affGreatCount = affGreatCount + 5
                Case "5"
                    affGreatCount = affGreatCount + 6
            End Select
            count = count + 1
        Loop
     
ElseIf champRarity = "Epic" Then
    If curAsc >= 4 Then
        affGreatCount = 0
    Else
        Do While count < wantedAsc
            Select Case count
                Case "0"
                    affGreatCount = affGreatCount + 4
                Case "1"
                    affGreatCount = affGreatCount + 7
                Case "2"
                    affGreatCount = affGreatCount + 9
                Case "3"
                    
                Case "4"
                    
                Case "5"
                    
            End Select
            count = count + 1
        Loop
    End If
    
ElseIf champRarity = "Legendary" Then
    affGreatCount = 0
Else

End If

End Sub
Sub affSupCalc(champRarity As String, curAsc As Integer, wantedAsc As Integer)

Dim count As Integer
count = curAsc

If champRarity = "Uncommon" Then
     If curAsc >= 5 Then
     MsgBox "Sorry I don't have that information yet!"
     
     Else
        affSupCount = 0
     End If
     
ElseIf champRarity = "Rare" Then
    affSupCount = 0
    
ElseIf champRarity = "Epic" Then
    Do While count < wantedAsc
            Select Case count
                Case "0"
                    
                Case "1"
                    
                Case "2"
                    
                Case "3"
                    affSupCount = affSupCount + 3
                Case "4"
                    affSupCount = affSupCount + 3
                Case "5"
                    affSupCount = affSupCount + 4
            End Select
            count = count + 1
    Loop
ElseIf champRarity = "Legendary" Then
    Do While count < wantedAsc
            Select Case count
                Case "0"
                    affSupCount = affSupCount + 1
                Case "1"
                    affSupCount = affSupCount + 2
                Case "2"
                    affSupCount = affSupCount + 3
                Case "3"
                    affSupCount = affSupCount + 4
                Case "4"
                    affSupCount = affSupCount + 5
                Case "5"
                    affSupCount = affSupCount + 6
            End Select
            count = count + 1
    Loop
Else
    
End If
End Sub
Sub arcLesCalc(champRarity As String, curAsc As Integer, wantedAsc As Integer)

Dim count As Integer
count = curAsc

If champRarity = "Uncommon" Then
     If curAsc >= 5 Then
     MsgBox "Sorry I don't have that information yet!"
     
     Else
        Do While count < wantedAsc
            Select Case count
                Case "0"
                    arcLesCount = arcLesCount + 2
                Case "1"
                    arcLesCount = arcLesCount + 2
                Case "2"
                    arcLesCount = arcLesCount + 2
                Case "3"
                    arcLesCount = arcLesCount + 3
            End Select
            count = count + 1
        Loop
     End If
     
ElseIf champRarity = "Rare" Then
    If curAsc > 2 Then
        arcLesCount = 0
    Else
        Do While count < wantedAsc
            Select Case count
                Case "0"
                    arcLesCount = arcLesCount + 2
                Case "1"
                    arcLesCount = arcLesCount + 3
            End Select
            count = count + 1
        Loop
    End If
    
     
ElseIf champRarity = "Epic" Then
    affLesCount = 0
ElseIf champRarity = "Legendary" Then
    affLesCount = 0
Else
    
End If

End Sub
Sub arcGreatCalc(champRarity As String, curAsc As Integer, wantedAsc As Integer)

Dim count As Integer
count = curAsc

If champRarity = "Uncommon" Then
     If curAsc >= 5 Then
     MsgBox "Sorry I don't have that information yet!"
     
     Else
     
     End If
     
ElseIf champRarity = "Rare" Then
    Do While count < wantedAsc
            Select Case count
                Case "0"
                    
                Case "1"
                    
                Case "2"
                    arcGreatCount = arcGreatCount + 1
                Case "3"
                    arcGreatCount = arcGreatCount + 2
                Case "4"
                    arcGreatCount = arcGreatCount + 3
                Case "5"
                    arcGreatCount = arcGreatCount + 4
            End Select
            count = count + 1
    Loop
     
ElseIf champRarity = "Epic" Then
    If curAsc > 3 Then
       arcGreatCount = 0
    Else
        Do While count < wantedAsc
            Select Case count
                Case "0"
                    arcGreatCount = arcGreatCount + 3
                Case "1"
                    arcGreatCount = arcGreatCount + 5
                Case "2"
                    arcGreatCount = arcGreatCount + 7
                Case "3"
                    
                Case "4"
                    
                Case "5"
                    
            End Select
            count = count + 1
        Loop
    End If
    
ElseIf champRarity = "Legendary" Then
    Do While count < wantedAsc
            Select Case count
                Case "0"
                    arcGreatCount = arcGreatCount + 5
                Case "1"
                    
                Case "2"
                    
                Case "3"
                    
                Case "4"
                    
                Case "5"
                    
            End Select
            count = count + 1
    Loop
End If

End Sub
Sub arcSupCalc(champRarity As String, curAsc As Integer, wantedAsc As Integer)

Dim count As Integer
count = curAsc

If champRarity = "Uncommon" Then
    arcSupCount = 0
ElseIf champRarity = "Rare" Then
    arcSupCount = 0
ElseIf champRarity = "Epic" Then
    Do While count < wantedAsc
            Select Case count
                Case "0"
                    
                Case "1"
                    
                Case "2"
                    
                Case "3"
                    arcSupCount = arcSupCount + 1
                Case "4"
                    arcSupCount = arcSupCount + 2
                Case "5"
                    arcSupCount = arcSupCount + 2
            End Select
            count = count + 1
    Loop
    
ElseIf champRarity = "Legendary" Then
    Do While count < wantedAsc
            Select Case count
                Case "0"
                    
                Case "1"
                    arcSupCount = arcSupCount + 2
                Case "2"
                    arcSupCount = arcSupCount + 2
                Case "3"
                    arcSupCount = arcSupCount + 3
                Case "4"
                    arcSupCount = arcSupCount + 4
                Case "5"
                    arcSupCount = arcSupCount + 4
            End Select
            count = count + 1
    Loop
Else
    
End If

End Sub
'Function that calls all calculate functions
Sub Calc(champRarity As String, curAsc As Integer, wantedAsc As Integer)
affLesCalc champRarity, curAsc, wantedAsc
affGreatCalc champRarity, curAsc, wantedAsc
affSupCalc champRarity, curAsc, wantedAsc
arcLesCalc champRarity, curAsc, wantedAsc
arcGreatCalc champRarity, curAsc, wantedAsc
arcSupCalc champRarity, curAsc, wantedAsc
subtractCurrent 'Subtract current potions from Required potions
End Sub
'Function to subtract current potions from required
Sub subtractCurrent()

If affLesCount > affLes Then
    affLesCount = affLesCount - affLes
Else
    affLesCount = 0
End If

If affGreatCount > affGreat Then
    affGreatCount = affGreatCount - affGreat
Else
    affGreatCount = 0
End If

If affSupCount > affSup Then
    affSupCount = affSupCount - affSup
Else
    affSupCount = 0
End If

If arcLesCount > arcLes Then
    arcLesCount = arcLesCount - arcLes
Else
    arcLesCount = 0
End If

If arcGreatCount > arcGreat Then
    arcGreatCount = arcGreatCount - arcGreat
Else
    arcGreatCount = 0
End If

If arcSupCount > arcSup Then
    arcSupCount = arcSupCount - arcSup
Else
    arcSupCount = 0
End If

End Sub
Sub calc_Btn()

'Variable Delcarations
Dim wantedAsc As Integer
Dim champRarity As String
Dim curAsc As Integer

Dim test As Integer

'Assign Values to Variables
wantedAsc = Range("I21").Value
champRarity = Range("I22").Value
curAsc = Range("I20").Value
affLesCount = 0
affGreatCount = 0
affSupCount = 0
arcLesCount = 0
arcGreatCount = 0
arcSupCount = 0
affLes = Range("I18").Value 'Set current amount of Affinity Lesser Potions
affGreat = Range("J18").Value 'Set current amount of Affinity Greater Potions
affSup = Range("K18").Value 'Set current amount of Affinity Superior Potions
arcLes = Range("I19").Value 'Set current amount of Arcane Lesser Potions
arcGreat = Range("J19").Value 'Set current amount of Arcane Greater Potions
arcSup = Range("K19").Value 'Set current amount of Arcane Superior Potions

'test = greatCalc(champRarity, curAsc, wantedAsc)
If curAsc >= wantedAsc Then
    MsgBox "Ya done messed up A-a-ron!"
Else
Calc champRarity, curAsc, wantedAsc
End If

assignVals

End Sub
