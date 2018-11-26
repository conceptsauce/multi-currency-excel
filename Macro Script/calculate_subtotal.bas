Attribute VB_Name = "calculate_subtotal"
Option Explicit
Option Base 1

Sub CalculateSubtotal()

    Dim lastRow As Long, lastColumn As Long, rowCnt As Long, colCnt As Long, exCnt As Long
    Dim vCosts As Variant, vLabels As Variant, vExchange As Variant
    Dim convRate As Double, convVal As Double
    Dim mSubtotal As Double, qSubtotal As Double, ySubtotal As Double, y5Total As Double
    Dim errCell As String
    Dim matchFound As Boolean

    With ActiveSheet
    
        lastRow = .Cells.Find(what:="*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        If lastRow <= FirstCostRow Then Exit Sub 'no cost data
        lastColumn = .Cells.Find(what:="*", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
        vCosts = .Range(.Cells(FirstCostRow, CurrencyCol), .Cells(lastRow, lastColumn))
        vLabels = .Range(.Cells(FirstCostRow, 1), .Cells(lastRow, 1))
        
        With Sheet2 'XE Rates
            lastRow = .Cells.Find(what:="*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
            vExchange = .Range(.Cells(FirstExchangeRow, CurrencyCol), .Cells(lastRow, lastColumn))
        End With
    
        Application.ScreenUpdating = False
        
        For colCnt = LBound(vCosts, 2) + 1 To UBound(vCosts, 2) 'go through all months, skip currency column
        
            mSubtotal = 0
            If colCnt Mod 3 = 2 Then qSubtotal = 0
            If colCnt Mod 12 = 2 Then ySubtotal = 0
            
            For rowCnt = LBound(vCosts) To UBound(vCosts)
            
                If vLabels(rowCnt, 1) = SubTotalName Then 'subtotal row
                    .Cells(FirstCostRow + rowCnt - 1, colCnt + 1) = mSubtotal: mSubtotal = 0
                ElseIf vLabels(rowCnt, 1) = TotalName5y Then '5 year sum
                    If colCnt = UBound(vCosts, 2) Then .Cells(FirstCostRow + rowCnt - 1, 2) = y5Total
                ElseIf vLabels(rowCnt, 1) = TotalName1y Then '1st year sum
                    If colCnt = 13 Then .Cells(FirstCostRow + rowCnt - 1, 2) = y5Total
                Else
            
                    If Not IsEmpty(vCosts(rowCnt, colCnt)) Then 'cost value entered
                        If IsNumber(CStr(vCosts(rowCnt, colCnt))) Then 'cost value is really a numeric value representing price, not some text comment
                            If Not IsEmpty(vCosts(rowCnt, 1)) Then 'currency entered
                                
                                If vCosts(rowCnt, 1) = DefaultCurrency Then
                                    convRate = 1
                                Else
                                
                                    matchFound = False
                                
                                    For exCnt = LBound(vExchange) To UBound(vExchange)
                                        If vCosts(rowCnt, 1) = vExchange(exCnt, 1) Then
                                            convRate = vExchange(exCnt, colCnt)
                                            matchFound = True
                                            Exit For
                                        End If
                                    Next exCnt
                                    
                                    If matchFound = False Then 'currency not found
                                        errCell = Replace(.Cells(FirstCostRow + rowCnt - 1, CurrencyCol).Address, "$", vbNullString)
                                        MsgBox ("Currency name " & vCosts(rowCnt, 1) & " entered in cell " & errCell & " of Hosting and Services Costs worksheet not found." & String(2, vbLf) & "Operation aborted.")
                                        Application.ScreenUpdating = True
                                        Exit Sub
                                    End If
                                    
                                End If
                                
                                convVal = Round(vCosts(rowCnt, colCnt) * convRate, 2)
                                
                                mSubtotal = mSubtotal + convVal
                                'qSubtotal = qSubtotal + convVal
                                'ySubtotal = ySubtotal + convVal
                                y5Total = y5Total + convVal
                                
                            Else 'currency not entered
                                errCell = Replace(.Cells(FirstCostRow + rowCnt - 1, CurrencyCol).Address, "$", vbNullString)
                                MsgBox ("Currency not specified in cell " & errCell & " of Hosting and Services Costs worksheet." & String(2, vbLf) & "Operation aborted.")
                                Application.ScreenUpdating = True
                                Exit Sub
                            End If
                            
                        End If
                    End If
                    
                End If
                
            Next rowCnt
            
        Next colCnt
        
        Application.ScreenUpdating = True
        
    End With

End Sub

Function IsNumber(ByVal myVal As String) As Boolean

    Dim nCounter As Long
    
    myVal = Replace(myVal, ".", vbNullString)
    myVal = Replace(myVal, ",", vbNullString)
    
    For nCounter = 1 To Len(myVal)
        Select Case Asc(Mid(myVal, nCounter, 1))
        Case 48 To 57
        Case Else
            Exit Function
        End Select
    Next nCounter
    
    IsNumber = True
    
End Function
