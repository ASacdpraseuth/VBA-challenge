Attribute VB_Name = "Test"
Sub Runthisone()

    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    
    For Each ws In Worksheets
    
        ws.Select
        Call Calculator
        
    Next
    
    Application.ScreenUpdating = True
    
    Application.ScreenUpdating = False
    
    For Each ws In Worksheets
    
        ws.Select
        Call Color
        
    Next
    
    Application.ScreenUpdating = True

End Sub
Sub Calculator()


        'Establish lastrow counts
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
    
        'Create headers
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
    
        'Establish variables
        Dim percent As Double
     
        'Get ticker list
        j = 2
        For i = 2 To lastrow
    
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
                Cells(j, 9).Value = Cells(i, 1).Value
                j = j + 1
            
            End If
                
        Next i
    
        'Get stock volume
        For j = 2 To lastrow2
        
            'Establishing null value for total stock volume
            Cells(j, 12).Value = 0
        
        Next j
        
        'Reseting i and j
        j = 2
        i = 2
    
        'Establish total stock volume
        For i = 2 To lastrow
    
                If Cells(i, 1).Value = Cells(j, 9) Then
        
                    Cells(j, 12).Value = Cells(j, 12).Value + Cells(i, 7).Value
                
                Else
            
                    j = j + 1
            
                End If
    
        Next i
    
        'Reset column count
        i = 2
        j = 2
    
        'Establish yearly change
        For i = 2 To lastrow
        
            If Cells(i, 1) <> Cells(i - 1, 1) Then
        
                If Cells(i - 1, 1) = Range("A1") Then
            
                    Cells(j, 10).Value = Cells(i, 3).Value
                    
                ElseIf Cells(i, 3) = 0 Then
                
                    i = i
                    j = j + 1
                
                ElseIf Cells(i - 1, 6) = 0 Then
                
                    i = i
                    j = j + 1
                
                Else
                
                    Cells(j, 10) = Cells(i - 1, 6).Value - Cells(j, 10).Value
                    percent = (Cells(j, 10).Value / Cells(i - 1, 6).Value)
                    Cells(j, 11).NumberFormat = "0.00%"
                    Cells(j, 11) = percent
                    i = i + 1
                    j = j + 1
            
                End If
                
            ElseIf Cells(j, 10) <> 0 Then
        
                i = i
            
            Else
        
                Cells(j, 10).Value = Cells(i, 3).Value
  
            End If
    
        Next i
    
        'Reset columns
        i = 2
        j = 2
End Sub

Sub Color()
    
    lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color yearly change
    For j = 2 To lastrow2
    
        If Cells(j, 10) > 0 Then
                
            Cells(j, 10).Interior.ColorIndex = 4
                    
        ElseIf Cells(j, 10) < 0 Then
                
            Cells(j, 10).Interior.ColorIndex = 3
                    
        End If
        
    Next j

    
End Sub
