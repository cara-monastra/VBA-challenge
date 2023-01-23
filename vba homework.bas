Attribute VB_Name = "Module1"

Sub Ticker()
    
    Dim openprice As Double
    Dim closeprice As Double
    
    Dim Ticker As String
    
    Dim yearlychange As Double
    
    Dim summarytablerow As Integer
    summarytablerow = 2
    
    Dim totalvolume As Double
    
    Dim percentchange As Double

    
    totalvolume = 0
    
    yearlychange = 0
    percentchange = 0
    openprice = 0
    closeprice = 0
    totalvolume = 0
    
    For i = 2 To 22770
                     
     
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
         
            Ticker = Cells(i, 1).Value
            
            openprice = Cells(i - 250, 3).Value
    
            closeprice = Cells(i, 6).Value

            totalvolume = totalvolume + Cells(i, 7).Value
    
            yearlychange = closeprice - openprice

            percentchange = (yearlychange / openprice) * 100
            
            'print the ticker name in the summary table
            Range("I" & summarytablerow).Value = Ticker
     
            Range("J" & summarytablerow).Value = yearlychange
            
  
            Range("K" & summarytablerow).Value = percentchange
    
            Range("L" & summarytablerow).Value = totalvolume
    
            If yearlychange < 0 Then
            
                'if the yearly change is neg, cell is red
                Range("J" & summarytablerow).Interior.ColorIndex = 3
            Else
                
                'if the yearly change is pos, the cell is green
                Range("J" & summarytablerow).Interior.ColorIndex = 4
            End If
            
            
            'summary
            summarytablerow = summarytablerow + 1
            
            'reset
            totalvolume = 0
            openprice = 0
            closeprice = 0
            percentchange = 0
                
        
        Else
            
            totalvolume = totalvolume + Cells(i, 7).Value
            
    
        End If
        
    Next i
    
End Sub
