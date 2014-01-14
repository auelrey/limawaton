limawaton
=========

limawaton

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()

Dim startdate As Date
    Dim enddate As Date
    
    startdate = DTPicker1.Value
    enddate = DTPicker2.Value
     counter = 2
    counter1 = 2
    x = 0
    y = 0
    g = 0
    
    If ComboBox1.Value = "Verizon wireless" Then
        Sheets(1).Select
        Unload Me
        
        Do
            If Sheets("Verzon Data").Cells(counter, 1).Value = startdate Then
            End If
            counter = counter + 1
        Loop Until Sheets("Verzon Data").Cells(counter, 1).Value = startdate
        x = counter
        
        Do
            If Sheets("Verzon Data").Cells(counter1, 1).Value = enddate Then
            End If
            counter1 = counter1 + 1
        Loop Until Sheets("Verzon Data").Cells(counter1, 1).Value = enddate
        y = counter1
        
        g = x - y
        
        dataSelected = Sheets("Verzon Data").Range("A" & y & ":" & "G" & x).Value
        Sheets("Sheet1").Range("G24:M" & g + 24).Value = dataSelected
        
        dataSelected2 = Sheets("Verzon Data").Range("A1:G1").Value
        Sheets("Sheet1").Range("G23:M23").Value = dataSelected2
        
        With ActiveSheet
            If .ChartObjects.Count > 0 Then
            .ChartObjects.Select
            .ChartObjects(1).Delete
            End If
            .Shapes.AddChart.Select
        End With

        With ActiveChart
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete
            Loop
            With .Parent
                .Left = 300
                .Top = 0
                .Width = 500
                .Height = 300
            End With
            .ChartType = xlLine 'this is the type of chart we'll make. see if you can find other types
            .ChartArea.Select
            .HasTitle = True
            .ChartTitle.Text = "Price Graph" 'you can change the title here
        End With
        
       
        
        With ActiveChart
            .SeriesCollection.NewSeries
            .SeriesCollection(1).Name = Sheets("Verzon Data").Cells(1, 5).Value
            .SeriesCollection(1).Values = Sheets("Verzon Data").Range("E" & y & ":" & "E" & x)
            .SeriesCollection(1).XValues = Sheets("Verzon Data").Range("A" & y & ":" & "A" & x)
        End With
    
    ElseIf ComboBox1.Value = "AT&T" Then
            Sheets(1).Select
            Unload Me
        
        Do
            If Sheets("Verzon Data").Cells(counter, 1).Value = startdate Then
            End If
            counter = counter + 1
        Loop Until Sheets("Verzon Data").Cells(counter, 1).Value = startdate
        x = counter
        
        Do
            If Sheets("Verzon Data").Cells(counter1, 1).Value = enddate Then
            End If
            counter1 = counter1 + 1
        Loop Until Sheets("Verzon Data").Cells(counter1, 1).Value = enddate
        y = counter1
        
        g = x - y
        
        dataSelected = Sheets("Verzon Data").Range("I" & y & ":" & "O" & x).Value
        Sheets("Sheet1").Range("G24:M" & g + 24).Value = dataSelected
        
        dataSelected2 = Sheets("Verzon Data").Range("A1:G1").Value
        Sheets("Sheet1").Range("G23:M23").Value = dataSelected2
        
        With ActiveSheet
            If .ChartObjects.Count > 0 Then
            .ChartObjects.Select
            .ChartObjects(1).Delete
            End If
            .Shapes.AddChart.Select
        End With

        With ActiveChart
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete
            Loop
            With .Parent
                .Left = 300
                .Top = 0
                .Width = 500
                .Height = 300
            End With
            .ChartType = xlLine 'this is the type of chart we'll make. see if you can find other types
            .ChartArea.Select
            .HasTitle = True
            .ChartTitle.Text = "Price Graph" 'you can change the title here
        End With
        
       
        
         With ActiveChart
            .SeriesCollection.NewSeries
            .SeriesCollection(1).Name = Sheets("Verzon Data").Cells(1, 5).Value
            .SeriesCollection(1).Values = Sheets("Verzon Data").Range("M" & y & ":" & "M" & x)
            .SeriesCollection(1).XValues = Sheets("Verzon Data").Range("I" & y & ":" & "I" & x)
        End With
    
    
    
    
        
        
        
        
    ElseIf ComboBox1.Value = "PTR" Then 'collect data for PTR
        Sheets(1).Select
        Unload Me
        
        Do
            If Sheets("DataSheet").Cells(counter, 1).Value = startdate Then
            End If
            counter = counter + 1
        Loop Until Sheets("DataSheet").Cells(counter, 1).Value = startdate
        x = counter
        
        Do
            If Sheets("DataSheet").Cells(counter1, 1).Value = enddate Then
            End If
            counter1 = counter1 + 1
        Loop Until Sheets("DataSheet").Cells(counter1, 1).Value = enddate
        y = counter1
        
        g = x - y
        
        dataSelected = Sheets("DataSheet").Range("A" & y & ":" & "G" & x).Value
        Sheets(1).Range("I23:O" & g + 23).Value = dataSelected
        
        Range("I23:O" & g + 23).Cells.Font.Color = RGB(255, 0, 0)
        Range("I23:O" & g + 23).Interior.Color = RGB(162, 205, 90)
        
        dataSelected2 = Sheets("DataSheet").Range("A1:G1").Value
        Sheets(1).Range("I22:O22").Value = dataSelected2
        
        Range("I22:O22").Cells.Font.Color = RGB(255, 0, 0)
        Range("I22:O22").Interior.Color = RGB(202, 255, 112)
        
        With ActiveSheet
            If .ChartObjects.Count > 0 Then
            .ChartObjects.Select
            .ChartObjects(1).Delete
            End If
            .Shapes.AddChart.Select
        End With

        With ActiveChart
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete
            Loop
            With .Parent
                .Left = 300
                .Top = 0
                .Width = 500
                .Height = 300
            End With
            .ChartType = xlLine
            .ChartArea.Select
            .HasTitle = True
            .ChartTitle.Text = "Closing Price" 'you can change the title here
        End With
    
        With ActiveChart
            .SeriesCollection.NewSeries
            .SeriesCollection(1).Name = Sheets("DataSheet").Cells(1, 5).Value
            .SeriesCollection(1).Values = Sheets("DataSheet").Range("E" & y & ":" & "E" & x)
            .SeriesCollection(1).XValues = Sheets("DataSheet").Range("A" & y & ":" & "A" & x)
        End With
        
   
    
    ElseIf ComboBox1.Value = "SHI" Then 'collect data for SHI
        Sheets(1).Select
        Unload Me
        
        Do
            If Sheets("DataSheet").Cells(counter, 9).Value = startdate Then
            End If
            counter = counter + 1
        Loop Until Sheets("DataSheet").Cells(counter, 9).Value = startdate
        x = counter
        
        Do
            If Sheets("DataSheet").Cells(counter1, 9).Value = enddate Then
            End If
            counter1 = counter1 + 1
        Loop Until Sheets("DataSheet").Cells(counter1, 9).Value = enddate
        y = counter1
        
        g = x - y
        
        dataSelected = Sheets("DataSheet").Range("I" & y & ":" & "O" & x).Value
        Sheets(1).Range("I23:O" & g + 23).Value = dataSelected
        
        Range("I23:O" & g + 23).Cells.Font.Color = RGB(255, 0, 0)
        Range("I23:O" & g + 23).Interior.Color = RGB(162, 205, 90)
        
        dataSelected2 = Sheets("DataSheet").Range("A1:G1").Value
        Sheets(1).Range("I22:O22").Value = dataSelected2
        
        Range("I22:O22").Cells.Font.Color = RGB(255, 0, 0)
        Range("I22:O22").Interior.Color = RGB(202, 255, 112)
        
        With ActiveSheet
            If .ChartObjects.Count > 0 Then
            .ChartObjects.Select
            .ChartObjects(1).Delete
            End If
            .Shapes.AddChart.Select
        End With

        With ActiveChart
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete
            Loop
            With .Parent
                .Left = 300
                .Top = 0
                .Width = 500
                .Height = 300
            End With
            .ChartType = xlLine
            .ChartArea.Select
            .HasTitle = True
            .ChartTitle.Text = "Closing Price" 'you can change the title here
        End With
    
        With ActiveChart
            .SeriesCollection.NewSeries
            .SeriesCollection(1).Name = Sheets("DataSheet").Cells(1, 5).Value
            .SeriesCollection(1).Values = Sheets("DataSheet").Range("M" & y & ":" & "M" & x)
            .SeriesCollection(1).XValues = Sheets("DataSheet").Range("I" & y & ":" & "I" & x)
        End With
        
     End If
     
     
        
        
      
        
    
    
    

  
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub




Private Sub UserForm_Initialize()
ComboBox1.List = Sheets("sheet2").Range("B1:B4").Value



End Sub
