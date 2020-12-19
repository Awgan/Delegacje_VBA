Attribute VB_Name = "Module1"
Sub nowa_pozycja()



Dim nr1, nr2
Dim rok As String

    If Arkusz1.Range("H4").Value <> "" Then     'sprawdzenie czy pole z nazwiskiem zosta³o wype³nione
        
        'Generowanie numeru delegacji
        
            rok = Arkusz2.Range("H2").Value
            
            
            nr1 = Arkusz3.Range("B6").Value
        
            Sheets("Arkusz1").Select
            Range("B4").Value = nr1 + 1
        
            nr2 = nr1 + 1
        
        
        
                 
        
        
        
        'dezaktywancja zabezpieczenia Arkusza 3
            
                Sheets("Arkusz3").Select
                ActiveSheet.Unprotect Password:="toropol12"
                    
        'tu wpisujemu kod grzebi¹cy po arkuszu, pobieramy b¹dŸ modyfukujemy
                    
        'Wstawianie nowego wiersza w Arkuszu 3
        
        Range("A6").Select
        Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
        
        
        'Wstawienie informacji do listy w ARKUSZU 3 '
        
        Arkusz3.Range("A6").Value = Arkusz1.Range("A4").Value
        Arkusz3.Range("B6").Value = Arkusz1.Range("B4").Value
        Arkusz3.Range("C6").Value = Arkusz1.Range("C4").Value
        Arkusz3.Range("D6").Value = Arkusz1.Range("D4").Value
        Arkusz3.Range("E6").Value = Arkusz1.Range("E4").Value
        Arkusz3.Range("F6").Value = Arkusz1.Range("F4").Value
        Arkusz3.Range("G6").Value = Arkusz1.Range("G4").Value
        Arkusz3.Range("H6").Value = Arkusz1.Range("H4").Value
        Arkusz3.Range("I6").Value = Arkusz1.Range("I4").Value
        
        Worksheets("Arkusz5").Range("I4") = Worksheets("Arkusz1").Range("B4") & " / " & rok & " / " & Year(Now)
        Worksheets("Arkusz5").Range("B8") = Worksheets("Arkusz1").Range("H4")
        Worksheets("Arkusz5").Range("B11") = Worksheets("Arkusz1").Range("E4") & ", " & Worksheets("Arkusz1").Range("F4")
        Worksheets("Arkusz5").Range("B13") = Worksheets("Arkusz1").Range("D4")
        
        Worksheets("Arkusz5").Range("B15") = Worksheets("Arkusz1").Range("G4")
        
        Worksheets("Arkusz5").Range("A29") = "Warszawa"
        Worksheets("Arkusz5").Range("E29") = Worksheets("Arkusz1").Range("E4")
        Worksheets("Arkusz5").Range("A30") = Worksheets("Arkusz1").Range("E4")
        Worksheets("Arkusz5").Range("E30") = "Warszawa"
        
        Worksheets("Arkusz5").Range("C29") = Worksheets("Arkusz1").Range("D4")
        Worksheets("Arkusz5").Range("C30") = Worksheets("Arkusz1").Range("D4")
        Worksheets("Arkusz5").Range("G29") = Worksheets("Arkusz1").Range("D4")
        Worksheets("Arkusz5").Range("G30") = Worksheets("Arkusz1").Range("D4")
        
        
        Range("B6:H6").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
        
            
        'wtawianie listy rozwijanej gwarancji'
        
        'Arkusz1.Range("I6").Select
            
            'With Selection.Validation
              '  .Delete
              '  .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
              '  xlBetween, Formula1:="=taknie"
               ' .IgnoreBlank = True
              '  .InCellDropdown = True
              '  .InputTitle = ""
              '  .ErrorTitle = ""
              '  .InputMessage = ""
              ' .ErrorMessage = ""
              '  .ShowInput = True
              '  .ShowError = True
           ' End With
        
        
        
        
        'Ustalenie szerokoœci kolumn'
        Arkusz3.Columns("B:D").EntireColumn.AutoFit
        
        Arkusz3.Columns("E:H").EntireColumn.ColumnWidth = 35
        
        Arkusz3.Columns("A").EntireColumn.ColumnWidth = 10
        
        
        
        
        'Okienko komunikatu
        
        nr2 = nr1 + 1
        
        
        Arkusz1.Select
        MsgBox "Twój numer delagacj to: " & " >> " & nr2 & " / " & rok & " / " & Year(Now) & " << " & Chr(13) & Chr(13) & "W DROGÊ !!!" & Chr(13) & Chr(13) & "Podró¿e kszta³c¹ wykszta³conych.", vbSystemModal, ">> TWÓJ NUMER <<"
        
        'Zerowanie komórek
        
        Range("D4:M4").Value = ""
        
        'zmiana numeru delegacji na kolejny
        
        Range("B4").Value = nr2 + 1
        
        
        
        
        
        
        'Centrowanie na komórce
        
        Sheets("Arkusz3").Select
        Range("B6").Select
        
        
        
        'Okienko
        
        'ochrony Arkusza 3
        Sheets("Arkusz3").Select
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="toropol12"
                
        
        
        
        
        'Formatowanie arkusza wydruku delegacji
        
        Worksheets("Arkusz5").Select
         
                    
                    With Worksheets("Arkusz5")
                    
                        .Range("C29:C34,G29:G34").Select
                        Selection.NumberFormat = "dd/mm/yyyy"
                         
                        .Range("D29:D34,H29:H34").Select
                        Selection.NumberFormat = "h:mm;@"
                    
                        .Range("I29:I34, F59, F64").Select
                        Selection.NumberFormat = "General"
                        
                        .Range("D29:D34,H29:H34,I29:I34").Value = " "
                        .Range("A18").Value = " "
                        .Range("F8").Value = " "
                        .Range("K29:K43, C41, B43") = " "
                        .Range("K29:K43").Font.Size = 10
                        
                        .Range("E45") = Now()
                        
                        .Range("A23,A51,A55,A61,A66").Value = " "
                        
                        .Range("A23") = Now()
                        
                        .Range("C59,C64").Value = " "
                        .Range("F59,F64").Value = " "
                        .Range("I29").Value = "=A18"
                        .Range("I30").Value = "=A18"
                        .Range("K41").Formula = " "
                        .Range("K43").Formula = " "
                        .Range("K41").Formula = "=sum(K29:K40)"
                        .Range("K43").Formula = " "
                        
                        .Range("C59").Value = "=K43"
                        .Range("F59").Value = "=C41"
                        
                        .Range("C64").Value = "=K43"
                        .Range("F64").Value = "=C41"
                        
                        
                    
                    End With
        
        
        'Zapisywanie Skoroszytu
        
        ActiveWorkbook.Save
        
        Sheets("Arkusz5").Range("F8").Select
        
    Else
    MsgBox "Formularz nie zosta³ wype³niony ca³kowicie !!!"
    
    End If

End Sub



