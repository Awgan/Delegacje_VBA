Attribute VB_Name = "Ten_skorosz"
Private Sub Workbook_Open()


Worksheets("Arkusz1").Select        'inicjalizacja dokumentu
Range("D4").Select



Application.DisplayFullScreen = False

Application.ScreenUpdating = False




If Worksheets("Arkusz2").Range("I2").Value <> Year(Now) Then        'je¿eli nast¹pi³a zmiana roku, to generuje siê nowa lista

        
    
        
        Worksheets("Arkusz3").Select

        
ActiveSheet.Unprotect Password:="toropol12"
    
            
            With Worksheets("Arkusz3")
                
                .Range("B6").EntireRow.Insert           '3 nowe wiersze na liœcie
                .Range("B6").EntireRow.Insert
                .Range("B6").EntireRow.Insert
                
                .Range("A7").Value = "? " & Year(Now) & " ** " & "? " & Year(Now) & " ** " & "? " & Year(Now) & " ** "
                .Range("A7").EntireRow.Interior.Color = 5296274
                                   
                
            End With



ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="toropol12"



        
        
        
        
Sheets("Arkusz2").Visible = True                'odkrycie ukrytego arkusza2, inaczej nie mo¿na wprowadziæ zmian do arkusza
Worksheets("Arkusz2").Select
            
            
            
            
            Dim literka As String
            
            literka = InputBox("Zapytaj prezesa jak¹ literke wymyœli³ na ten rok " & Year(Now), "Literka Na rok " & Year(Now))
                        
            
            
            'zmiana na nowy rok w Arkuszu2
            
            ActiveSheet.Unprotect Password:="toropol12" 'odhaœlenie
            
            
                    Range("I2").Value = Year(Now)       'zmiana roku na obecny
                    Range("H2").Value = literka         'zmiana literki roku
                    
            
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="toropol12"
        
        
        
       
Sheets("Arkusz2").Visible = False               'ukrycie arkusza2




End If







Dim nr1, nr2

Worksheets("Arkusz1").Select
Range("D4").Select


        'Data w okienku: data wpisu delegacji

ActiveSheet.Unprotect Password:="toropol12"

        
        
        'Generowanie numeru delegacji

        nr1 = Arkusz3.Range("B6").Value

        Sheets("Arkusz1").Select
        Range("B4").Value = nr1 + 1

        nr2 = nr1 + 1
    
        Range("C4").Value = Date






ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="toropol12"

Application.DisplayFullScreen = False

Application.ScreenUpdating = True



End Sub




