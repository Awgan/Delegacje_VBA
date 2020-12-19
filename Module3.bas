Attribute VB_Name = "Module3"
Sub wykres()

'#                              #
'#  usuwanie starych wykresów   #
'#                              #

Do Until Charts.Count = 0
    
        Application.DisplayAlerts = False
        Charts(1).Delete
        Application.DisplayAlerts = True
Loop


'#                              #
'#  wstawianie nowego wykresu   #
'#  i usuwanie wszystkich serii #

Dim wykres As Chart
Set wykres = Charts.Add

    wykres.Move after:=Sheets(Sheets.Count)
    wykres.Name = "Grand_Prix"
    wykres.ChartType = xlPie
    wykres.HasTitle = True
    wykres.ChartTitle.Text = "Grand Prix"

Do Until wykres.SeriesCollection.Count = 0
                    
    wykres.SeriesCollection(1).Delete
Loop


'#                                  #
'#  wstawianie nowej serii danych   #
'#                                  #

Dim s_last As String

Dim wsh_GP As Worksheet
Set wsh_GP = Worksheets("Grand_Prix_temp")
Dim r_temp As Range
Set r_temp = wsh_GP.Range("B1")

Do Until r_temp = ""
    Set r_temp = r_temp.Offset(1, 0)
Loop
Let s_last = r_temp.Row - 1

wykres.SeriesCollection.Add _
        Source:=wsh_GP.Range("b1:" & "b" & s_last)

wykres.SeriesCollection(1).XValues = wsh_GP.Range("a1:" & "a" & s_last)
wykres.SeriesCollection(1).ApplyDataLabels
wykres.SeriesCollection(1).DataLabels.ShowCategoryName = True

End Sub

Sub aktualizuj_listy_osob()

Dim rng1, rng2 As Range

Set rng1 = ActiveWorkbook.Worksheets("Arkusz2").Range("B2")
Set rng2 = ActiveWorkbook.Worksheets("Grand_Prix_temp").Range("A1")


Dim rng_x As Range
Set rng_x = rng2

'#                          #
'#  usuwanie starej listy   #
'#                          #
Do Until rng_x = ""
    
    rng_x.Clear
    Set rng_x = rng_x.Offset(1, 0)
Loop


'#                              #
'#  dodawanie aktualnej listy   #
'#                              #
Do Until rng1 = ""
    
    rng2.Value = rng1.Value
    
    Set rng2 = rng2.Offset(1, 0)
    Set rng1 = rng1.Offset(1, 0)
Loop

'#                              #
'#  obliczanie iloœci wyjazdów  #
'#                              #

Set rng1 = ActiveWorkbook.Worksheets("Grand_Prix_temp").Range("A1")
Set rng2 = ActiveWorkbook.Worksheets("Grand_Prix_temp").Range("B1")

Dim i As Integer
Let i = 1
Do Until rng1 = ""
    
    rng2.FormulaLocal = "=LICZ.JE¯ELI(Arkusz3!$H$6:$H$119;A" & i & ")"
    Set rng2 = rng2.Offset(1, 0)
    Set rng1 = rng1.Offset(1, 0)
    Let i = i + 1
    
Loop


'#                              #
'#  usuwanie pozycji z zerow¹   #
'#  iloœci¹ wyjazdów            #
'#                              #
Set rng2 = ActiveWorkbook.Worksheets("Grand_Prix_temp").Range("B1")
Dim temp_adr As String

Do Until rng2 = ""

    If rng2.Value = 0 Then
        Let temp_adr = rng2.Address
            rng2.EntireRow.Delete
        Set rng2 = ActiveWorkbook.Worksheets("Grand_Prix_temp").Range(temp_adr)
    Else

        Set rng2 = rng2.Offset(1, 0)

    End If
Loop
    
    
End Sub
