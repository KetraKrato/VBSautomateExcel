Sub automation()

     Dim obj As New WebDriver
     Dim sheetName As Variant
     sheetName = Array("Sheet1", "Sheet2")
     obj.Start "edge", "" '--For chrome browser replace edge with chrome
     obj.Get "https://www.automationandagile.com/p/sample-form.html"
     
     
        For Each Name In sheetName
        i = 2
        Do Until i = 20
        If IsEmpty(ThisWorkbook.Sheets(Name).Range("A" & i).Value) = False Then
     obj.FindElementByName("fname").SendKeys (ThisWorkbook.Sheets(Name).Range("A" & i).Value)
     obj.FindElementByName("lname").SendKeys (ThisWorkbook.Sheets(Name).Range("B" & i).Value)
     'obj.FindElementByXpath("//*[@value='m']").Click
     'obj.FindElementByXpath("//*[@value='Reading']").Click
     'obj.FindElementByName("occupation").AsSelect.selectByText ("Business")
     'obj.FindElementByName("day").AsSelect.selectByText ("31")
     'obj.FindElementByName("month").AsSelect.selectByText ("Dec")
     'obj.FindElementByXpath("(//*[@name='day'])[2]").AsSelect.selectByText ("1991")
     obj.FindElementById("btnSubmit").Click
     obj.Wait 1000
     obj.FindElementByName("fname").Clear
     obj.FindElementByName("lname").Clear
     i = i + 1
        ElseIf IsEmpty(ThisWorkbook.Sheets(Name).Range("A" & i).Value) = True And IsEmpty(ThisWorkbook.Sheets(Name).Range("A" & i + 1).Value) = False Then
        i = i + 1
        Else: Exit Do
        End If
        Loop
    Next Name

End Sub
