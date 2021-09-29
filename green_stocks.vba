Sub MacroCheck()

Dim testMessage As String

testMessage = "Hello Jorge"

MsgBox (testMessage)

End Sub


Sub DQAnalysis1()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

End Sub


Sub DQAnalysis2()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"



End Sub
