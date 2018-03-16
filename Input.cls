VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Public Sub getPrice(strStock As String)
    Dim appIE As Object, JSON As Object, pDate As Variant, objDates As Variant, price As Variant, index As Integer, arrPrices As Variant, objPrices As Object, scriptControl As Object, ws As Object, script As Object, strStockJS As String, doc As Object, scripts As Object, i As Integer
    Set appIE = CreateObject("internetexplorer.application")
    Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    'VBA honestly sucks with JSON, but this AddCode is pretty cool as it lets me work with JavaScript'
    scriptControl.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    scriptControl.AddCode "function getDates(jsonObj) { var dates = jsonObj['context']['dispatcher']['stores']['HistoricalPriceStore']['prices']; var new_dates = [];  for(var i =0; i < dates.length; i++){if(dates[i]['close'])new_dates.push(dates[i]['date'])}; return new_dates}"
    scriptControl.AddCode "function getPrices(jsonObj) { var dates = jsonObj['context']['dispatcher']['stores']['HistoricalPriceStore']['prices']; var new_prices = [];  for(var i =0; i < dates.length; i++){if(dates[i]['close'])new_prices.push(dates[i]['close'])}; return new_prices}"
    scriptControl.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
    'scriptControl.AddCode "function getProperty(jsonObj, propertyName)"
    With appIE
        .Navigate "https://finance.yahoo.com/quote/" & strStock & "/history/"
        .Visible = False
    End With
    Do While appIE.Busy
        DoEvents
    Loop
    Set doc = appIE.document
    appIE.Quit
    strStockJS = doc.GetElementsByTagName("script")(57).innerHTML
    'Ever since finance.yahoo.com got rid of their API, they've stored all the data in a react object. Have to find it then do splits to get the data. Kind of a pain'
    strStockJS = Split(Split(strStockJS, "root.App.main = ")(1), "(this)")(0)
    strStockJS = Left(strStockJS, Len(strStockJS) - 3)
    With ThisWorkbook
        Set ws = .Worksheets.Add(After:=.Sheets(.Sheets.Count))
    End With
    ws.Name = strStock & " prices"
    ws.Activate
    ws.Cells(1, 2) = "Price"
    ws.Cells(1, 1) = "Date"
    Set JSON = scriptControl.Eval("(" + strStockJS + ")")
    Set objPrices = scriptControl.Run("getPrices", JSON)
    Set objDates = scriptControl.Run("getDates", JSON)
    index = 2
    For Each price In objPrices
        ws.Cells(index, 2).Value = price
        index = index + 1
    Next
    index = 2
    For Each pDate In objDates
        ws.Cells(index, 1).Formula = "=(((" & pDate & "/60)/60)/24)+DATE(1970,1,1)"
        ws.Cells(index, 1).Value = Format(ws.Cells(index, 1).Value, "Short Date")
        index = index + 1
    Next
End Sub

Private Sub getStock_Click()
    If IsEmpty(TextBox1.Value) = True Then
        MsgBox ("Please enter a stock")
    Else
        Application.DisplayAlerts = False
        Application.ScreenUpdating = True
        Call deleteSheets
        Call getPrice(TextBox1.Value)
        Call createChart(TextBox1.Value)
        Application.ScreenUpdating = False
        Application.DisplayAlerts = True
    End If
    
End Sub

Private Sub TextBox1_Change()

End Sub

Public Sub deleteSheets()
    Dim xWs As Worksheet
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "Input" Then
            xWs.Delete
        End If
    Next
End Sub

Public Sub createChart(strStock As String)
    Dim d1 As Range: Set d1 = ActiveSheet.Range("A2:A30")
    Dim d2 As Range: Set d2 = ActiveSheet.Range("B2:B30")
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim chsource As Range
    Set chsource = ws.Range("A2:B30")
    '~~> Clean up existing chart
    Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
        If sh.Type = msoChart Then sh.Delete
    Next
    '~~> Create the chart
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=ActiveSheet.Range("A2:B30"), PlotBy:= _
    xlColumns
    ActiveChart.ChartType = xlLine
    ActiveChart.SeriesCollection(1).Name = strStock & "'s stock price"
End Sub

