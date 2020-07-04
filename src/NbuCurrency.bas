Attribute VB_Name = "NbuCurrency"
Option Explicit

Public Function NbuTodayCurrency(currCode As String, volatileArg As Variant) As Variant 'decimal
    NbuTodayCurrency = GetNbuCurrency(currCode, Now)
End Function

Public Function GetNbuCurrency(currCode As String, selectedDate As Date) As Variant 'decimal
    Dim col As Collection
    Set col = GetCurrency(selectedDate)
    
    If col.Count = 0 Then
        Exit Function
    End If
    
    Dim foundItem As ExchengeRateItem
    Dim v As ExchengeRateItem
    For Each v In col
        If v.CurrencyCodeL = currCode Then
            Set foundItem = v
            Exit For
        End If
    Next
    
    If foundItem Is Nothing Then
        Dim errText As String
        errText = "Not found by currency code " & currCode & " on " & selectedDate & "!"
        Debug.Print errText
        Exit Function
        'If you want to have #VALUE! error on invalid currency code, _
         remove "Exit Function" call and uncomment next line with "Err.Raise"
         
        'Err.Raise 555, Description:=errText
    End If
    
    GetNbuCurrency = foundItem.Amount / foundItem.Units
    
    Debug.Print "Currency " & currCode & " loaded on " & selectedDate & ", value is " + CStr(foundItem.Amount)
End Function

Private Sub TestCurrLoad()
    Dim col As Collection
    Set col = GetCurrency(#12/31/2019#)
    
    Dim v As Variant
    v = GetNbuCurrency("USD", #12/31/2019#)
    Debug.Assert v = CDec(23.6862)
    
    Stop
End Sub


Private Function GetCurrency(currencyDate As Date) As Collection

    Dim resultXmlDocument As DOMDocument60
    Dim recordItem As IXMLDOMElement
    Dim col As New Collection
    Dim resultItem As ExchengeRateItem
    
    Set resultXmlDocument = RequestGetXml(currencyDate)
    
    For Each recordItem In resultXmlDocument.LastChild.ChildNodes
        Set resultItem = New ExchengeRateItem
        With resultItem
            .Amount = ParseDecimal(recordItem.SelectSingleNode("Amount").Text)
            .CurrencyCode = CLng(recordItem.SelectSingleNode("CurrencyCode").Text)
            .CurrencyCodeL = recordItem.SelectSingleNode("CurrencyCodeL").Text
            .StartDate = CDate(recordItem.SelectSingleNode("StartDate").Text)
            .TimeSign = recordItem.SelectSingleNode("TimeSign").Text
            .Units = CLng(recordItem.SelectSingleNode("Units").Text)
        End With
        col.Add resultItem
    Next
    
    Set GetCurrency = col

End Function

Private Function RequestGetXml(currencyDate As Date) As DOMDocument60
    Dim XMLHTTP As New MSXML2.XMLHTTP60
    Dim requestUrl As String
    
    requestUrl = "https://bank.gov.ua/NBU_Exchange/exchange?date=" & _
        Day(currencyDate) & "." & _
        Month(currencyDate) & "." & _
        Year(currencyDate)
        
    XMLHTTP.Open "GET", requestUrl, False
    XMLHTTP.send
    
    Dim resultXmlDocument As DOMDocument60
    Set resultXmlDocument = XMLHTTP.responseXML
    
    Set RequestGetXml = resultXmlDocument
End Function

Private Function ParseDecimal(inputString As String) As Variant
    ParseDecimal = CDec(Replace(inputString, ".", Application.DecimalSeparator))
End Function
