VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExchengeRateItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'https://bank.gov.ua/NBU_Exchange/exchange?date=31.12.2019

'<StartDate>31.12.2019</StartDate>
'<TimeSign>0000</TimeSign>
'<CurrencyCode>949</CurrencyCode>
'<CurrencyCodeL>TRY</CurrencyCodeL>
'<Units>1</Units>
'<Amount>3.9757</Amount>

Private pStartDate As Date

Private pTimeSign As String
Private pCurrencyCodeL As String

Private pCurrencyCode As Long
Private pUnits As Long

Private pAmount As Variant ' decimal

Public Property Get StartDate() As Date
    StartDate = pStartDate
End Property
Public Property Let StartDate(ByVal vNewValue As Date)
    pStartDate = CDate(vNewValue)
End Property

Public Property Get TimeSign() As String
    TimeSign = pTimeSign
End Property
Public Property Let TimeSign(ByVal vNewValue As String)
    pTimeSign = vNewValue
End Property


Public Property Get CurrencyCode() As Long
    CurrencyCode = pCurrencyCode
End Property
Public Property Let CurrencyCode(ByVal vNewValue As Long)
    pCurrencyCode = CLng(vNewValue)
End Property

Public Property Get CurrencyCodeL() As String
    CurrencyCodeL = pCurrencyCodeL
End Property
Public Property Let CurrencyCodeL(ByVal vNewValue As String)
    pCurrencyCodeL = vNewValue
End Property

Public Property Get Units() As Long
    Units = pUnits
End Property
Public Property Let Units(ByVal vNewValue As Long)
    pUnits = CLng(vNewValue)
End Property

Public Property Get Amount() As Variant 'decimal
    Amount = pAmount
End Property
Public Property Let Amount(ByVal vNewValue As Variant)
    pAmount = CDec(vNewValue)
End Property
