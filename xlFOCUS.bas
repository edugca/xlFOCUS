Attribute VB_Name = "xlFOCUS"
''' xlFOCUS
'''
''' This module contains routines for fetching data from the BCB FOCUS webservice
''' For these routines to work, it requires VBA-JSON routines developed by Tim Hall and available at https://github.com/VBA-tools/VBA-JSON. Tested on v2.3.1.
''' To put these resources to work on your own spreadsheet, copy this module and VBA-JSON's module named 'JsonConverter' to your spreadsheet
'''
''' Available on:
''' Developed by Eduardo G. C. Amaral
''' Version: 0.1
''' Last update: 2021-12-18
'''
''' It is intended to help researchers and the general public, so have fun, but use at your own risk!

Option Explicit

Private Const recalculateWhenFunctionWizardIsOpen = False

Private Function xlFOCUS_CheckArguments(Optional ByRef Indicador As String, Optional ByRef IndicadorDetalhe As String, _
    Optional ByRef DataReferencia As Variant, _
    Optional ByRef DataInicial As Variant, _
    Optional ByRef DataFinal As Variant, _
    Optional ByRef baseCalculo As String, _
    Optional ByRef tipoCalculo As String, _
    Optional ByRef Suavizada As String, _
    Optional ByRef Instituicao As Variant, _
    Optional ByRef Periodicidade As String, _
    Optional ByRef Campos As Variant) As String

Dim result As String

'Force range to value or array
DataReferencia = DataReferencia
DataInicial = DataInicial
DataFinal = DataFinal
Instituicao = Instituicao

If LenB(Indicador) = 0 Then
    result = "# Indicador não especificado"
    GoTo Final
End If

If IsMissing(DataInicial) Or IsEmpty(DataInicial) Then
    DataInicial = ""
ElseIf IsNumeric(DataInicial) Or IsDate(DataInicial) Then
    DataInicial = Format(Year(DataInicial), "0000") & "-" & Format(Month(DataInicial), "00") & "-" & Format(Day(DataInicial), "00")
End If

If IsMissing(DataFinal) Or IsEmpty(DataFinal) Then
    DataFinal = ""
ElseIf IsNumeric(DataFinal) Or IsDate(DataFinal) Then
    DataFinal = Format(Year(DataFinal), "0000") & "-" & Format(Month(DataFinal), "00") & "-" & Format(Day(DataFinal), "00")
End If

If IsArray(Campos) = False Then
    Campos = Array(Campos)
End If

result = "OK"

Final:

xlFOCUS_CheckArguments = result

End Function

Function xlFOCUS_ExpectativasMensais(Indicador As String, Optional IndicadorDetalhe As String, Optional DataReferencia As Variant, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional baseCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim Suavizada As String
Dim Instituicao As String
Dim Periodicidade As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasMensais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais?"

DataReferencia = Format(Month(DataReferencia), "00") & "%2F" & Format(Year(DataReferencia), "0000")

xlFOCUS_ExpectativasMensais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)
    
End Function

Function xlFOCUS_ExpectativasTop5Mensais(Indicador As String, Optional IndicadorDetalhe As String, Optional DataReferencia As Variant, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional tipoCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim baseCalculo As String
Dim Suavizada As String
Dim Instituicao As String
Dim Periodicidade As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasTop5Mensais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoTop5Mensais?"
    
DataReferencia = Format(Month(DataReferencia), "00") & "%2F" & Format(Year(DataReferencia), "0000")
    
xlFOCUS_ExpectativasTop5Mensais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)

End Function

Function xlFOCUS_ExpectativasTrimestrais(Indicador As String, Optional IndicadorDetalhe As String, Optional DataReferencia As Variant, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional baseCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim trimestre As Long
Dim Suavizada As String
Dim Instituicao As String
Dim Periodicidade As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasTrimestrais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoTrimestrais?"

If IsNumeric(DataReferencia) Or IsDate(DataReferencia) Then
    trimestre = Application.WorksheetFunction.RoundUp(Month(DataReferencia) / 3, 0)
    DataReferencia = Format(trimestre, "0") & "%2F" & Format(Year(DataReferencia), "0000")
End If

xlFOCUS_ExpectativasTrimestrais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)
    
End Function

Function xlFOCUS_ExpectativasAnuais(Indicador As String, Optional IndicadorDetalhe As String, Optional DataReferencia As Variant, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional baseCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim Suavizada As String
Dim Instituicao As String
Dim Periodicidade As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasAnuais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?"

DataReferencia = Format(Year(DataReferencia), "0000")

xlFOCUS_ExpectativasAnuais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)

End Function

Function xlFOCUS_ExpectativasTop5Anuais(Indicador As String, IndicadorDetalhe As String, DataReferencia As Variant, _
    DataInicial As Variant, DataFinal As Variant, tipoCalculo As String, Campos As Variant) As Variant

Dim Sistema As String
Dim baseCalculo As String
Dim Suavizada As String
Dim Instituicao As String
Dim Periodicidade As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasTop5Anuais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoTop5Anuais?"

DataReferencia = Format(Year(DataReferencia), "0000")

xlFOCUS_ExpectativasTop5Anuais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)
    
End Function

Function xlFOCUS_Expectativas12Meses(Indicador As String, Optional IndicadorDetalhe As String, Optional Suavizada As String, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional baseCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim DataReferencia As String
Dim Instituicao As String
Dim Periodicidade As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_Expectativas12Meses = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoInflacao12Meses?"

xlFOCUS_Expectativas12Meses = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)

End Function

Function xlFOCUS_ExpectativasInstituicoes(Indicador As String, Optional IndicadorDetalhe As String, Optional DataReferencia As Variant, Optional Instituicao As String, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional Periodicidade As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim baseCalculo As String
Dim Suavizada As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasInstituicoes = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoInstituicoes?"

xlFOCUS_ExpectativasInstituicoes = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)

End Function

Private Function sistema_xlFOCUS_Expectativas(Sistema As String, Indicador As String, IndicadorDetalhe As String, DataReferencia As Variant, _
    DataInicial As Variant, DataFinal As Variant, baseCalculo As String, tipoCalculo As String, Suavizada As String, Instituicao As Variant, Periodicidade As String, Campos As Variant) As Variant

Dim URL As String
Dim jsonScript As String
Dim Indicador_str As String, IndicadorDetalhe_str As String, DataReferencia_str As String
Dim DataInicial_str As String, DataFinal_str As String, baseCalculo_str As String
Dim tipoCalculo_str As String, Campos_str As String, Suavizada_str As String
Dim Instituicao_str As String, Periodicidade_str As String
Dim result As Variant
Dim colData As Long
Dim iObs As Long
Dim Passed As String

''''''' First checks
Passed = xlFOCUS_CheckArguments(Indicador, _
    IndicadorDetalhe, _
    DataReferencia, _
    DataInicial, _
    DataFinal, _
    baseCalculo, _
    tipoCalculo, _
    Suavizada, _
    Instituicao, _
    Periodicidade, _
    Campos)

If Passed <> "OK" Then
    result = Passed
    GoTo Final
End If
'''''''''''''''''''''''''''''''''''

Indicador_str = Application.WorksheetFunction.EncodeURL(Indicador)
IndicadorDetalhe_str = Application.WorksheetFunction.EncodeURL(IndicadorDetalhe)
DataReferencia_str = CStr(DataReferencia)
DataInicial_str = CStr(DataInicial)
DataFinal_str = CStr(DataFinal)
baseCalculo_str = CStr(baseCalculo)
tipoCalculo_str = CStr(tipoCalculo)
Suavizada_str = CStr(Suavizada)
Instituicao_str = CStr(Instituicao)
Periodicidade_str = CStr(Periodicidade)

'Force array format
Campos = Application.Transpose(Application.Transpose(Campos))
Campos_str = Join(Campos, ",")

URL = Sistema _
        & "$top=10000" _
        & "&$filter=Indicador%20eq%20'" & Indicador_str & "'" _
        & IIf(LenB(IndicadorDetalhe_str) = 0, "", "%20and%20IndicadorDetalhe%20eq%20'" & IndicadorDetalhe_str & "'") _
        & IIf(LenB(DataReferencia_str) = 0, "", "%20and%20DataReferencia%20eq%20'" & DataReferencia_str & "'") _
        & IIf(LenB(DataInicial_str) = 0, "", "%20and%20Data%20ge%20'" & DataInicial_str & "'") _
        & IIf(LenB(DataFinal_str) = 0, "", "%20and%20Data%20le%20'" & DataFinal_str & "'") _
        & IIf(LenB(baseCalculo_str) = 0, "", "%20and%20baseCalculo%20eq%20" & baseCalculo_str) _
        & IIf(LenB(tipoCalculo_str) = 0, "", "%20and%20tipoCalculo%20eq%20'" & tipoCalculo_str & "'") _
        & IIf(LenB(Suavizada_str) = 0, "", "%20and%20Suavizada%20eq%20'" & Suavizada_str & "'") _
        & IIf(LenB(Instituicao_str) = 0, "", "%20and%20Instituicao%20eq%20" & Instituicao_str) _
        & IIf(LenB(Periodicidade_str) = 0, "", "%20and%20Periodicidade%20eq%20'" & Periodicidade_str & "'") _
        & "&$format=json" _
        & "&$select=" & Campos_str

jsonScript = Application.WorksheetFunction.WebService(URL)
result = xlFOCUS_ReadJSON(jsonScript, False, Campos)

'Check returned values
If VarType(result) = vbString Then
    GoTo Final
End If

'Format values
colData = -1
On Error Resume Next
colData = -1 + Application.WorksheetFunction.Match("Data", Campos, 0)
On Error GoTo 0

If colData > -1 Then
    For iObs = 0 To UBound(result, 1)
        result(iObs, colData) = DateValue(result(iObs, colData))
    Next iObs
End If


Final:

sistema_xlFOCUS_Expectativas = result

End Function

Function xlFOCUS_ReadJSON(JsonText As String, Optional showHeaders As Boolean = False, Optional Fields As Variant) As Variant

Dim result As Variant
Dim Parsed As Scripting.Dictionary
Dim nCols As Long, jCol As Long
Dim nameCols As Variant
Dim Val As Variant
Dim colNamesStart As Long
Dim Value As Dictionary
Dim i As Long
Dim Values As Variant

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ReadJSON = "# Barra de fórmulas aberta"
    Exit Function
End If

' Parse json to Dictionary
' "values" is parsed as Collection
' each item in "values" is parsed as Dictionary
Set Parsed = JsonConverter.ParseJson(JsonText)

' Check structure
If Parsed("value").Count > 0 Then
    nCols = Parsed("value")(1).Count - 1
    nameCols = Parsed("value")(1).Keys
    
    If Not IsMissing(Fields) Then
        colNamesStart = LBound(Fields)
        If nCols + 1 <> Application.WorksheetFunction.CountA(Fields) Then
            result = "# Ao menos um campo está errado"
            GoTo Final
        End If
    Else
        colNamesStart = 0
        Fields = nameCols
    End If
Else
    result = "# Consulta retornou vazia"
    GoTo Final
End If

' Prepare and write values to sheet
If showHeaders Then
    ReDim Values(Parsed("value").Count, nCols)
    i = 1
    
    'Fill in header
    For jCol = 0 To nCols
        Values(0, jCol) = Fields(colNamesStart + jCol)
    Next jCol
Else
    ReDim Values(Parsed("value").Count - 1, nCols)
    i = 0
End If


For Each Value In Parsed("value")
    For jCol = 0 To nCols
        Val = Value(Fields(colNamesStart + jCol))
        Val = IIf(IsNull(Val), "", Val)
        Values(i, jCol) = Val
    Next jCol
    i = i + 1
Next Value

result = Values

Final:

xlFOCUS_ReadJSON = result

End Function

Function xlFOCUS_IfError(ValueToBeChecked As Variant, ValueToReturnInCaseOfError As Variant) As Variant

If IsError(ValueToBeChecked) Or VarType(ValueToBeChecked) = vbString Then
    
    xlFOCUS_IfError = ValueToReturnInCaseOfError
    Exit Function
    
End If

xlFOCUS_IfError = ValueToBeChecked

End Function
