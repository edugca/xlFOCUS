Attribute VB_Name = "xlFOCUS"
''' xlFOCUS
''' https://github.com/edugca/xlFOCUS
'''
''' This module contains routines for fetching economic data from webservices
''' For these routines to work, it requires VBA-JSON routines developed by Tim Hall and available at https://github.com/VBA-tools/VBA-JSON. Tested on v2.3.1.
''' To put these resources to work on your own spreadsheet, copy this module and the module named 'xlFOCUS_JsonConverter' to your spreadsheet
'''
''' Available on:
''' Developed by Eduardo G. C. Amaral
''' Version: 0.6
''' Last update: 2022-08-25
'''
''' It is intended to help researchers and the general public, so have fun, but use at your own risk!
'''
''' Copyright (c) 2022, Eduardo G. C. Amaral
''' All rights reserved.
'''
''' Redistribution and use in source and binary forms, with or without
''' modification, are permitted provided that the following conditions are met:
'''     * Redistributions of source code must retain the above copyright
'''       notice, this list of conditions and the following disclaimer.
'''     * Redistributions in binary form must reproduce the above copyright
'''       notice, this list of conditions and the following disclaimer in the
'''       documentation and/or other materials provided with the distribution.
'''     * Neither the name of the <organization> nor the
'''       names of its contributors may be used to endorse or promote products
'''       derived from this software without specific prior written permission.
'''
''' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
''' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
''' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
''' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
''' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
''' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
''' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
''' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
''' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
''' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
''' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

Private Const recalculateWhenFunctionWizardIsOpen = False

Function xlFOCUS_ipeadata_Metadados(Optional SERCODIGO As String, Optional NIVNOME As String, Optional TERCODIGO As String, Optional Campos As Variant) As Variant

'Details about the webservice: http://www.ipeadata.gov.br/api/

Dim URL As String
Dim jsonScript As String
Dim SERCODIGO_str As String, NIVNOME_str As String, TERCODIGO_str As String
Dim result As Variant
Dim colData As Long
Dim iObs As Long
Dim datesVector() As Long, seqDates() As Variant
Dim minDateIdx As Long, maxDateIdx As Long
Dim dateAux As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ipeadata_Metadados = "# Barra de fórmulas aberta"
    Exit Function
End If

SERCODIGO_str = Application.WorksheetFunction.EncodeURL(CStr(SERCODIGO))
NIVNOME_str = Application.WorksheetFunction.EncodeURL(CStr(NIVNOME))
TERCODIGO_str = Application.WorksheetFunction.EncodeURL(CStr(TERCODIGO))

'Check arguments
If (Len(SERCODIGO_str) = 0 And Len(NIVNOME_str) = 0 And Len(TERCODIGO_str) = 0) Or _
    (Len(SERCODIGO_str) > 0 And (Len(NIVNOME_str) > 0 Or Len(TERCODIGO_str) > 0)) Or _
    (Len(SERCODIGO_str) = 0 And (Len(NIVNOME_str) > 0 Xor Len(TERCODIGO_str) > 0)) Then
    
    xlFOCUS_ipeadata_Metadados = "# Preencha apenas SERCODIGO ou o par NIVNOME e TERCODIGO"
    Exit Function

End If

'Force array format
Campos = Application.Transpose(Application.Transpose(Campos))
If Not IsArray(Campos) Then
    Campos = Array(Campos)
End If

'Check URL
If Len(SERCODIGO) > 0 Then
    URL = "http://www.ipeadata.gov.br/api/odata4/Metadados('" & SERCODIGO_str & "')"
Else
    URL = "http://www.ipeadata.gov.br/api/odata4/Territorios(TERCODIGO='" & TERCODIGO_str & "',NIVNOME='" & NIVNOME_str & "')"
End If

'If Not (LenB(DataInicial) = 0 And LenB(DataFinal) = 0) Then
'    If LenB(DataInicial) = 0 Then
'        result = "# Data inicial ausente"
'        GoTo Final
'    End If
'    If LenB(DataFinal) = 0 Then
'        result = "# Data final ausente"
'        GoTo Final
'    End If
'End If

'If LenB(DataInicial) <> 0 And LenB(DataFinal) <> 0 Then
'    URL = URL & "&dataInicial=" & DataInicial_str & "&dataFinal=" & DataFinal_str
'End If

'Campos = Array("VALDATA", "VALVALOR")

''''''''NEW METHOD
jsonScript = xlFOCUS_WEBSERVICE(URL)

''''''''OLD METHOD
'jsonScript = Application.WorksheetFunction.WebService(URL)

result = xlFOCUS_ipeadata_ReadJSON(jsonScript, False, Campos)

'Check returned values
If VarType(result) = vbString Then
    GoTo Final
End If

Final:

xlFOCUS_ipeadata_Metadados = result
    
End Function

Function xlFOCUS_ipeadata(SERCODIGO As String, Optional NIVNOME As String, Optional TERCODIGO As String, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, _
    Optional nUltimos As Variant, Optional RetornarDatas As Variant, Optional AlfaNumerico As Boolean = False) As Variant

'Details about the webservice: http://www.ipeadata.gov.br/api/

Dim URL As String
Dim jsonScript As String
Dim SERCODIGO_str As String, NIVNOME_str As String, TERCODIGO_str As String, nUltimos_str As String, DataInicial_str As String, DataFinal_str As String
Dim result As Variant
Dim colData As Long
Dim iObs As Long
Dim Campos As Variant
Dim datesVector() As Long, seqDates() As Variant
Dim minDateIdx As Long, maxDateIdx As Long
Dim dateAux As String
Dim allPeriods As Long
Dim nCols As Long
Dim listMatches As Variant
Dim colNIVNOME As Long, colTERCODIGO As Long
Dim oldDates As Boolean

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ipeadata = "# Barra de fórmulas aberta"
    Exit Function
End If

' Avoid recalculation when the function wizard is being used
If (Len(NIVNOME) > 0) <> (Len(TERCODIGO) > 0) Then
    xlFOCUS_ipeadata = "# NIVNOME e TERCODIGO devem ser conjuntamente fornecidos"
    Exit Function
End If

SERCODIGO_str = CStr(SERCODIGO)
NIVNOME_str = CStr(NIVNOME)
TERCODIGO_str = CStr(TERCODIGO)
DataInicial_str = Format(DataInicial, "dd/MM/yyyy")
DataFinal_str = Format(DataFinal, "dd/MM/yyyy")
nUltimos_str = CStr(nUltimos)

'Check URL
'Return observations between dates
If AlfaNumerico = False Then
    URL = "http://www.ipeadata.gov.br/api/odata4/Metadados('" & SERCODIGO_str & "')/Valores"
Else
    URL = "http://www.ipeadata.gov.br/api/odata4/Metadados('" & SERCODIGO_str & "')/ValoresStr"
End If

'If Not (LenB(DataInicial) = 0 And LenB(DataFinal) = 0) Then
'    If LenB(DataInicial) = 0 Then
'        result = "# Data inicial ausente"
'        GoTo Final
'    End If
'    If LenB(DataFinal) = 0 Then
'        result = "# Data final ausente"
'        GoTo Final
'    End If
'End If

'If LenB(DataInicial) <> 0 And LenB(DataFinal) <> 0 Then
'    URL = URL & "&dataInicial=" & DataInicial_str & "&dataFinal=" & DataFinal_str
'End If

If Len(NIVNOME_str) = 0 Then
    Campos = Array("VALDATA", "VALVALOR")
Else
    Campos = Array("NIVNOME", "TERCODIGO", "VALDATA", "VALVALOR")
End If

''''''''NEW METHOD
jsonScript = xlFOCUS_WEBSERVICE(URL)

''''''''OLD METHOD
'jsonScript = Application.WorksheetFunction.WebService(URL)

result = xlFOCUS_ipeadata_ReadJSON(jsonScript, False, Campos, AlfaNumerico)

'Check returned values
If VarType(result) = vbString Then
    GoTo Final
End If

'Filter territory
If Len(NIVNOME_str) = 0 Then
    'Rebase the array matrix to be 1-based
    result = Application.Index(result, 0, 0)
Else
    colNIVNOME = 0
    colTERCODIGO = 1
    For iObs = LBound(result, 1) To UBound(result, 1)
        If result(iObs, colNIVNOME) = NIVNOME_str And result(iObs, colTERCODIGO) = TERCODIGO_str Then
            If IsEmpty(listMatches) Then
                ReDim listMatches(1 To 1) As Long
            Else
                ReDim Preserve listMatches(1 To (UBound(listMatches) + 1))
            End If
            listMatches(UBound(listMatches)) = iObs + 1
        End If
    Next iObs
    If IsEmpty(listMatches) Then
        xlFOCUS_ipeadata = "# NIVNOME não encontrado"
        Exit Function
    End If
    result = Application.Index(result, Application.Transpose(listMatches), Array(3, 4))
    Campos = Array("VALDATA", "VALVALOR")
End If

'Format values
colData = -1
On Error Resume Next
colData = colData + 1 + Application.WorksheetFunction.Match("VALDATA", Campos, 0)
On Error GoTo 0

If colData > -1 Then
    ReDim Preserve datesVector(LBound(result, 1) To UBound(result, 1))
    For iObs = LBound(result, 1) To UBound(result, 1)
        dateAux = Left$(result(iObs, colData), 10)
        datesVector(iObs) = CLng(DateValue(dateAux))
    Next iObs
End If

'Slice dates
If Len(DataInicial_str) = 0 Then
    minDateIdx = 1
ElseIf datesVector(1) >= DateValue(DataInicial_str) Then
    minDateIdx = 1
Else
    minDateIdx = Application.WorksheetFunction.Match(CLng(DateValue(DataInicial_str)), datesVector, 1)
    If CLng(DateValue(DataInicial_str)) > datesVector(minDateIdx) Then
        minDateIdx = minDateIdx + 1
    End If
End If
If Len(DataFinal_str) = 0 Then
    maxDateIdx = UBound(datesVector)
ElseIf datesVector(UBound(datesVector)) <= CLng(DateValue(DataFinal_str)) Then
    maxDateIdx = UBound(datesVector)
Else
    maxDateIdx = Application.WorksheetFunction.Match(CLng(DateValue(DataFinal_str)), datesVector, 1)
End If
seqDates = Application.WorksheetFunction.Sequence(maxDateIdx - minDateIdx + 1, 1, minDateIdx)

'Define return table format
If IsMissing(RetornarDatas) Then
    'Only values
    nCols = 1
    colData = -1
    result = Application.Index(result, seqDates, 2)
ElseIf RetornarDatas = False Then
    'Only values
    nCols = 1
    colData = -1
    result = Application.Index(result, seqDates, 2)
ElseIf RetornarDatas = True Then
    'Only dates
    nCols = 1
    result = Application.Index(result, seqDates, 1)
ElseIf RetornarDatas = 2 Then
    'Dates and values
    nCols = 2
    'No need to change (but array must start at 1)
    result = Application.Index(result, seqDates, Array(1, 2))
Else
    result = "# RetornarDatas é inválido"
    GoTo Final
End If

'Slice periods
If Len(nUltimos_str) > 0 Then
    allPeriods = UBound(result, 1) - LBound(result, 1) + 1
    
    If CLng(nUltimos) > allPeriods Then nUltimos = allPeriods
    
    result = Application.Index(result, _
        Application.WorksheetFunction.Sequence(CLng(nUltimos), 1, allPeriods - CLng(nUltimos) + 1, 1), _
        Application.WorksheetFunction.Sequence(1, nCols, 1, 1))
End If

'Format values
oldDates = False
If colData > 0 Then
    'Check whether it is a singleton
    If minDateIdx = maxDateIdx Then
        result(colData) = DateValue(result(colData))
    Else
        For iObs = 1 To UBound(result, 1)
            dateAux = Left$(result(iObs, colData), 10)
            'Check whether dates are older than Excel's first date
            If oldDates = False And DateValue(dateAux) > DateValue("1900-01-01") Then
                result(iObs, colData) = DateValue(dateAux)
            Else
                result(iObs, colData) = dateAux
                oldDates = True
            End If
        Next iObs
    End If
End If


Final:

xlFOCUS_ipeadata = result
    
End Function

Function xlFOCUS_SGS(Codigo As Long, Optional DataInicial As Variant, Optional DataFinal As Variant, _
    Optional nUltimos As Variant = "", Optional RetornarDatas As Variant = False) As Variant

Dim URL As String
Dim jsonScript As String
Dim Codigo_str As String, nUltimos_str As String, DataInicial_str As String, DataFinal_str As String
Dim result As Variant
Dim colData As Long
Dim iObs As Long
Dim Campos As Variant
Dim datesVector() As Long, seqDates() As Variant
Dim minDateIdx As Long, maxDateIdx As Long

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_SGS = "# Barra de fórmulas aberta"
    Exit Function
End If

Codigo_str = CStr(Codigo)
DataInicial_str = Format(DataInicial, "dd/MM/yyyy")
DataFinal_str = Format(DataFinal, "dd/MM/yyyy")
nUltimos_str = CStr(nUltimos)

'Check URL
If LenB(nUltimos_str) = 0 Then
    'Return observations between dates
    URL = "http://api.bcb.gov.br/dados/serie/bcdata.sgs." & Codigo_str & "/dados?formato=json"
    
    'If Not (LenB(DataInicial) = 0 And LenB(DataFinal) = 0) Then
    '    If LenB(DataInicial) = 0 Then
    '        result = "# Data inicial ausente"
    '        GoTo Final
    '    End If
    '    If LenB(DataFinal) = 0 Then
    '        result = "# Data final ausente"
    '        GoTo Final
    '    End If
    'End If
    
    If LenB(DataInicial) <> 0 And LenB(DataFinal) <> 0 Then
        URL = URL & "&dataInicial=" & DataInicial_str & "&dataFinal=" & DataFinal_str
    End If
    
Else
    'Return N last observations
    URL = "http://api.bcb.gov.br/dados/serie/bcdata.sgs." & Codigo_str & "/dados/ultimos/" & nUltimos_str & "?formato=json"
End If

Campos = Array("data", "valor")

''''''''NEW METHOD
jsonScript = xlFOCUS_WEBSERVICE(URL)

''''''''OLD METHOD
'jsonScript = Application.WorksheetFunction.WebService(URL)

result = xlFOCUS_SGS_ReadJSON(jsonScript, False, Campos)

'Check returned values
If VarType(result) = vbString Then
    GoTo Final
End If

'Format values
colData = -1
On Error Resume Next
colData = colData + Application.WorksheetFunction.Match("Data", Campos, 0)
On Error GoTo 0

If colData > -1 Then
    ReDim Preserve datesVector(0 To UBound(result, 1))
    For iObs = 0 To UBound(result, 1)
        datesVector(iObs) = CLng(DateValue(result(iObs, colData)))
    Next iObs
End If

'Slice dates
If datesVector(0) >= CDbl(DataInicial) Or Len(DataInicial_str) = 0 Then
    minDateIdx = 1
Else
    minDateIdx = Application.WorksheetFunction.Match(CDbl(DataInicial), datesVector, 1)
    If CDbl(DataInicial) > datesVector(-1 + minDateIdx) Then
        minDateIdx = minDateIdx + 1
    End If
End If
If datesVector(UBound(datesVector)) <= CDbl(DataFinal) Or CDbl(DataFinal) = 0 Then
    maxDateIdx = UBound(datesVector) + 1
Else
    maxDateIdx = Application.WorksheetFunction.Match(CDbl(DataFinal), datesVector, 1)
End If
seqDates = Application.WorksheetFunction.Sequence(maxDateIdx - minDateIdx + 1, 1, minDateIdx)

'Define return table format
If IsMissing(RetornarDatas) Then
    'Only values
    result = Application.Index(result, seqDates, 2)
ElseIf RetornarDatas = False Then
    'Only values
    result = Application.Index(result, seqDates, 2)
ElseIf RetornarDatas = True Then
    'Only dates
    colData = colData + 1
    result = Application.Index(result, seqDates, 1)
ElseIf RetornarDatas = 2 Then
    'Dates and values
    'No need to change (but array must start at 1)
    colData = colData + 1
    result = Application.Index(result, seqDates, Array(1, 2))
Else
    result = "# RetornarDatas é inválido"
    GoTo Final
End If

'Format values
If colData > 0 Then
    'Check whether it is a singleton
    If minDateIdx = maxDateIdx Then
        result(colData) = DateValue(result(colData))
    Else
        For iObs = 1 To UBound(result, 1)
            result(iObs, colData) = DateValue(result(iObs, colData))
        Next iObs
    End If
End If

Final:

xlFOCUS_SGS = result
    
End Function

Private Function xlFOCUS_CheckArguments(Optional ByRef Indicador As String, Optional ByRef IndicadorDetalhe As String, _
    Optional ByRef DataReferencia As Variant, _
    Optional ByRef Reuniao As Variant, _
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
Reuniao = Reuniao
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
Dim Reuniao As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasMensais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais?"

DataReferencia = Format(Month(DataReferencia), "00") & "%2F" & Format(Year(DataReferencia), "0000")

xlFOCUS_ExpectativasMensais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, Reuniao, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)
    
End Function

Function xlFOCUS_ExpectativasTop5Mensais(Indicador As String, Optional IndicadorDetalhe As String, Optional DataReferencia As Variant, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional tipoCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim baseCalculo As String
Dim Suavizada As String
Dim Instituicao As String
Dim Periodicidade As String
Dim Reuniao As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasTop5Mensais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoTop5Mensais?"
    
DataReferencia = Format(Month(DataReferencia), "00") & "%2F" & Format(Year(DataReferencia), "0000")
    
xlFOCUS_ExpectativasTop5Mensais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, Reuniao, _
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
Dim Reuniao As String

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

xlFOCUS_ExpectativasTrimestrais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, Reuniao, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)
    
End Function

Function xlFOCUS_ExpectativasAnuais(Indicador As String, Optional IndicadorDetalhe As String, Optional DataReferencia As Variant, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional baseCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim Suavizada As String
Dim Instituicao As String
Dim Periodicidade As String
Dim Reuniao As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasAnuais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?"

DataReferencia = Format(Year(DataReferencia), "0000")

xlFOCUS_ExpectativasAnuais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, Reuniao, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)

End Function

Function xlFOCUS_ExpectativasTop5Anuais(Indicador As String, IndicadorDetalhe As String, DataReferencia As Variant, _
    DataInicial As Variant, DataFinal As Variant, tipoCalculo As String, Campos As Variant) As Variant

Dim Sistema As String
Dim baseCalculo As String
Dim Suavizada As String
Dim Instituicao As String
Dim Periodicidade As String
Dim Reuniao As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasTop5Anuais = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoTop5Anuais?"

DataReferencia = Format(Year(DataReferencia), "0000")

xlFOCUS_ExpectativasTop5Anuais = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, Reuniao, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)
    
End Function

Function xlFOCUS_Expectativas12Meses(Indicador As String, Optional IndicadorDetalhe As String, Optional Suavizada As String, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional baseCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim DataReferencia As String
Dim Instituicao As String
Dim Periodicidade As String
Dim Reuniao As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_Expectativas12Meses = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoInflacao12Meses?"

xlFOCUS_Expectativas12Meses = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, Reuniao, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)

End Function

Function xlFOCUS_ExpectativasMercadoSelic(Indicador As String, Optional Reuniao As String, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional baseCalculo As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim DataReferencia As String
Dim Instituicao As String
Dim Periodicidade As String
Dim IndicadorDetalhe As String
Dim Suavizada As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasMercadoSelic = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoSelic?"

xlFOCUS_ExpectativasMercadoSelic = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, Reuniao, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)

End Function

Function xlFOCUS_ExpectativasInstituicoes(Indicador As String, Optional IndicadorDetalhe As String, Optional DataReferencia As Variant, Optional Instituicao As String, _
    Optional DataInicial As Variant, Optional DataFinal As Variant, Optional Periodicidade As String, Optional Campos As Variant) As Variant

Dim Sistema As String
Dim tipoCalculo As String
Dim baseCalculo As String
Dim Suavizada As String
Dim Reuniao As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ExpectativasInstituicoes = "# Barra de fórmulas aberta"
    Exit Function
End If

Sistema = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoInstituicoes?"

xlFOCUS_ExpectativasInstituicoes = sistema_xlFOCUS_Expectativas(Sistema, Indicador, IndicadorDetalhe, DataReferencia, Reuniao, _
    DataInicial, DataFinal, baseCalculo, tipoCalculo, Suavizada, Instituicao, Periodicidade, Campos)

End Function

Private Function sistema_xlFOCUS_Expectativas(Sistema As String, Indicador As String, IndicadorDetalhe As String, DataReferencia As Variant, Reuniao As Variant, _
    DataInicial As Variant, DataFinal As Variant, baseCalculo As String, tipoCalculo As String, Suavizada As String, Instituicao As Variant, Periodicidade As String, Campos As Variant) As Variant

Dim URL As String
Dim jsonScript As String
Dim Indicador_str As String, IndicadorDetalhe_str As String, DataReferencia_str As String, Reuniao_str As String
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
    Reuniao, _
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
Reuniao_str = CStr(Reuniao)
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
        & IIf(LenB(Reuniao_str) = 0, "", "%20and%20Reuniao%20eq%20'" & Reuniao_str & "'") _
        & IIf(LenB(DataInicial_str) = 0, "", "%20and%20Data%20ge%20'" & DataInicial_str & "'") _
        & IIf(LenB(DataFinal_str) = 0, "", "%20and%20Data%20le%20'" & DataFinal_str & "'") _
        & IIf(LenB(baseCalculo_str) = 0, "", "%20and%20baseCalculo%20eq%20" & baseCalculo_str) _
        & IIf(LenB(tipoCalculo_str) = 0, "", "%20and%20tipoCalculo%20eq%20'" & tipoCalculo_str & "'") _
        & IIf(LenB(Suavizada_str) = 0, "", "%20and%20Suavizada%20eq%20'" & Suavizada_str & "'") _
        & IIf(LenB(Instituicao_str) = 0, "", "%20and%20Instituicao%20eq%20" & Instituicao_str) _
        & IIf(LenB(Periodicidade_str) = 0, "", "%20and%20Periodicidade%20eq%20'" & Periodicidade_str & "'") _
        & "&$format=json" _
        & "&$select=" & Campos_str

''''''''NEW METHOD
jsonScript = xlFOCUS_WEBSERVICE(URL)

''''''''OLD METHOD
'jsonScript = Application.WorksheetFunction.WebService(URL)

result = xlFOCUS_ReadJSON(jsonScript, False, Campos, "value")

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

Function xlFOCUS_ReadJSON(JsonText As String, Optional showHeaders As Boolean = False, Optional Fields As Variant = "", Optional subField As Variant) As Variant

Dim result As Variant
Dim Parsed As Object
Dim nCols As Long, jCol As Long
Dim nameCols As Variant
Dim Val As Variant
Dim colNamesStart As Long
Dim Value As Object, ValueVar As Variant
Dim i As Long
Dim Values As Variant
Dim dicItems As Variant, dicKeys As Variant
Dim iField As Variant

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ReadJSON = "# Barra de fórmulas aberta"
    Exit Function
End If

'Force to be array
If Not IsMissing(Fields) Then
    If Not IsArray(Fields) Then
        Fields = Array(Fields)
    End If
End If

' Parse json to Dictionary
' "values" is parsed as Collection
' each item in "values" is parsed as Dictionary
Set Parsed = xlFOCUS_JsonConverter.ParseJson(JsonText)

' Check structure
If Not IsMissing(subField) Then
    If IsArray(subField) Then
        For Each iField In subField
            If Parsed(iField).Count > 0 Then
                Set Parsed = Parsed(iField)
            Else
                result = "# Consulta retornou vazia"
                GoTo Final
            End If
        Next iField
    Else
        If Len(subField) > 0 Then
            If Parsed(subField).Count > 0 Then
                Set Parsed = Parsed(subField)
            Else
                result = "# Consulta retornou vazia"
                GoTo Final
            End If
        End If
    End If
End If

If Parsed.Count > 0 Then
    'Check whether it is a collection or a dictionary
    If TypeName(Parsed) = "Collection" Then
        nCols = Parsed(1).Count - 1
        nameCols = Parsed(1).Keys
    Else
        nCols = 1
        nameCols = Array("1", "2")
    End If
    
    If Not IsMissing(Fields) Then
        colNamesStart = LBound(Fields)
        If nCols + 1 <> (UBound(Fields) - LBound(Fields) + 1) Then
            Fields = nameCols
            'result = "# Ao menos um campo está errado"
            'GoTo Final
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
    ReDim Values(Parsed.Count, nCols)
    i = 1
    
    'Fill in header
    For jCol = 0 To nCols
        Values(0, jCol) = Fields(colNamesStart + jCol)
    Next jCol
Else
    ReDim Values(Parsed.Count - 1, nCols)
    i = 0
End If

'Check whether it is a collection or a dictionary
If TypeName(Parsed) = "Collection" Then
    For Each Value In Parsed
        For jCol = 0 To nCols
            Val = Value(Fields(colNamesStart + jCol))
            Val = IIf(IsNull(Val), "", Val)
            Values(i, jCol) = Val
        Next jCol
        i = i + 1
    Next Value
Else
    dicItems = Parsed.Items
    dicKeys = Parsed.Keys
    For Each ValueVar In dicItems
        Val = dicKeys(i - CLng(showHeaders))
        Val = IIf(IsNull(Val), "", Val)
        Values(i, 0) = Val
        
        Val = ValueVar
        Val = IIf(IsNull(Val), "", Val)
        Values(i, 1) = Val
        
        i = i + 1
    Next ValueVar
End If

result = Values

Final:

xlFOCUS_ReadJSON = result

End Function

Function xlFOCUS_SGS_ReadJSON(JsonText As String, Optional showHeaders As Boolean = False, Optional Fields As Variant) As Variant

Dim result As Variant
Dim Parsed As Object
Dim nCols As Long, jCol As Long
Dim nameCols As Variant
Dim Val As Variant
Dim colNamesStart As Long
Dim Value As Object
Dim i As Long
Dim Values As Variant

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_SGS_ReadJSON = "# Barra de fórmulas aberta"
    Exit Function
End If

' Parse json to Dictionary
' "values" is parsed as Collection
' each item in "values" is parsed as Dictionary
Set Parsed = xlFOCUS_JsonConverter.ParseJson(JsonText)

' Check structure
If Parsed.Count > 0 Then
    nCols = UBound(Fields)
    nameCols = Parsed(1).Keys
Else
    result = "# Consulta retornou vazia"
    GoTo Final
End If

' Prepare and write values to sheet
If showHeaders Then
    ReDim Values(Parsed.Count, nCols)
    i = 1
    
    'Fill in header
    For jCol = 0 To nCols
        Values(0, jCol) = Fields(colNamesStart + jCol)
    Next jCol
Else
    ReDim Values(Parsed.Count - 1, nCols)
    i = 0
End If


For Each Value In Parsed
    For jCol = 0 To UBound(Fields)
        Val = Value(Fields(colNamesStart + jCol))

        'Check whether value is null
        If IsNull(Val) Then
            Values(i, jCol) = CVErr(xlErrNA)
        Else
            Values(i, jCol) = Val
            If Fields(colNamesStart + jCol) = "valor" Then
                Values(i, jCol) = VBA.Val((Values(i, jCol)))
            End If
        End If
        
    Next jCol
    i = i + 1
Next Value

result = Values

Final:

xlFOCUS_SGS_ReadJSON = result

End Function

Function xlFOCUS_ipeadata_ReadJSON(JsonText As String, Optional showHeaders As Boolean = False, Optional Fields As Variant, Optional alphaNumeric As Boolean = False) As Variant

Dim result As Variant
Dim Parsed As Object
Dim nCols As Long, jCol As Long
Dim nameCols As Variant
Dim Val As Variant
Dim colNamesStart As Long
Dim Value As Object
Dim i As Long
Dim Values As Variant

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ipeadata_ReadJSON = "# Barra de fórmulas aberta"
    Exit Function
End If

' Parse json to Dictionary
' "values" is parsed as Collection
' each item in "values" is parsed as Dictionary
Set Parsed = xlFOCUS_JsonConverter.ParseJson(JsonText)

' Check structure
If Parsed("value").Count > 0 Then
    nCols = UBound(Fields)
    nameCols = Parsed("value")(1).Keys
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
    For jCol = 0 To UBound(Fields)
        Val = Value(Fields(colNamesStart + jCol))
        
        'Check whether value is null
        If IsNull(Val) Then
            Values(i, jCol) = CVErr(xlErrNA)
        Else
            Values(i, jCol) = Val
            If alphaNumeric = False Then
                If Fields(colNamesStart + jCol) = "VALVALOR" Then
                    If VarType(Values(i, jCol)) = vbString Then
                        Values(i, jCol) = VBA.Val(Values(i, jCol))
                    End If
                End If
            End If
        End If
    Next jCol
    i = i + 1
Next Value

result = Values

Final:

xlFOCUS_ipeadata_ReadJSON = result

End Function

Function xlFOCUS_ReadJSONFile(JsonFilePath As String, Optional showHeaders As Boolean = False, Optional Fields As Variant, Optional subField As String) As Variant

Dim result As Variant
Dim JsonText As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_ReadJSONFile = "# Barra de fórmulas aberta"
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' CODE NOT USED ANYMORE
'
'Dim FSO As New Scripting.FileSystemObject
'Dim FileToRead As Scripting.TextStream
'Set FSO = CreateObject("Scripting.FileSystemObject")
'
'On Error Resume Next
'Set FileToRead = FSO.OpenTextFile(JsonFilePath, ForReading)
'If FileToRead Is Nothing Then
'    xlFOCUS_ReadJSONFile = "# Arquivo não localizado"
'    Exit Function
'End If
'On Error GoTo 0
'
'JsonText = FileToRead.ReadAll
'FileToRead.Close
'
''Clear memory
'Set FSO = Nothing
'Set FileToRead = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

JsonText = Read_UTF_8_Text_File(JsonFilePath)

result = xlFOCUS_ReadJSON(JsonText, showHeaders, Fields, subField)

Final:

xlFOCUS_ReadJSONFile = result

End Function

Function xlFOCUS_IfError(ValueToBeChecked As Variant, ValueToReturnInCaseOfError As Variant) As Variant

If IsError(ValueToBeChecked) Or VarType(ValueToBeChecked) = vbString Then
    
    xlFOCUS_IfError = ValueToReturnInCaseOfError
    Exit Function
    
End If

xlFOCUS_IfError = ValueToBeChecked

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' AUXILIARY FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function Read_UTF_8_Text_File(filePath As String)
'Adapted from https://www.ozgrid.com/forum/index.php?thread/107893-read-load-utf-8-text-file-with-vba/

Dim adoStream As Object
Dim var_String As Variant
Dim text As String

Set adoStream = CreateObject("ADODB.Stream")
adoStream.Charset = "UTF-8"
adoStream.Open
adoStream.LoadFromFile filePath

text = adoStream.ReadText

Set adoStream = Nothing

Read_UTF_8_Text_File = text
 
End Function

Function xlFOCUS_SCR_TaxasDeJurosDiario(Modalidade As String, Segmento As String, Optional InicioPeriodo As Variant, Optional FimPeriodo As Variant, _
    Optional Posicao As Variant, Optional InstituicaoFinanceira As Variant, Optional Campos As Variant) As Variant

Dim URL As String
Dim jsonScript As String
Dim Modalidade_str As String, Segmento_str As String, Posicao_str As String, InicioPeriodo_str As String, FimPeriodo_str As String, InstituicaoFinanceira_str As String, Campos_str As String
Dim result As Variant
Dim colData As Long
Dim iObs As Long

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_SCR_TaxasDeJurosDiario = "# Barra de fórmulas aberta"
    Exit Function
End If

Modalidade_str = Application.WorksheetFunction.EncodeURL(CStr(Modalidade))
Segmento_str = Application.WorksheetFunction.EncodeURL(CStr(Segmento))
Posicao_str = CStr(Posicao)
InicioPeriodo_str = Format(InicioPeriodo, "yyyy-MM-dd")
FimPeriodo_str = Format(FimPeriodo, "yyyy-MM-dd")
InstituicaoFinanceira_str = Application.WorksheetFunction.EncodeURL(CStr(InstituicaoFinanceira))

'Force array format
Campos = Application.Transpose(Application.Transpose(Campos))
Campos_str = Join(Campos, ",")

URL = "https://olinda.bcb.gov.br/olinda/servico/taxaJuros/versao/v2/odata/TaxasJurosDiariaPorInicioPeriodo?$format=json&$top=10000&$orderby=InicioPeriodo%20asc"

If LenB(Modalidade_str) <> 0 Then
    URL = URL & "&$filter=Modalidade%20eq%20'" & Modalidade_str & "'"
End If
If LenB(Segmento_str) <> 0 Then
    URL = URL & "%20and%20Segmento%20eq%20'" & Segmento_str & "'"
End If
If LenB(Posicao_str) <> 0 Then
    URL = URL & "%20and%20Posicao%20eq%20" & Posicao_str
End If
If LenB(InicioPeriodo_str) <> 0 Then
    URL = URL & "%20and%20InicioPeriodo%20ge%20'" & InicioPeriodo_str & "'"
End If
If LenB(FimPeriodo_str) <> 0 Then
    URL = URL & "%20and%20FimPeriodo%20le%20'" & FimPeriodo_str & "'"
End If
If LenB(InstituicaoFinanceira_str) <> 0 Then
    URL = URL & "%20and%20InstituicaoFinanceira%20eq%20'" & InstituicaoFinanceira_str & "'"
End If

If LenB(Campos_str) <> 0 Then
    URL = URL & "&$select=" & Campos_str
End If

''''''''NEW METHOD
jsonScript = xlFOCUS_WEBSERVICE(URL)

''''''''OLD METHOD
'jsonScript = Application.WorksheetFunction.WebService(URL)

result = xlFOCUS_ReadJSON(jsonScript, False, Campos, "value")

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

xlFOCUS_SCR_TaxasDeJurosDiario = result
    
End Function

Function xlFOCUS_SCR_TaxasDeJurosMensal(Modalidade As String, Optional InicioPeriodo As Variant, Optional FimPeriodo As Variant, _
    Optional Posicao As Variant, Optional InstituicaoFinanceira As Variant, Optional Campos As Variant) As Variant

Dim URL As String
Dim jsonScript As String
Dim Modalidade_str As String, Posicao_str As String, InicioPeriodo_str As String, FimPeriodo_str As String, InstituicaoFinanceira_str As String, Campos_str As String
Dim result As Variant
Dim colData As Long
Dim iObs As Long

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_SCR_TaxasDeJurosMensal = "# Barra de fórmulas aberta"
    Exit Function
End If

Modalidade_str = Application.WorksheetFunction.EncodeURL(CStr(Modalidade))
Posicao_str = CStr(Posicao)
InicioPeriodo_str = Format(InicioPeriodo, "yyyy-MM")
FimPeriodo_str = Format(FimPeriodo, "yyyy-MM")
InstituicaoFinanceira_str = Application.WorksheetFunction.EncodeURL(CStr(InstituicaoFinanceira))

'Force array format
Campos = Application.Transpose(Application.Transpose(Campos))
Campos_str = Join(Campos, ",")

URL = "https://olinda.bcb.gov.br/olinda/servico/taxaJuros/versao/v2/odata/TaxasJurosMensalPorMes?$format=json&$top=10000&$orderby=anoMes%20asc"

If LenB(Modalidade_str) <> 0 Then
    URL = URL & "&$filter=Modalidade%20eq%20'" & Modalidade_str & "'"
End If
If LenB(Posicao_str) <> 0 Then
    URL = URL & "%20and%20Posicao%20eq%20" & Posicao_str
End If
If LenB(InicioPeriodo_str) <> 0 Then
    URL = URL & "%20and%20anoMes%20ge%20'" & InicioPeriodo_str & "'"
End If
If LenB(FimPeriodo_str) <> 0 Then
    URL = URL & "%20and%20anoMes%20le%20'" & FimPeriodo_str & "'"
End If
If LenB(InstituicaoFinanceira_str) <> 0 Then
    URL = URL & "%20and%20InstituicaoFinanceira%20eq%20'" & InstituicaoFinanceira_str & "'"
End If

If LenB(Campos_str) <> 0 Then
    URL = URL & "&$select=" & Campos_str
End If

''''''''NEW METHOD
jsonScript = xlFOCUS_WEBSERVICE(URL)

''''''''OLD METHOD
'jsonScript = Application.WorksheetFunction.WebService(URL)

result = xlFOCUS_ReadJSON(jsonScript, False, Campos, "value")

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

xlFOCUS_SCR_TaxasDeJurosMensal = result
    
End Function

Function xlFOCUS_MercadoImobiliario(Codigo As String, Optional DataInicial As Variant, Optional DataFinal As Variant, _
    Optional nUltimos As Variant, Optional RetornarDatas As Variant) As Variant

Dim URL As String
Dim jsonScript As String
Dim Codigo_str As String, nUltimos_str As String, DataInicial_str As String, DataFinal_str As String
Dim result As Variant
Dim colData As Long
Dim iObs As Long
Dim Campos As Variant
Dim datesVector() As Long, seqDates() As Variant
Dim minDateIdx As Long, maxDateIdx As Long

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_MercadoImobiliario = "# Barra de fórmulas aberta"
    Exit Function
End If

Codigo_str = CStr(Codigo)
DataInicial_str = Format(DataInicial, "dd/MM/yyyy")
DataFinal_str = Format(DataFinal, "dd/MM/yyyy")
nUltimos_str = CStr(nUltimos)

'Check URL
'Return observations between dates
URL = "https://olinda.bcb.gov.br/olinda/servico/MercadoImobiliario/versao/v1/odata/mercadoimobiliario?$top=100&$format=json&$orderby=Data%20asc"

'If Not (LenB(DataInicial) = 0 And LenB(DataFinal) = 0) Then
'    If LenB(DataInicial) = 0 Then
'        result = "# Data inicial ausente"
'        GoTo Final
'    End If
'    If LenB(DataFinal) = 0 Then
'        result = "# Data final ausente"
'        GoTo Final
'    End If
'End If

URL = URL & "&$select=Data,Valor&$filter=" & "Info%20eq%20'" & Codigo_str & "'"

If LenB(DataInicial) <> 0 And LenB(DataFinal) <> 0 Then
    URL = URL & "%20and%20Data%20ge%20'" & DataInicial_str & "'%20and%20Data%20le%20'" & DataFinal_str & "'"
End If

Campos = Array("Data", "Valor")
''''''''NEW METHOD
jsonScript = xlFOCUS_WEBSERVICE(URL)

''''''''OLD METHOD
'jsonScript = Application.WorksheetFunction.WebService(URL)

result = xlFOCUS_ReadJSON(jsonScript, False, Campos, "value")

'Check returned values
If VarType(result) = vbString Then
    GoTo Final
End If

'Format values
colData = -1
On Error Resume Next
colData = colData + Application.WorksheetFunction.Match("Data", Campos, 0)
On Error GoTo 0

For iObs = 0 To UBound(result, 1)
    result(iObs, colData + 1) = Val(result(iObs, colData + 1))
Next iObs

If colData > -1 Then
    ReDim Preserve datesVector(0 To UBound(result, 1))
    For iObs = 0 To UBound(result, 1)
        datesVector(iObs) = CLng(DateValue(result(iObs, colData)))
    Next iObs
End If

'Slice dates
If datesVector(0) >= CDbl(DataInicial) Or Len(DataInicial_str) = 0 Then
    minDateIdx = 1
Else
    minDateIdx = Application.WorksheetFunction.Match(CDbl(DataInicial), datesVector, 1)
    If CDbl(DataInicial) > datesVector(-1 + minDateIdx) Then
        minDateIdx = minDateIdx + 1
    End If
End If
If datesVector(UBound(datesVector)) <= CDbl(DataFinal) Or CDbl(DataFinal) = 0 Then
    maxDateIdx = UBound(datesVector) + 1
Else
    maxDateIdx = Application.WorksheetFunction.Match(CDbl(DataFinal), datesVector, 1)
End If
seqDates = Application.WorksheetFunction.Sequence(maxDateIdx - minDateIdx + 1, 1, minDateIdx)

'Define return table format
If IsMissing(RetornarDatas) Then
    'Only values
    result = Application.Index(result, seqDates, 2)
ElseIf RetornarDatas = False Then
    'Only values
    result = Application.Index(result, seqDates, 2)
ElseIf RetornarDatas = True Then
    'Only dates
    colData = colData + 1
    result = Application.Index(result, seqDates, 1)
ElseIf RetornarDatas = 2 Then
    'Dates and values
    'No need to change (but array must start at 1)
    colData = colData + 1
    result = Application.Index(result, seqDates, Array(1, 2))
Else
    result = "# RetornarDatas é inválido"
    GoTo Final
End If

'Format values
If colData > 0 Then
    'Check whether it is a singleton
    If minDateIdx = maxDateIdx Then
        result(colData) = DateValue(result(colData))
    Else
        For iObs = 1 To UBound(result, 1)
            result(iObs, colData) = DateValue(result(iObs, colData))
        Next iObs
    End If
End If

Final:

xlFOCUS_MercadoImobiliario = result
    
End Function

Function xlFOCUS_WEBSERVICE(URL As String) As Variant

'This function circumvents the Excel's WEBSERVICE limitations, but still suffers from the String size limitation when passing strings directly to Excel

''''''''NEW METHOD
Dim xmlhttp As Object
Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
xmlhttp.Open "GET", URL, False
xmlhttp.Send

xlFOCUS_WEBSERVICE = xmlhttp.responseText

'Clear memory
Set xmlhttp = Nothing

End Function

Function xlFOCUS_ReadJSONFromWEBSERVICE(URL As String, Optional showHeaders As Boolean = False, Optional Fields As Variant, Optional subField As Variant) As Variant

'This function circumvents the String size limitation when passing strings directly to Excel

''''''''NEW METHOD
Dim xmlhttp As Object
Dim jsonScript As String

Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
xmlhttp.Open "GET", URL, False
xmlhttp.Send

jsonScript = xmlhttp.responseText

xlFOCUS_ReadJSONFromWEBSERVICE = xlFOCUS_ReadJSON(jsonScript, showHeaders, Fields, subField)

'Clear memory
Set xmlhttp = Nothing

End Function

Function xlFOCUS_IBGE_SIDRA(Optional Path As String, Optional Tabela As String, Optional Variavel As String, _
    Optional Classificacao As String, Optional NivelTerritorial As String, Optional CampoData As String, _
    Optional DataInicial As Variant = "", Optional DataFinal As Variant = "", _
    Optional nUltimos As Variant = "", Optional RetornarDatas As Variant = False) As Variant

' Details about the webservice:
' https://apisidra.ibge.gov.br/

Dim URL As String, URLMeta As String
Dim jsonScript As String, jsonScriptMeta As String
Dim pPath_str As String, Tabela_str As String, Variavel_str As String, NivelTerritorial_str As String, Classificacao_str As String
Dim Campos As Variant
Dim result As Variant, resultMeta As Variant
Dim frequenciaStr As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_IBGE_SIDRA = "# Barra de fórmulas aberta"
    Exit Function
End If

pPath_str = CStr(Path)
Tabela_str = CStr(Tabela)
Variavel_str = CStr(Variavel)
NivelTerritorial_str = CStr(NivelTerritorial)
Classificacao_str = CStr(Classificacao)

'Check parameters consistency
If Len(pPath_str) > 0 And _
    (Len(Variavel_str) > 0 _
    Or Len(NivelTerritorial_str) _
    Or Len(Classificacao_str) > 0) Then

    xlFOCUS_IBGE_SIDRA = "# Se Path é especificado, apenas Tabela precisa ser especificado"
    Exit Function

End If

'Check URL
'Return observations between dates
If Len(pPath_str) > 0 Then
    URL = "http://api.sidra.ibge.gov.br/values" _
            & pPath_str _
            & "/p/all" _
            & "/f/a" _
            & "/d/m" _
            & "/h/n" _
            & "?formato=json"
Else
    URL = "http://api.sidra.ibge.gov.br/values" _
            & IIf(Len(Tabela_str) = 0, "", "/t/" & Tabela_str) _
            & IIf(Len(Variavel_str) = 0, "", "/v/" & Variavel_str) _
            & IIf(Len(Classificacao_str) = 0, "", "/" & Classificacao_str) _
            & IIf(Len(NivelTerritorial_str) = 0, "", "/" & NivelTerritorial_str) _
            & "/p/all" _
            & "/f/a" _
            & "/d/m" _
            & "/h/n" _
            & "?formato=json"
End If

URLMeta = "https://servicodados.ibge.gov.br/api/v3/agregados/" & Tabela_str & "/metadados"

Campos = Array(CampoData, "V")

'Fetch webservice
jsonScript = xlFOCUS_WEBSERVICE(URL)
jsonScriptMeta = xlFOCUS_WEBSERVICE(URLMeta)

result = xlFOCUS_ReadJSON(jsonScript, False, Campos)
resultMeta = xlFOCUS_ReadJSON(jsonScriptMeta, False, , "periodicidade")
frequenciaStr = LCase(resultMeta(0, 1))


xlFOCUS_IBGE_SIDRA = sistema_xlFOCUS_IBGE_SIDRA(result, frequenciaStr, Campos, CampoData, _
    DataInicial, DataFinal, _
    nUltimos, RetornarDatas)

End Function

Function xlFOCUS_IBGE_Agregados(Optional Tabela As String, Optional Variavel As String, _
   Optional Classificacao As String, Optional lLocalidade As String, _
    Optional DataInicial As Variant = "", Optional DataFinal As Variant = "", _
    Optional nUltimos As Variant = "", Optional RetornarDatas As Variant = False) As Variant

' Details about the webservice:
' https://servicodados.ibge.gov.br/api/docs/agregados?versao=3

Dim URL As String, URLMeta As String
Dim jsonScript As String, jsonScriptMeta As String
Dim Tabela_str As String, Variavel_str As String, lLocalidade_str As String, Classificacao_str As String
Dim DataInicial_str As String, DataFinal_str As String, Periodos_str As String
Dim Campos As Variant
Dim result As Variant, resultMeta As Variant
Dim frequenciaStr As String

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    xlFOCUS_IBGE_Agregados = "# Barra de fórmulas aberta"
    Exit Function
End If

'Force range to be values
DataInicial = DataInicial
DataFinal = DataFinal
If IsEmpty(DataInicial) Then: DataInicial = ""
If IsEmpty(DataFinal) Then: DataFinal = ""

Tabela_str = CStr(Tabela)
Variavel_str = CStr(Variavel)
lLocalidade_str = CStr(lLocalidade)
Classificacao_str = CStr(Classificacao)

'Check URL Metadata
URLMeta = "https://servicodados.ibge.gov.br/api/v3/agregados/" & Tabela_str & "/metadados"

Campos = Array("Data", "V")

'Fetch metadata webservice
jsonScriptMeta = xlFOCUS_WEBSERVICE(URLMeta)
resultMeta = xlFOCUS_ReadJSON(jsonScriptMeta, False, , "periodicidade")
frequenciaStr = LCase(resultMeta(0, 1))

'Build periods expression
If Len(DataInicial) = 0 Or Len(DataFinal) = 0 Then
    Periodos_str = "all"
Else
    Select Case frequenciaStr
        Case "anual":
            DataInicial_str = Format(DataInicial, "yyyy")
            DataFinal_str = Format(DataFinal, "yyyy")
        Case "semestral":
            DataInicial_str = Format(DateSerial(Year(DataInicial), 1, 1), "yyyy") & Format(IIf(Month(DataInicial) < 7, 1, 2), "00")
            DataFinal_str = Format(DateSerial(Year(DataFinal), 1, 1), "yyyy") & Format(IIf(Month(DataFinal) < 7, 1, 2), "00")
        Case "trimestral":
            DataInicial_str = Format(DateSerial(Year(DataInicial), 1, 1), "yyyy") & Format(Application.WorksheetFunction.RoundUp(Month(DataInicial) / 3, 0), "00")
            DataFinal_str = Format(DateSerial(Year(DataFinal), 1, 1), "yyyy") & Format(Application.WorksheetFunction.RoundUp(Month(DataFinal) / 3, 0), "00")
        Case "mensal":
            DataInicial_str = Format(DataInicial, "yyyymm")
            DataFinal_str = Format(DataFinal, "yyyymm")
        Case "trimestral móvel":
            DataInicial_str = Format(DataInicial, "yyyymm")
            DataFinal_str = Format(DataFinal, "yyyymm")
    End Select
    
    'Check whether dates are the same
    Periodos_str = DataInicial_str & "-" & DataFinal_str
End If

'Check URL Data
URL = "https://servicodados.ibge.gov.br/api/v3/agregados" _
    & "/" & Tabela_str _
    & "/periodos/" & Periodos_str _
    & "/variaveis/" & Variavel_str _
    & "?" _
    & "localidades=" & lLocalidade_str _
    & "&classificacao=" & Classificacao_str

'Fetch data webservice
jsonScript = xlFOCUS_WEBSERVICE(URL)
result = xlFOCUS_ReadJSON(jsonScript, False, , Array(1, "resultados", 1, "series", 1, "serie"))

xlFOCUS_IBGE_Agregados = sistema_xlFOCUS_IBGE_SIDRA(result, frequenciaStr, Campos, "Data", _
    DataInicial, DataFinal, _
    nUltimos, RetornarDatas)

End Function


Private Function sistema_xlFOCUS_IBGE_SIDRA(result As Variant, frequenciaStr As String, Campos As Variant, Optional CampoData As String, _
    Optional DataInicial As Variant = "", Optional DataFinal As Variant = "", _
    Optional nUltimos As Variant = "", Optional RetornarDatas As Variant = False) As Variant

'Details about the webservice: https://apisidra.ibge.gov.br/
' https://servicodados.ibge.gov.br/api/docs/agregados?versao=3

Dim nUltimos_str As String, DataInicial_str As String, DataFinal_str As String
Dim colData As Long, colValor As Long
Dim iObs As Long
Dim datesVector() As Long, seqDates() As Variant
Dim minDateIdx As Long, maxDateIdx As Long
Dim dateAux As String
Dim allPeriods As Long
Dim nCols As Long
Dim listMatches As Variant
Dim colNIVNOME As Long, colTERCODIGO As Long
Dim oldDates As Boolean
Dim resultAux As Variant
Dim nDim As Long

' Avoid recalculation when the function wizard is being used
If (Not Application.CommandBars("Standard").Controls(1).Enabled) And recalculateWhenFunctionWizardIsOpen = False Then
    sistema_xlFOCUS_IBGE_SIDRA = "# Barra de fórmulas aberta"
    Exit Function
End If

DataInicial_str = Format(DataInicial, "dd/MM/yyyy")
DataFinal_str = Format(DataFinal, "dd/MM/yyyy")
nUltimos_str = CStr(nUltimos)

'If Not (LenB(DataInicial) = 0 And LenB(DataFinal) = 0) Then
'    If LenB(DataInicial) = 0 Then
'        result = "# Data inicial ausente"
'        GoTo Final
'    End If
'    If LenB(DataFinal) = 0 Then
'        result = "# Data final ausente"
'        GoTo Final
'    End If
'End If

'If LenB(DataInicial) <> 0 And LenB(DataFinal) <> 0 Then
'    URL = URL & "&dataInicial=" & DataInicial_str & "&dataFinal=" & DataFinal_str
'End If

'Check returned values
If VarType(result) = vbString Then
    GoTo Final
End If

'Fix for the case of a single observation
If UBound(result, 1) - LBound(result, 1) = 0 Then
    resultAux = result
    ReDim result(1 To 1, 1 To 2)
    result(1, 1) = resultAux(LBound(resultAux, 1), LBound(resultAux, 2))
    result(1, 2) = resultAux(LBound(resultAux, 1), UBound(resultAux, 2))
Else
    result = Application.Index(result, 0, 0)
End If

'Format values
colData = -1
On Error Resume Next
colData = colData + 1 + Application.WorksheetFunction.Match(CampoData, Campos, 0)
On Error GoTo 0
colValor = Application.WorksheetFunction.Match("V", Campos, 0)

If colData > -1 Then
    ReDim Preserve datesVector(LBound(result, 1) To UBound(result, 1))
    For iObs = LBound(result, 1) To UBound(result, 1)
        dateAux = Left$(result(iObs, colData), 10)
        
        'Format date
        Select Case frequenciaStr
            Case "trimestral"
                datesVector(iObs) = DateSerial(CLng(Left$(dateAux, 4)), (CLng(Right$(dateAux, 2)) - 1) * 3 + 1, 1)
            Case "anual"
                datesVector(iObs) = DateSerial(CLng(dateAux), 1, 1)
            Case Else
                datesVector(iObs) = CLng(DateValue(dateAux))
        End Select
        
    Next iObs
End If

'Slice dates
If Len(DataInicial_str) = 0 Then
    minDateIdx = 1
ElseIf datesVector(1) >= DateValue(DataInicial_str) Then
    minDateIdx = 1
Else
    minDateIdx = Application.WorksheetFunction.Match(CLng(DateValue(DataInicial_str)), datesVector, 1)
    If CLng(DateValue(DataInicial_str)) > datesVector(minDateIdx) Then
        minDateIdx = minDateIdx + 1
    End If
End If
If Len(DataFinal_str) = 0 Then
    maxDateIdx = UBound(datesVector)
ElseIf datesVector(UBound(datesVector)) <= CLng(DateValue(DataFinal_str)) Then
    maxDateIdx = UBound(datesVector)
Else
    maxDateIdx = Application.WorksheetFunction.Match(CLng(DateValue(DataFinal_str)), datesVector, 1)
End If
seqDates = Application.WorksheetFunction.Sequence(maxDateIdx - minDateIdx + 1, 1, minDateIdx)

'Define return table format
If IsMissing(RetornarDatas) Then
    'Only values
    nCols = 1
    colData = -1
    result = Application.Index(result, seqDates, colValor)
    colValor = 1
ElseIf RetornarDatas = False Then
    'Only values
    nCols = 1
    colData = -1
    result = Application.Index(result, seqDates, colValor)
    colValor = 1
ElseIf RetornarDatas = True Then
    'Only dates
    nCols = 1
    result = Application.Index(result, seqDates, colData)
    colValor = 1
ElseIf RetornarDatas = 2 Then
    'Dates and values
    nCols = 2
    'No need to change (but array must start at 1)
    
    'Check whether is a single observation
    If UBound(result, 1) - LBound(result, 1) <> 0 Then
        result = Application.Index(result, seqDates, Array(colData, colValor))
    End If
    colData = 1
    colValor = 2
Else
    result = "# RetornarDatas é inválido"
    GoTo Final
End If

'Slice periods
If Len(nUltimos_str) > 0 Then
    allPeriods = UBound(result, 1) - LBound(result, 1) + 1
    
    If CLng(nUltimos) > allPeriods Then nUltimos = allPeriods
    
    'Check whether is a single observation
    If UBound(result, 1) - LBound(result, 1) <> 0 Then
        result = Application.Index(result, _
            Application.WorksheetFunction.Sequence(CLng(nUltimos), 1, allPeriods - CLng(nUltimos) + 1, 1), _
            Application.WorksheetFunction.Sequence(1, nCols, 1, 1))
    End If
End If

'Get dimensions
nDim = getDimension(result)

'Format dates
oldDates = False
If colData > 0 Then
    For iObs = 1 To UBound(result, 1)
        If nDim = 2 Then
            dateAux = result(iObs, colData)
        Else
            dateAux = result(iObs)
        End If
        'Check whether dates are older than Excel's first date
        
        'Format date
        Select Case frequenciaStr
            Case "trimestral"
                dateAux = DateSerial(CLng(Left$(dateAux, 4)), (CLng(Right$(dateAux, 2)) - 1) * 3 + 1, 1)
            Case "anual"
                dateAux = DateSerial(CLng(dateAux), 1, 1)
            Case Else
                dateAux = CLng(DateValue(dateAux))
        End Select
        
        If oldDates = False And DateValue(dateAux) > DateValue("1900-01-01") Then
            dateAux = DateValue(dateAux)
        Else
            oldDates = True
        End If
        
        If nDim = 2 Then
            result(iObs, colData) = dateAux
        Else
            result(iObs) = dateAux
        End If
        
    Next iObs
End If

'Format values
If colValor > 0 And RetornarDatas <> True Then
    If nDim = 2 Then
        For iObs = 1 To UBound(result, 1)
            result(iObs, colValor) = CDbl(result(iObs, colValor))
        Next iObs
    Else
        For iObs = 1 To UBound(result, 1)
            result(iObs) = CDbl(result(iObs))
        Next iObs
    End If
End If


Final:

sistema_xlFOCUS_IBGE_SIDRA = result
    
End Function

Function xlFOCUS_INDEX(Values As Variant, Optional Row As Long, Optional Column As Long)

'Force range to be matrix
Values = Application.Transpose(Application.Transpose(Values))

If Row > 0 Then
    Row = Row
ElseIf Row < 0 Then
    Row = UBound(Values, 1) + Row + 1
End If

If Column > 0 Then
    Column = Column
ElseIf Row < 0 Then
    Column = UBound(Values, 2) + Column + 1
End If

xlFOCUS_INDEX = Application.WorksheetFunction.Index(Values, Row, Column)

End Function


''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''' OTHER FUNCTIONS'
''''''''''''''''''''''''''''''''''''''

Private Function getDimension(var As Variant) As Long
' https://stackoverflow.com/questions/6901991/how-to-return-the-number-of-dimensions-of-a-variant-variable-passed-to-it-in-v

On Error GoTo Err
Dim i As Long
Dim tmp As Long
i = 0
Do While True
    i = i + 1
    tmp = UBound(var, i)
Loop

Err:
getDimension = i - 1

End Function
