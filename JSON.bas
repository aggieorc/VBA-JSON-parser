Option Explicit

Sub Scrape_picknpay_co_za()

    Dim sResponse As String
    Dim sState As String
    Dim vJSON As Variant
    Dim aRows() As Variant
    Dim aHeader() As Variant

    ' Retrieve JSON data
    XmlHttpRequest "POST", "https://app.estimateone.com/awarded/fetchMapData?daysToShow=60", "", "", "", sResponse
    ' Parse JSON response
    JSON.Parse sResponse, vJSON, sState
    If sState <> "Array" Then
        MsgBox "Invalid JSON response"
        Exit Sub
    End If
    ' Convert result to arrays for output
    JSON.ToArray vJSON, aRows, aHeader
    ' Output
    With ThisWorkbook.Sheets(1)
        OutputArray .Cells(1, 1), aHeader
        Output2DArray .Cells(2, 1), aRows
        .Columns.AutoFit
    End With

    MsgBox "Completed"

End Sub

Sub XmlHttpRequest(sMethod As String, sUrl As String, arrSetHeaders, sFormData, sRespHeaders As String, sContent As String)

    Dim arrHeader

    'With CreateObject("Msxml2.ServerXMLHTTP")
    '    .SetOption 2, 13056 ' SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    With CreateObject("MSXML2.XMLHTTP")
        .Open sMethod, sUrl, False
        If IsArray(arrSetHeaders) Then
            For Each arrHeader In arrSetHeaders
                .SetRequestHeader arrHeader(0), arrHeader(1)
            Next
        End If
        .send sFormData
        sRespHeaders = .GetAllResponseHeaders
        sContent = .responseText
    End With

End Sub

Sub OutputArray(oDstRng As Range, aCells As Variant)

    With oDstRng
        .Parent.Select
        With .Resize(1, UBound(aCells) - LBound(aCells) + 1)
            .NumberFormat = "@"
            .Value = aCells
        End With
    End With

End Sub

Sub Output2DArray(oDstRng As Range, aCells As Variant)

    With oDstRng
        .Parent.Select
        With .Resize( _
                UBound(aCells, 1) - LBound(aCells, 1) + 1, _
                UBound(aCells, 2) - LBound(aCells, 2) + 1)
            .NumberFormat = "@"
            .Value = aCells
        End With
    End With

End Sub
