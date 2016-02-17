VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCOPAPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function PostData(pOperatingConcern As String, pTypeOfProfitAnalysis As String, pTestRun As String, pData As Collection) As String
    On Error GoTo PostData_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCOPAPlanning As Object
    Dim lOperatingConcern
    Dim oSelectedFields As Object
    Dim oSelection As Object
    Dim oPervalue As Object
    Dim oRETURN As Object
    Set boCOPAPlanning = MySAPCon.SAPBapiControl.GetSAPObject("COPAPlanning")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCOPAPlanning, "PostData", "Return")
    Set oSelectedFields = MySAPCon.SAPBapiControl.DimAs(boCOPAPlanning, "PostData", "SelectedFields")
    Set oSelection = MySAPCon.SAPBapiControl.DimAs(boCOPAPlanning, "PostData", "Selection")
    Set oData = MySAPCon.SAPBapiControl.DimAs(boCOPAPlanning, "PostData", "Data")
    lOperatingConcern = pOperatingConcern

    Dim aSFRow As Object
    Dim aSRow As Object
    Dim aDRow As Object
    Dim aRow As Object
    Dim aItem As Object
    lCnt = 0
    For Each aRow In pData
        lCnt = lCnt + 1
        For Each aItem In aRow
            If lCnt = 1 Then
                Set aSFRow = oSelectedFields.Rows.Add
                aSFRow.Value("FIELDNAME") = aItem.gFIELDNAME

                If aItem.gSELECTION = "X" Then
                    Set aSRow = oSelection.Rows.Add
                    aSRow.Value("FIELDNAME") = aItem.gFIELDNAME
                    aSRow.Value("SIGN") = "I"
                    aSRow.Value("OPTION") = "EQ"
                    aSRow.Value("LOW") = aItem.gVALUE
                End If
            End If

            Set aDRow = oData.Rows.Add
            aDRow.Value("RECORD_ID") = lCnt
            aDRow.Value("FIELDNAME") = aItem.gFIELDNAME
            aDRow.Value("VALUE") = aItem.gVALUE
            If Not IsNull(aItem.gCURRENCY) And aItem.gCURRENCY <> "" Then
                aDRow.Value("CURRENCY") = aItem.gCURRENCY
            End If
        Next aItem
    Next aRow
    Application.Cursor = xlWait
    boCOPAPlanning.PostData OperatingConcern:=lOperatingConcern, _
    TypeOfProfitAnalysis:=pTypeOfProfitAnalysis, _
    TestRun:=pTestRun, _
    SelectedFields:=oSelectedFields, _
    Selection:=oSelection, _
    Data:=oData, _
    Return:=oRETURN
    Application.Cursor = xlDefault
    If oRETURN.RowCount = 0 Then
        PostData = "Success"
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        aSAPBapiTranctionCommit.commit
    Else
        Dim RetRow As Object
        For Each RetRow In oRETURN.Rows
            PostData = PostData & ";" & RetRow("MESSAGE")
        Next RetRow
    End If
    Set boCOPAPlanning = Nothing
    Exit Function
PostData_Error:
    Application.Cursor = xlDefault
    PostData = "Error"
    MySAPErr.MSGProt "SAPCOPAPlanning", "PostData", "", Err.Number, Err.Description
    Exit Function
End Function


