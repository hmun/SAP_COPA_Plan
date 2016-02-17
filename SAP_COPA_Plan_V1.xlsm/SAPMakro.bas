Attribute VB_Name = "SAPMakro"
Sub SAP_COPA_Plan()
    Dim SAPCOPAPlanning As New SAPCOPAPlanning
    Dim aSAPCOPAItem As New SAPCOPAItem
    Dim aSAPFormat As New SAPFormat
    Dim aSAPProjectDefinition As New SAPProjectDefinition
    Dim aSAPWbsElement As New SAPWbsElement
    Dim aData As New Collection
    Dim aDataRow As New Collection
    Dim aLines As Integer
    Dim aStartLine As Integer
    Dim aEndLine As Integer
    Dim aLineCnt As Integer

    Dim i As Integer
    Dim j As Integer
    Dim maxJ As Integer
    Dim aRetStr As String

    Dim aFIELDNAME As String
    Dim aVALUE As Variant
    Dim aCURRENCY As String

    Dim aOperatingConcern As String
    Dim aTypeOfProfitAnalysis As String
    Dim aTestRun As String

    Worksheets("Parameter").Activate
    aOperatingConcern = Cells(2, 2).Value
    aLines = Cells(3, 2).Value
    aTypeOfProfitAnalysis = Cells(4, 2).Value
    aTestRun = Cells(5, 2).Value
    If IsNull(aOperatingConcern) Or aOperatingConcern = "" Then
        MsgBox "Bitte die Pflichtfelder im Blatt Parameter füllen!", vbCritical + vbOKOnly
        Exit Sub
    End If
    If IsNull(aLines) Or aLines = 0 Then
        aLines = 1
    End If

    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Verbindung zu SAP fehlgeschlagen!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Items
    Worksheets("Data").Activate
    i = 6
    ' determine the last column
    maxJ = 1
    Do
        maxJ = maxJ + 1
    Loop While Not IsNull(Cells(1, maxJ)) And Cells(1, maxJ) <> ""

    aStartLine = i
    aLineCnt = 0
    Set aData = New Collection
    Do
        If Left(Cells(i, maxJ), 7) <> "Success" Then
            Set aDataRow = New Collection
            j = 1
            Do
                Set aSAPCOPAItem = New SAPCOPAItem
                If Not IsNull(Cells(2, j).Value) And Cells(2, j).Value <> "" Then
                    aCURRENCY = Cells(2, j).Value
                Else
                    aCURRENCY = ""
                End If
                Select Case Cells(3, j).Value
                    Case "DATE"
                        aVALUE = Format$(CDate(Cells(i, j).Value), "YYYYMMDD")
                    Case "PERIO"
                        aVALUE = Right(Cells(i, j).Value, 4) & Left(Cells(i, j).Value, 3)
                    Case "PROJ"
                        If Cells(i, j).Value <> "" Then
                            aVALUE = aSAPProjectDefinition.GetPspnr(Cells(i, j).Value)
                        Else
                            aVALUE = ""
                        End If
                    Case "WBS"
                        If Cells(i, j).Value <> "" Then
                            aVALUE = aSAPWbsElement.GetPspnr(Cells(i, j).Value)
                        Else
                            aVALUE = ""
                        End If
                    Case Else
                        If Left(Cells(3, j).Value, 1) = "U" Then
                            aVALUE = aSAPFormat.unpack(Cells(i, j).Value, CInt(Right(Cells(3, j).Value, Len(Cells(3, j).Value) - 1)))
                        ElseIf Left(Cells(3, j).Value, 1) = "P" Then
                            aVALUE = aSAPFormat.pspid(Cells(i, j).Value, CInt(Right(Cells(3, j).Value, Len(Cells(3, j).Value) - 1)))
                        Else
                            aVALUE = Cells(i, j).Value
                        End If
                End Select
                aFIELDNAME = Cells(1, j).Value
                aSAPCOPAItem.create aFIELDNAME, aVALUE, aCURRENCY, Cells(4, j).Value
                aDataRow.Add aSAPCOPAItem
                j = j + 1
        Loop While Not IsNull(Cells(1, j)) And Cells(1, j) <> ""
        aData.Add aDataRow
        aLineCnt = aLineCnt + 1
        If aLineCnt >= aLines Then
            aEndLine = i
            '     post the lines
            Application.StatusBar = "Posting at line " & aEndLine
            aRetStr = SAPCOPAPlanning.PostData(aOperatingConcern, aTypeOfProfitAnalysis, aTestRun, aData)
            For Each c In Range(Cells(aStartLine, j), Cells(aEndLine, j))
                If Left(c.Value, 7) <> "Success" Then
                    c.Value = aRetStr
                End If
            Next c
            aStartLine = i + 1
            aLineCnt = 0
            Set aData = New Collection
        End If
    Else
        Cells(i, maxJ + 1) = "ignored - already posted"
    End If
    i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
    ' post the rest
    If aData.Count > 0 Then
        aEndLine = i - 1
        Application.StatusBar = "Posting at line " & aEndLine
        aRetStr = SAPCOPAPlanning.PostData(aOperatingConcern, aTypeOfProfitAnalysis, aTestRun, aData)
        For Each c In Range(Cells(aStartLine, j), Cells(aEndLine, j))
            If Left(c.Value, 7) <> "Success" Then
                c.Value = aRetStr
            End If
        Next c
    End If

    Application.Cursor = xlDefault
End Sub
