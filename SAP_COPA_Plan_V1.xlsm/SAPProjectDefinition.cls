VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPProjectDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Private Sub Class_Initialize()
    IsConnected = SAPCheck()
End Sub


Public Function GetPspnr(pPSPID As String) As String
    On Error GoTo GetPspnr_Error
    Dim oFunction As Object

    Set oFunction = MySAPCon.SAPFunctions.Add("Z_CO_PS_PROJ_INTERNAL")

    oFunction.Exports("I_PSPID") = pPSPID
    Ret = oFunction.call
    If Ret Then
        GetPspnr = oFunction.Imports("E_PSPNR")
    End If
    MySAPCon.SAPFunctions.Remove ("Z_CO_PS_PROJ_INTERNAL")
    Exit Function

GetPspnr_Error:
    GetPspnr = "Fehler"
    MySAPErr.MSGProt "SAPProjectDefinition", "GetPspnr", "", Err.Number, Err.Description
    Exit Function
End Function



