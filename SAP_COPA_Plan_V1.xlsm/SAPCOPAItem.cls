VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCOPAItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public gFIELDNAME As String
Public gVALUE As Variant
Public gCURRENCY As String
Public gSELECTION As String

Public Function create(pFIELDNAME As String, pVALUE As Variant, pCURRENCY As String, Optional ByVal pSELECTION = "")
    gFIELDNAME = pFIELDNAME
    gVALUE = pVALUE
    gCURRENCY = pCURRENCY
    gSELECTION = pSELECTION
End Function


