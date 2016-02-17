Attribute VB_Name = "SAPMenu"
Function delSAPCommandbar()
    Dim aCmdBars As CommandBars
    Dim aCmdBar As CommandBar
    Dim aCmdBarExists As Boolean
    Set aCmdBars = Application.CommandBars
    For Each aCmdBar In aCmdBars
      If aCmdBar.Name = "SAPCOPAPlanning" Then
        aCmdBarExists = True
        Exit For
      End If
    Next
    If aCmdBarExists Then
      aCmdBar.Delete
    End If
End Function

Function addSAPCommandbar()
Attribute addSAPCommandbar.VB_Description = "Makro am 8/12/2008 von Hermann Mundprecht aufgezeichnet"
Attribute addSAPCommandbar.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim aCmdBars As CommandBars
    Dim aCmdBar As CommandBar
    Dim aCmdBarExists As Boolean
    Set aCmdBars = Application.CommandBars
    For Each aCmdBar In aCmdBars
      If aCmdBar.Name = "SAPCOPAPlanning" Then
        aCmdBarExists = True
        Exit For
      End If
    Next
    If aCmdBarExists Then
      aCmdBar.Visible = True
    Else
      Set aCmdBar = aCmdBars.Add("SAPCOPAPlanning", msoBarTop, , True)
        Dim aButton As CommandBarControl
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .Style = msoButtonCaption
            .Caption = "COPA Planung"
            .TooltipText = "SAP-COPA Plan-Einzelposten verbuchen"
            .OnAction = "SAP_COPA_Plan"
        End With
        Set aButton = aCmdBar.Controls.Add(msoControlButton)
        With aButton
            .BeginGroup = True
            .Style = msoButtonCaption
            .Caption = "Logoff"
            .TooltipText = "Logoff from SAP"
            .OnAction = "SAPLogoff"
        End With
        aCmdBar.Visible = True
    End If
End Function
