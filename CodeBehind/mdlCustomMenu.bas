Attribute VB_Name = "mdlCustomMenu"

Public Const cCustomMenuName = "&Custom Menu"
Public Const cWsMenuBar = "Worksheet Menu Bar"

Public Sub LoadCustomMenus()
    Dim cmbBar As CommandBar
    Dim cmbControl As CommandBarControl
    Dim cmbSettings As CommandBarControl
    Dim cmbDBLink As CommandBarControl
    
    'add custom menus to Add-In ribon of Excel
    Set cmbBar = Application.CommandBars(cWsMenuBar)
    
    'check if the custom menu already exists. If it exists, delete it; it will be recreated in the later code
    Dim i As Integer ', boolMenuExists As Boolean
    For i = cmbBar.Controls.Count To 1 Step -1
        If cmbBar.Controls.Item(i).Caption = cCustomMenuName Then
            'boolMenuExists = True
            cmbBar.Controls.Item(i).Delete
            Exit For
        End If
    Next
    
    'create menu bar entries
    Set cmbControl = cmbBar.Controls.Add(Type:=msoControlPopup, Temporary:=True) 'adds a menu item to the Menu Bar
    With cmbControl
        .Caption = cCustomMenuName 'names the menu item
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Import Detection Manifest" 'adds a description to the menu item
            .OnAction = "ImportDetectionFile" 'runs the specified macro
            .FaceId = 109 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Export Validated Detection File" 'adds a description to the menu item
            .OnAction = "SavePreparedData" 'runs the specified macro
            .FaceId = 526 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Refresh Validation Results" 'adds a description to the menu item
            .OnAction = "RefreshWorkbookData" 'runs the specified macro
            .FaceId = 37 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Refresh Database Links" 'adds a description to the menu item
            .OnAction = "RefreshDBConnections" 'runs the specified macro
            .FaceId = 688 '638 '1098 'assigns an icon to the dropdown
        End With
'        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
'            .Caption = "Help" 'adds a description to the menu item
'            .OnAction = "OpenHelpLink" 'runs the specified macro
'            .FaceId = 49 '638 '1098 'assigns an icon to the dropdown
'        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "About" 'adds a description to the menu item
            .OnAction = "ShowVersionMsg" 'runs the specified macro
            .FaceId = 279 '501 '638 '1098 'assigns an icon to the dropdown
        End With
        
    End With
End Sub



