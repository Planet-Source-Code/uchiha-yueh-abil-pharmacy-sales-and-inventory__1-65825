Attribute VB_Name = "Modular"
Public UserType As String
Public User As String
Public UPass As String
Public ActiveReport As String

Sub Menus(ByVal UType As String)

With mdiMain

    'enable menus in mdiMain mdiform..
    
    .mnuPOS.Enabled = Not .mnuPOS.Enabled
    .mnuProd.Enabled = Not .mnuProd.Enabled
    .mnuExpired.Enabled = Not .mnuExpired.Enabled
    .mnuLocate.Enabled = Not .mnuLocate.Enabled
    
    .mnuInvRept.Enabled = Not .mnuInvRept.Enabled
    
    .mnuChange.Enabled = Not .mnuChange.Enabled

    'enable toolbar..
    .Toolbar1.Buttons(1).Enabled = _
        Not .Toolbar1.Buttons(1).Enabled
    .Toolbar1.Buttons(2).Enabled = _
        Not .Toolbar1.Buttons(2).Enabled
    .Toolbar1.Buttons(3).Enabled = _
        Not .Toolbar1.Buttons(3).Enabled
    .Toolbar1.Buttons(4).Enabled = _
        Not .Toolbar1.Buttons(4).Enabled
    
    'check usertype..
    If UType = "Administrator" Then
    
        'enable User account menu..
        .mnuUser.Enabled = True
        
        'enable report..
        .mnuInvRept.Enabled = True
        
        .Toolbar1.Buttons(3).Enabled = True
    
    'if UType is user..
    ElseIf UType = "User" Then
    
        'disable mnuUser..
        .mnuUser.Enabled = False
        
        'enable report..
        .mnuInvRept.Enabled = False
    
        .Toolbar1.Buttons(3).Enabled = False
    
    End If

    'check mnuLog caption..
    If .mnuLog.Caption = "Login" Then
        
        .mnuLog.Caption = "Log-out"
    
    Else
        
        'show toolbar..
        mdiMain.Toolbar1.Visible = True

        .mnuLog.Caption = "Login"
    End If
    
End With

End Sub
