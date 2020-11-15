Attribute VB_Name = "modGeneral"
Option Explicit

' Client Executes Here.
Public Sub Main()
    frmSendGetData.Visible = True
    
    ' Check to make sure all the folder exist.
    Call SetStatus("Checking Folders...")
    Call CheckFolders

    ' Check to make sure all the files exist.
    Call SetStatus("Checking Files...")
    Call SystemFileChecker

    If Not FileExists("config.ini") Then
        Call FileCreateConfigINI
    End If

    If Not FileExists("Font.ini") Then
        Call FileCreateFontINI
    End If

    ' Load the configuration settings.
    Call SetStatus("Loading Configuration...")
    Call LoadConfig
    Call LoadColors
    Call LoadFont

    ' Prepare the socket for communication.
    Call SetStatus("Preparing Socket...")
    Call TcpInit

    frmMainMenu.lblVersion.Caption = "Version: 2.0"

    frmSendGetData.Visible = False
    frmMainMenu.Visible = True
End Sub

Private Sub CheckFolders()
    If LCase$(Dir$(App.Path & "\Maps", vbDirectory)) <> "maps" Then
        Call MkDir$(App.Path & "\Maps")
    End If

    If UCase$(Dir$(App.Path & "\GFX", vbDirectory)) <> "GFX" Then
        Call MkDir$(App.Path & "\GFX")
    End If

    If UCase$(Dir$(App.Path & "\GUI", vbDirectory)) <> "GUI" Then
        Call MkDir$(App.Path & "\GUI")
    End If

    If UCase$(Dir$(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Call MkDir$(App.Path & "\Music")
    End If

    If UCase$(Dir$(App.Path & "\SFX", vbDirectory)) <> "SFX" Then
        Call MkDir$(App.Path & "\SFX")
    End If

    If UCase$(Dir$(App.Path & "\DATA", vbDirectory)) <> "DATA" Then
        Call MkDir$(App.Path & "\Data")
    End If
End Sub

Private Sub LoadConfig()
    Dim FileName As String

    FileName = App.Path & "\config.ini"

    frmMirage.chkBubbleBar.Value = CInt(ReadINI("CONFIG", "SpeechBubbles", FileName))
    frmMirage.chkNpcBar.Value = CInt(ReadINI("CONFIG", "NpcBar", FileName))
    frmMirage.chkNpcName.Value = CInt(ReadINI("CONFIG", "NPCName", FileName))
    frmMirage.chkPlayerBar.Value = CInt(ReadINI("CONFIG", "PlayerBar", FileName))
    frmMirage.chkPlayerName.Value = CInt(ReadINI("CONFIG", "PlayerName", FileName))
    frmMirage.chkPlayerDamage.Value = CInt(ReadINI("CONFIG", "NPCDamage", FileName))
    frmMirage.chkNpcDamage.Value = CInt(ReadINI("CONFIG", "PlayerDamage", FileName))
    frmMirage.chkSound.Value = CInt(ReadINI("CONFIG", "Sound", FileName))
    frmMirage.chkAutoScroll.Value = CInt(ReadINI("CONFIG", "AutoScroll", FileName))
    AutoLogin = CInt(ReadINI("CONFIG", "Auto", FileName))
    
    If ReadINI("CONFIG", "AutoRun", FileName) = vbNullString Then
        Call WriteINI("CONFIG", "AutoRun", "0", FileName)
        
        frmMirage.chkAutoRun.Value = Unchecked
        Exit Sub
    End If
    
    frmMirage.chkAutoRun.Value = CInt(ReadINI("CONFIG", "AutoRun", FileName))
End Sub

Private Sub FileCreateConfigINI()
    WriteINI "CONFIG", "Account", vbNullString, App.Path & "\config.ini"
    WriteINI "CONFIG", "Password", vbNullString, App.Path & "\config.ini"
    WriteINI "CONFIG", "SpeechBubbles", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "NpcBar", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "NPCName", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "NPCDamage", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerBar", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerName", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerDamage", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "MapGrid", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "Music", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "Sound", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "AutoScroll", "1", App.Path & "\config.ini"
    WriteINI "CONFIG", "AutoRun", "0", App.Path & "\config.ini"
    WriteINI "CONFIG", "Auto", "0", App.Path & "\config.ini"
    WriteINI "CONFIG", "MenuMusic", "New Super Mario Bros. - Title Screen.mp3", App.Path & "\config.ini"
End Sub

Private Sub FileCreateFontINI()
    Call WriteINI("FONT", "Font", "fixedsys", App.Path & "\Font.ini")
    Call WriteINI("FONT", "Font2", "Comic Sans MS", App.Path & "\Font.ini")
    Call WriteINI("FONT", "Size", "18", App.Path & "\Font.ini")
End Sub

Private Sub LoadColors()
    ' Chat box color
    frmMirage.txtChat.BackColor = RGB(152, 146, 120)

    ' Chat box text color
    frmMirage.txtMyTextBox.BackColor = RGB(152, 146, 120)

    ' Special Attacks listbox
    frmMirage.lstSpells.BackColor = RGB(152, 146, 120)

    ' Online listbox
    frmMirage.lstOnline.BackColor = RGB(152, 146, 120)
    
    ' Item Description
    frmMirage.itmDesc.BackColor = RGB(152, 146, 120)
End Sub

Private Sub LoadFont()
    On Error GoTo ErrorHandle

    Font = ReadINI("FONT", "Font", App.Path & "\Font.ini")
    fontsize = CByte(ReadINI("FONT", "Size", App.Path & "\Font.ini"))
    Font2 = ReadINI("FONT", "Font2", App.Path & "\Font.ini")

    If Font = vbNullString Then
        Font = "fixedsys"
    End If

    If fontsize <= 0 Or fontsize > 32 Then
        fontsize = 18
    End If

    Call SetFont(GameFont, Font, fontsize)
    Call SetFont(GameFont2, Font2, 17)
    Call SetFont(GameFont3, "Comic Sans MS", 20)
    Call SetFont(GameFont4, "Comic Sans MS", 14)

    Exit Sub

ErrorHandle:
    Call WriteINI("FONT", "Font", "fixedsys", App.Path & "\Font.ini")
    Call WriteINI("FONT", "Size", 18, App.Path & "\Font.ini")
    Call WriteINI("FONT", "Font2", "Comic Sans MS", App.Path & "\Font.ini")

    Call SetFont(GameFont, "fixedsys", 18)
    Call SetFont(GameFont2, "Comic Sans MS", 17)
End Sub
