Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal Index As Long, ByVal Data As String)
    Dim Parse() As String
    Dim i As Long

    Parse = Split(Data, SEP_CHAR)

    ' Hackers would send parse(0) as a String
    On Error GoTo Hacking

    Select Case CInt(Parse(0))
        Case CPackets.Cgetclasses
            Call Packet_GetClasses(Index)
            Exit Sub
    
        Case CPackets.Cnewaccount
            If UBound(Parse) = 3 Then
                Call Packet_NewAccount(Index, Parse(1), Parse(2), CBool(Parse(3)))
            Else
                Call Packet_NewAccount(Index, Parse(1), Parse(2))
            End If
            
            Exit Sub
            
        Case CPackets.Cdelaccount
            If UBound(Parse) = 3 Then
                Call Packet_DeleteAccount(Index, Parse(1), Parse(2), CBool(Parse(3)))
            Else
                Call Packet_DeleteAccount(Index, Parse(1), Parse(2))
            End If
            
            Exit Sub
    
        Case CPackets.Cacclogin
            Call Packet_AccountLogin(Index, Parse(1), Parse(2), CLng(Parse(3)), CLng(Parse(4)), CLng(Parse(5)), Parse(6), CBool(Parse(7)))
            Exit Sub
    
        Case CPackets.Cgivemethemax
            Call Packet_GiveMeTheMax(Index)
            Exit Sub
    
        Case CPackets.Caddchar
            Call Packet_AddCharacter(Index, Parse(1), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub
    
        Case CPackets.Cdelchar
            Call Packet_DeleteCharacter(Index, CLng(Parse(1)))
            Exit Sub
    
        Case CPackets.Cusechar
            Call Packet_UseCharacter(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cguildchangeaccess
            Call Packet_GuildChangeAccess(Index, Parse(1), CLng(Parse(2)))
            Exit Sub

        Case CPackets.Cguilddisown
            Call Packet_GuildDisown(Index, Parse(1))
            Exit Sub

        Case CPackets.Cguildleave
            Call Packet_GuildLeave(Index)
            Exit Sub

        Case CPackets.Cguildmake
            Call Packet_GuildMake(Index, Parse(1), Parse(2))
            Exit Sub

        Case CPackets.Cguildmember
            Call Packet_GuildMember(Index, Parse(1))
            Exit Sub
            
        Case CPackets.Cguildmemberrequest
            Call Packet_GuildMemberRequest(Index, Parse(1), CInt(Parse(2)))
            Exit Sub
            
        Case CPackets.Cguildmemberdecline
            Call Packet_GuildMemberDecline(Index, Parse(1))
            Exit Sub
        
        Case CPackets.Cguildtrainee
            Call Packet_GuildTrainee(Index, Parse(1))
            Exit Sub
        
        Case CPackets.Cgroupmemberlist
            Call Packet_GroupMemberList(Index)
            Exit Sub

        Case CPackets.Csaymsg
            Call Packet_SayMessage(Index, Parse(1))
            Exit Sub

        Case CPackets.Cgroupmsg
            Call Packet_GroupMessage(Index, Parse(1))
            Exit Sub

        Case CPackets.Cbroadcastmsg
            Call Packet_BroadcastMessage(Index, Parse(1))
            Exit Sub

        Case CPackets.Cglobalmsg
            Call Packet_GlobalMessage(Index, Parse(1))
            Exit Sub

        Case CPackets.Cadminmsg
            Call Packet_AdminMessage(Index, Parse(1))
            Exit Sub

        Case CPackets.Cplayermsg
            Call Packet_PlayerMessage(Index, Parse(1), Parse(2))
            Exit Sub
            
        Case CPackets.Cothermsg
            Call Packet_OtherMessage(Index, Parse(1), Parse(2))
            Exit Sub
            
        Case CPackets.Cplayermove
            Call Packet_PlayerMove(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub

        Case CPackets.Cplayerdir
            Call Packet_PlayerDirection(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cuseitem
            Call Packet_UseItem(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cplayermovemouse
            Call Packet_PlayerMoveMouse(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cwarp
            Call Packet_Warp(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cattack
            Call Packet_Attack(Index)
            Exit Sub

        Case CPackets.Cusestatpoint
            Call Packet_UseStatPoint(Index, CInt(Parse(1)))
            Exit Sub

        Case CPackets.Csetplayersprite
            Call Packet_SetPlayerSprite(Index, Parse(1), CLng(Parse(2)))
            Exit Sub

        Case CPackets.Cgetstats
            Call Packet_GetStats(Index, Parse(1))
            Exit Sub

        Case CPackets.Crequestnewmap
            Call Packet_RequestNewMap(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cwarpmeto
            Call Packet_WarpMeTo(Index, Parse(1))
            Exit Sub

        Case CPackets.Cwarptome
            Call Packet_WarpToMe(Index, Parse(1))
            Exit Sub

        Case CPackets.Cmapdata
            Call Packet_MapData(Index, Parse)
            Exit Sub

        Case CPackets.Cneedmap
            Call Packet_NeedMap(Index, Parse(1))
            Exit Sub

        Case CPackets.Cmapgetitem
            Call Packet_MapGetItem(Index)
            Exit Sub
            
        Case CPackets.Cmapdropitem
            Call Packet_MapDropItem(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub

        Case CPackets.Cmaprespawn
            Call Packet_MapRespawn(Index)
            Exit Sub

        Case CPackets.Ckickplayer
            Call Packet_KickPlayer(Index, Parse(1))
            Exit Sub

        Case CPackets.Cmuteplayer
            Call Packet_MutePlayer(Index, Parse(1))
            Exit Sub
        
        Case CPackets.Cunmuteplayer
            Call Packet_UnmutePlayer(Index, Parse(1))
            Exit Sub
        
        Case CPackets.Cgetmutelist
            Call Packet_GetMuteList(Index)
            Exit Sub
        
        Case CPackets.Cgetbanlist
            Call Packet_GetBanList(Index)
            Exit Sub
        
        Case CPackets.Cbanplayer
            If UBound(Parse) = 2 Then
                Call Packet_BanPlayer(Index, Parse(1), CBool(Parse(2)))
            Else
                Call Packet_BanPlayer(Index, Parse(1))
            End If
            
            Exit Sub
            
        Case CPackets.Cunbanplayer
            If UBound(Parse) = 2 Then
                Call Packet_UnbanPlayer(Index, CInt(Parse(1)), CBool(Parse(2)))
            Else
                Call Packet_UnbanPlayer(Index, CInt(Parse(1)))
            End If
            
            Exit Sub

        Case CPackets.Crequesteditmap
            Call Packet_RequestEditMap(Index)
            Exit Sub

        Case CPackets.Crequestedititem
            Call Packet_RequestEditItem(Index)
            Exit Sub

        Case CPackets.Cedititem
            Call Packet_EditItem(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Csaveitem
            Call Packet_SaveItem(Index, Parse)
            Exit Sub

        Case CPackets.Crequesteditnpc
            Call Packet_RequestEditNPC(Index)
            Exit Sub

        Case CPackets.Ceditnpc
            Call Packet_EditNPC(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Csavenpc
            Call Packet_SaveNPC(Index, Parse)
            Exit Sub

        Case CPackets.Crequesteditshop
            Call Packet_RequestEditShop(Index)
            Exit Sub

        Case CPackets.Ceditshop
            Call Packet_EditShop(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Csaveshop
            Call Packet_SaveShop(Index, Parse)
            Exit Sub

        Case CPackets.Crequesteditspell
            Call Packet_RequestEditSpell(Index)
            Exit Sub

        Case CPackets.Ceditspell
            Call Packet_EditSpell(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Csavespell
            Call Packet_SaveSpell(Index, Parse)
            Exit Sub

        Case CPackets.Cforgetspell
            Call Packet_ForgetSpell(Index, CLng(Parse(1)))
            Exit Sub
        
        Case CPackets.Cspecialattackdetails
            Call SendSpecialAttackInfo(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Csetaccess
            Call Packet_SetAccess(Index, Parse(1), CByte(Parse(2)))
            Exit Sub

        Case CPackets.Conlinelist
            Call Packet_OnlineList(Index)
            Exit Sub

        Case CPackets.Csetmotd
            Call Packet_SetMOTD(Index, Parse(1))
            Exit Sub

        Case CPackets.Cbuy
            Call Packet_BuyItem(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)))
            Exit Sub

        Case CPackets.Csellitem
            Call Packet_SellItem(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub
        
        Case CPackets.Csearch
            Call Packet_Search(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub

        Case CPackets.Cplayerchat
            Call Packet_PlayerChat(Index, Parse(1))
            Exit Sub

        Case CPackets.Cachat
            Call Packet_AcceptChat(Index)
            Exit Sub

        Case CPackets.Cdchat
            Call Packet_DenyChat(Index)
            Exit Sub

        Case CPackets.Cqchat
            Call Packet_QuitChat(Index)
            Exit Sub

        Case CPackets.Csendchat
            Call Packet_SendChat(Index, Parse(1))
            Exit Sub

        Case CPackets.Ctraderequest
            Call Packet_TradeRequest(Index, Parse(1))
            Exit Sub
            
        Case CPackets.Caccepttrade
            Call Packet_AcceptTrade(Index)
            Exit Sub
            
        Case CPackets.Cdeclinetrade
            Call Packet_DeclineTrade(Index)
            Exit Sub
        
        Case CPackets.Cstoptrading
            Call Packet_StopTrading(Index)
            Exit Sub

        Case CPackets.Cupdatetradeoffers
            Call Packet_UpdateTradeOffers(Index, CLng(Parse(1)), Parse(2), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub
            
        Case CPackets.Ccompletetrade
            Call Packet_CompleteTrade(Index)
            Exit Sub

        Case CPackets.Cparty
            Call Packet_PartyRequest(Index, Parse(1))
            Exit Sub

        Case CPackets.Cjoinparty
            Call Packet_JoinParty(Index)
            Exit Sub
        
        Case CPackets.Cpartydecline
            Call Packet_PartyDecline(Index)
            Exit Sub
            
        Case CPackets.Cleaveparty
            Call Packet_LeaveParty(Index)
            Exit Sub

        Case CPackets.Cspells
            Call Packet_Spells(Index)
            Exit Sub

        Case CPackets.Chotscript
            Call Packet_HotScript(Index, CByte(Parse(1)))
            Exit Sub

        Case CPackets.Ccast
            Call Packet_Cast(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cprompt
            Call Packet_Prompt(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub

        Case CPackets.Cquerybox
            Call Packet_QueryBox(Index, Parse(1), CLng(Parse(2)))
            Exit Sub

        Case CPackets.Crequesteditarrow
            Call Packet_RequestEditArrow(Index)
            Exit Sub

        Case CPackets.Ceditarrow
            Call Packet_EditArrow(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Csavearrow
            Call Packet_SaveArrow(Index, CLng(Parse(1)), Parse(2), CLng(Parse(3)), CLng(Parse(4)), CLng(Parse(5)))
            Exit Sub

        Case CPackets.Crequesteditemoticon
            Call Packet_RequestEditEmoticon(Index)
            Exit Sub

        Case CPackets.Crequesteditelement
            Call Packet_RequestEditElement(Index)
            Exit Sub

        Case CPackets.Ceditemoticon
            Call Packet_EditEmoticon(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Ceditelement
            Call Packet_EditElement(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Csaveemoticon
            Call Packet_SaveEmoticon(Index, CLng(Parse(1)), Parse(2), CLng(Parse(3)))
            Exit Sub

        Case CPackets.Csaveelement
            Call Packet_SaveElement(Index, CLng(Parse(1)), Parse(2), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub

        Case CPackets.Ccheckemoticons
            Call Packet_CheckEmoticon(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cmapreport
            Call Packet_MapReport(Index)
            Exit Sub

        Case CPackets.Cweather
            Call Packet_Weather(Index, CLng(Parse(1)))
            Exit Sub

        Case CPackets.Cwarpto
            Call Packet_WarpTo(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)))
            Exit Sub

        Case CPackets.Clocalwarp
            Call Packet_LocalWarp(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
        
        Case CPackets.Cmakeadmin
            Call Packet_MakeAdmin(Index)
            Exit Sub
        
        Case CPackets.Carrowhit
            Call Packet_ArrowHit(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub
        
        Case CPackets.Carrowswitch
            Call Packet_ArrowSwitch(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
        
        Case CPackets.Cbankdeposit
            Call Packet_BankDeposit(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub

        Case CPackets.Cbankwithdraw
            Call Packet_BankWithdraw(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
        
        Case CPackets.Cbankdestroy
            Call Packet_BankDestroy(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
            
        Case CPackets.Cgetonline
            Call Packet_GetWhosOnline(Index)
            Exit Sub
            
        Case CPackets.Cdragdropinv
            Call DragDropInv(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
            
        Case CPackets.Crunfrombattle
            Call LeaveBattle(Index)
            Exit Sub
            
        Case CPackets.Chasturnbased
            Call SetHasTurnBased(Index, CInt(Parse(1)))
            Exit Sub
            
        Case CPackets.Csetbattleturn
            Call SetBattleTurn(Index, CInt(Parse(1)))
            Exit Sub
            
        Case CPackets.Cstartbattle
            Call StartBattle(Index, CLng(Parse(1)))
            Exit Sub
        
        Case CPackets.Cuseturnbaseditem
            Call Packet_UseTurnBasedItem(Index, CLng(Parse(1)))
            Exit Sub
        
        Case CPackets.Crequesteditrecipe
            Call Packet_RequestEditRecipe(Index)
            Exit Sub
        
        Case CPackets.Ceditrecipe
            Call Packet_EditRecipe(Index, CLng(Parse(1)))
            Exit Sub
            
        Case CPackets.Csaverecipe
            Call Packet_SaveRecipe(Index, Parse)
            Exit Sub
        
        Case CPackets.Ccookitem
            Call CookItem(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)))
            Exit Sub
        
        Case CPackets.Ccooking
            Call FinishCooking(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
        
        Case CPackets.Cfriendonlineoffline
            i = FindPlayer(Parse(1))
            
            Call SendDataTo(Index, SPackets.Sfriendstatus & SEP_CHAR & i & SEP_CHAR & Parse(2) & END_CHAR)
            Exit Sub
        
        Case CPackets.Ccaption
            i = CLng(Parse(2))
            
            Call PutVar(App.Path & "\SMBOAccounts\" & "Friend Lists.ini", GetPlayerName(Index), (i + 1), Parse(1))
            Exit Sub
           
        Case CPackets.Cwarn
            Call Packet_WarnPlayer(Index, Parse(1), Parse(2))
            Exit Sub
            
        Case CPackets.Cremovewarn
            Call Packet_RemoveWarn(Index, Parse(1))
            Exit Sub
        
        Case CPackets.Ccardshop
            Call UpdateCardShop(Index, CLng(Parse(1)))
            Exit Sub
            
        Case CPackets.Cunequip
            Call UnequipItem(Index, CLng(Parse(1)))
            Exit Sub
        
        Case CPackets.Cuseturnbasedspecial
            Call UseTurnBasedSpecial(Index, CLng(Parse(1)))
            Exit Sub
        
        Case CPackets.Cfinishplayerbattle
            Call EndBattle(Index, GetPlayerMap(Index), CLng(Parse(1)), GetPlayerOldX(Index), GetPlayerOldY(Index))
            Exit Sub
        
        Case CPackets.Chelp
            Call HelpCommands(Index)
            Exit Sub
        
        Case CPackets.Cplayerheight
            Player(Index).Char(Player(Index).CharNum).Height = CInt(Parse(1))
            Exit Sub
            
        Case CPackets.Cjumping
            Call Packet_Jumping(Index, CByte(Parse(1)), CByte(Parse(2)), CByte(Parse(3)))
            Exit Sub
        
        Case CPackets.Cendjump
            Call Packet_EndJump(CLng(Parse(1)))
            Exit Sub
            
        Case CPackets.Cusespecialbadge
            Call Packet_UseSpecialBadge(Index, CLng(Parse(1)))
            Exit Sub
            
        Case CPackets.Cdodgebillspawn
            Call Packet_DodgeBillSpawn(Index, CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
            
        Case CPackets.Cnotifyotherplayer
            Call Packet_NotifyOtherPlayer(Index, Parse(1))
            Exit Sub
            
        Case CPackets.Cjugemscloudwarp
            Call Packet_JugemsCloudWarp(Index)
            Exit Sub
            
        Case CPackets.Cgetplayerinfo
            Call Packet_GetPlayerInfo(Index, Parse(1))
            Exit Sub
        
        Case CPackets.Cdoctorheal
            Call Packet_DoctorHeal(Index)
            Exit Sub
    End Select

    Call HackingAttempt(Index, "Received invalid packet: " & Parse(0))
    Exit Sub
    
Hacking:
    If IsNumeric(Parse(0)) = False Then
        Call HackingAttempt(Index, "Received invalid packet: " & Parse(0))
    End If
End Sub

Public Sub Packet_GetClasses(ByVal Index As Long)
    Call SendNewCharClasses(Index)
End Sub

Public Sub Packet_NewAccount(ByVal Index As Long, ByVal Username As String, ByVal Password As String, Optional ByVal PlayerValidate As Boolean = False)
    ' Protect against unauthorized account creation
    If PlayerValidate = False Then
        Call HackingAttempt(Index, "Unauthorized account creation.")
        Exit Sub
    End If
    
    If Not IsLoggedIn(Index) Then
        If LenB(Username) < 6 Then
            Call PlainMsg(Index, "Your username must be at least three characters in length.", 1)
            Exit Sub
        End If

        If LenB(Password) < 6 Then
            Call PlainMsg(Index, "Your password must be at least three characters in length.", 1)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(Index, "Your username must consist of alpha-numeric characters!", 1)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(Index, "Your password must consist of alpha-numeric characters!", 1)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call AddAccount(Index, Username, Password)
            Call PlainMsg(Index, "Your account has been created!", 0)
        Else
            Call PlainMsg(Index, "Sorry, that account name is already taken!", 1)
        End If
    End If
End Sub

Public Sub Packet_DeleteAccount(ByVal Index As Long, ByVal Username As String, ByVal Password As String, Optional ByVal PlayerValidate As Boolean = False)
    ' Protect against unauthorized account deletion
    If PlayerValidate = False Then
        Call HackingAttempt(Index, "Unauthorized account deletion.")
        Exit Sub
    End If
    
    If Not IsLoggedIn(Index) Then
        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(Index, "Your username must consist of alpha-numeric characters!", 2)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(Index, "Your password must consist of alpha-numeric characters!", 2)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call PlainMsg(Index, "That account name does not exist.", 2)
            Exit Sub
        End If

        If Not PasswordOK(Username, Password) Then
            Call PlainMsg(Index, "You've entered an incorrect password.", 2)
            Exit Sub
        End If
    
        Call LoadPlayer(Index, Username)
        
        Dim CharName As String, StringIndex As String
        Dim i As Long, a As Long
        
        For i = 1 To MAX_CHARS
            CharName = Trim$(Player(Index).Char(i).Name)
            
            If LenB(CharName) <> 0 Then
                ' **** Deletes stored INI data for characters ****
                
                ' Level Up
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\Level Up.ini")
                ' Respawn Points
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\Respawn Points.ini")
                ' Sounds
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\Sounds.ini")
                ' Friends List
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\SMBOAccounts\" & "Friend Lists.ini")
                ' Favors
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\Scripts\" & "Quests.ini")
                ' Question Blocks
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\Question Blocks.ini")
                ' Welcome Message
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\Scripts\" & "WelcomeMsg.ini")
                ' Warnings
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\Warn.ini")
                ' Whack-A-Monty Hall Of Fame
                Call WritePrivateProfileString(CharName, vbNullString, vbNullString, App.Path & "\Scripts\" & "WhackFame.ini")
                
                ' Cards
                For a = 94 To MAX_ITEMS
                    StringIndex = CStr(a)
                    
                    If GetVar(App.Path & "\Scripts\" & "Cards.ini", StringIndex, CharName) = "Has" Then
                        Call PutVar(App.Path & "\Scripts\" & "Cards.ini", StringIndex, CharName, vbNullString)
                    End If
                Next a
                
                ' Recipes
                For a = 1 To MAX_RECIPES
                    StringIndex = CStr(a)
                    
                    If GetVar(App.Path & "\Scripts\" & "Recipes.ini", StringIndex, CharName) = "Has" Then
                        Call PutVar(App.Path & "\Scripts\" & "Recipes.ini", StringIndex, CharName, vbNullString)
                    End If
                Next a
                
                Call DeleteName(CharName)
            End If
        Next i
        
        Call ClearPlayer(Index)

        ' Remove the users main player profile.
        Kill App.Path & "\SMBOAccounts\" & Username & "_Info.ini"
        Kill App.Path & "\SMBOAccounts\" & Username & "\*.*"

        ' Delete the users account directory.
        RmDir App.Path & "\SMBOAccounts\" & Username & "\"
    
        Call PlainMsg(Index, "Your account has been deleted.", 0)
    End If
End Sub

Public Sub Packet_AccountLogin(ByVal Index As Long, ByVal Username As String, ByVal Password As String, ByVal Major As Long, ByVal Minor As Long, ByVal Revision As Long, ByVal Code As String, ByVal OwnerStatus As Boolean)
    If Not IsLoggedIn(Index) Then
        If Major < CLIENT_MAJOR Or Minor < CLIENT_MINOR Or Revision < CLIENT_REVISION Then
            Call PlainMsg(Index, "Your version of the game is outdated. Please visit www.supermariobrosonline.co.cc for the most recent news.", 3)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(Index, "Your username must consist of alpha-numeric characters!", 3)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(Index, "Your password must consist of alpha-numeric characters!", 3)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call PlainMsg(Index, "That account name does not exist.", 3)
            Exit Sub
        End If
    
        If Username = "Kimimaru" Or Username = "hydrakiller4000" Then
            If OwnerStatus = False Then
                Call AlertMsg(Index, "You are not authorized to use this account!")
                Exit Sub
            End If
        End If
    
        If Not PasswordOK(Username, Password) Then
            Call PlainMsg(Index, "You've entered an incorrect password.", 3)
            Exit Sub
        End If
        
        If IsMultiAccounts(Username) Then
            Call PlainMsg(Index, "Multiple account logins is not authorized.", 3)
            Exit Sub
        End If
    
        If frmServer.Closed.Value = Checked Then
            Call PlainMsg(Index, "The server is closed at the moment!", 3)
            Exit Sub
        End If
    
        If Code <> SEC_CODE Then
            Call AlertMsg(Index, "The client password does not match the server password.")
            Exit Sub
        End If
    
        Call LoadPlayer(Index, Username)
        Call SendChars(Index)
    
        Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
    End If
End Sub

Public Sub Packet_GiveMeTheMax(ByVal Index As Long)
    Dim packet As String

    packet = SPackets.Smaxinfo & SEP_CHAR & LEVEL & END_CHAR

    Call SendDataTo(Index, packet)
    Call SendNewsTo(Index)
End Sub

Public Sub Packet_AddCharacter(ByVal Index As Long, ByVal Name As String, ByVal Sex As Long, ByVal Class As Long, ByVal CharNum As Long)
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(Index, "Invalid CharNum")
        Exit Sub
    End If
    
    If LenB(Name) < 6 Then
        Call HackingAttempt(Index, "Invalid Name Length")
        Exit Sub
    End If
    
    If Class < 0 Or Class > MAX_CLASSES Then
        Call HackingAttempt(Index, "Invalid Class")
        Exit Sub
    End If

    If Not IsAlphaNumeric(Name) Then
        Call PlainMsg(Index, "Your username must consist of alpha-numeric characters!", 4)
        Exit Sub
    End If

    If CharExist(Index, CharNum) Then
        Call PlainMsg(Index, "Character already exists!", 4)
        Exit Sub
    End If
    
    If FindChar(Name) Then
        Call PlainMsg(Index, "Sorry, but that name is in use!", 4)
        Exit Sub
    End If

    Call ClearChar(Index, CharNum)
    Call AddChar(Index, Name, Sex, Class, CharNum)

    Call SendChars(Index)
    Call PlainMsg(Index, "Character has been created!", 5)
End Sub

Public Sub Packet_DeleteCharacter(ByVal Index As Long, ByVal CharNum As Long)
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(Index, "Invalid CharNum")
        Exit Sub
    End If
    
    If CharExist(Index, CharNum) Then
        Call DelChar(Index, CharNum)
        Call SendChars(Index)
    
        Call PlainMsg(Index, "Character has been deleted!", 5)
    Else
        Call PlainMsg(Index, "Character does not exist!", 5)
    End If
End Sub

Public Sub Packet_UseCharacter(ByVal Index As Long, ByVal CharNum As Long)
    Dim FileID As Integer

    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(Index, "Invalid CharNum")
        Exit Sub
    End If
    
    If CharExist(Index, CharNum) Then
        Player(Index).CharNum = CharNum
    
        If frmServer.GMOnly.Value = Checked Then
            If GetPlayerAccess(Index) < 2 Then
                Call PlainMsg(Index, "The server is only available to GMs at the moment!", 5)
                Exit Sub
            End If
        End If
    
        Call JoinGame(Index)

        Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing Super Mario Bros. Online.", True)
        Call UpdateTOP
    
        If Not FindChar(GetPlayerName(Index)) Then
            FileID = FreeFile
            Open App.Path & "\SMBOAccounts\CharList.txt" For Append As #FileID
                Print #FileID, GetPlayerName(Index)
            Close #FileID
        End If
    Else
        Call PlainMsg(Index, "Character does not exist!", 5)
    End If
End Sub

Public Sub Packet_GuildChangeAccess(ByVal Index As Long, ByVal Name As String, ByVal Rank As Long)
    Dim NameIndex As Long
    
    If LenB(Name) = 0 Then
        Call PlayerMsg(Index, "You must enter a player name to proceed.", WHITE)
        Exit Sub
    End If

    If Rank < 0 Or Rank > 4 Then
        Call PlayerMsg(Index, "You must provide a valid rank to proceed.", RED)
        Exit Sub
    End If

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(Index) Then
        Call PlayerMsg(Index, Name & " is not in your Group.", RED)
        Exit Sub
    End If

    If GetPlayerGuildAccess(Index) < 4 Then
        Call PlayerMsg(Index, "You are not the owner of this Group.", RED)
        Exit Sub
    End If

    Call SetPlayerGuildAccess(NameIndex, Rank)
    Call SendPlayerData(NameIndex)
    Call SendGuildMemberHP(Index)
End Sub

Public Sub Packet_GuildDisown(ByVal Index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(Index) Then
        Call PlayerMsg(Index, Name & " is not in your Group.", RED)
        Exit Sub
    End If

    If GetPlayerGuildAccess(NameIndex) > GetPlayerGuildAccess(Index) Then
        Call PlayerMsg(Index, Name & " has a higher Group level than you.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, vbNullString)
    Call SetPlayerGuildAccess(NameIndex, 0)
    Call SendPlayerData(NameIndex)
    Call SendGuildMemberHP(Index)
End Sub

Public Sub Packet_GuildLeave(ByVal Index As Long)
    If LenB(GetPlayerGuild(Index)) = 0 Then
        Call PlayerMsg(Index, "You are not in a Group.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(Index, vbNullString)
    Call SetPlayerGuildAccess(Index, 0)
    Call PlayerMsg(Index, "You have left your Group!", WHITE)
    Call SendPlayerData(Index)
End Sub

Public Sub Packet_GuildMake(ByVal Index As Long, ByVal Name As String, ByVal Guild As String)
    Dim NameIndex As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    NameIndex = FindPlayer(Name)
    
    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) <> 0 Then
        Call PlayerMsg(Index, Name & " is already in a Group.", RED)
        Exit Sub
    End If

    If LenB(Guild) = 0 Then
        Call PlayerMsg(Index, "Please enter a valid Group name.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, Guild)
    Call SetPlayerGuildAccess(NameIndex, 4)
    Call SendPlayerData(NameIndex)
    Call SendGuildMemberHP(Index)
End Sub

Public Sub Packet_GuildMemberRequest(ByVal Index As Long, ByVal Name As String, ByVal Trainee As Integer)
    Dim PlayerIndex As Long
    Dim Message As String

    PlayerIndex = FindPlayer(Name)
    
    ' Trainee = 1 - trainee request
    ' Trainee = 0 - not trainee request
    
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If
    
    If GetPlayerGuild(PlayerIndex) <> vbNullString And GetPlayerGuild(PlayerIndex) <> GetPlayerGuild(Index) Then
        Call PlayerMsg(Index, Name & " is already in a Group!", RED)
        Exit Sub
    End If

    If GetPlayerGuild(PlayerIndex) = GetPlayerGuild(Index) Then
        Call PlayerMsg(Index, Name & " has already been admitted to your Group.", WHITE)
        Exit Sub
    End If
    
    Message = GetPlayerName(Index) & " wants to invite you to join the Group: " & GetPlayerGuild(Index) & "."
    
    Call SendDataTo(PlayerIndex, SPackets.Sguildmemberrequest & SEP_CHAR & Message & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & Trainee & END_CHAR)
End Sub

Public Sub Packet_GuildMemberDecline(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long
    Dim PlayerName As String
    
    PlayerIndex = FindPlayer(Name)
    PlayerName = GetPlayerName(Index)
    
    ' Recruiter must've logged off
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If
    
    Call PlayerMsg(PlayerIndex, PlayerName & " has declined your Group member invitation.", WHITE)
End Sub

Public Sub Packet_GuildMember(ByVal Index As Long, ByVal Name As String)
    Dim NameIndex As Long
    Dim PlayerName As String
    
    ' Index = player you're recruiting in this Sub
    
    NameIndex = FindPlayer(Name)
    PlayerName = GetPlayerName(Index)
    
    ' Recruiter must've logged off
    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If
    
    If GetPlayerGuild(Index) <> vbNullString And GetPlayerGuild(Index) <> GetPlayerGuild(NameIndex) Then
        Call PlayerMsg(NameIndex, PlayerName & " is already in a Group.", RED)
        Exit Sub
    End If

    If GetPlayerGuild(Index) = GetPlayerGuild(NameIndex) Then
        Call PlayerMsg(NameIndex, PlayerName & " has already been admitted to your Group.", WHITE)
        Exit Sub
    End If
    
    Call SetPlayerGuild(Index, GetPlayerGuild(NameIndex))
    Call SetPlayerGuildAccess(Index, 1)
    Call SendPlayerData(Index)
    Call PlayerMsg(Index, "You have become a member of the Group: " & GetPlayerGuild(NameIndex) & "!", WHITE)
    Call PlayerMsg(NameIndex, PlayerName & " has accepted your Group invitation!", WHITE)
    Call SendGuildMemberHP(NameIndex)
End Sub

Public Sub Packet_GuildTrainee(ByVal Index As Long, ByVal Name As String)
    Dim NameIndex As Long
    Dim PlayerName As String

    ' Index = player you're recruiting in this Sub

    NameIndex = FindPlayer(Name)
    PlayerName = GetPlayerName(Index)
    
    ' Recruiter must've logged off
    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If
    
    If GetPlayerGuild(Index) <> vbNullString And GetPlayerGuild(Index) <> GetPlayerGuild(NameIndex) Then
        Call PlayerMsg(NameIndex, PlayerName & " is already in a Group.", RED)
        Exit Sub
    End If

    If GetPlayerGuild(Index) = GetPlayerGuild(NameIndex) Then
        Call PlayerMsg(NameIndex, PlayerName & " has already been admitted to your Group.", WHITE)
        Exit Sub
    End If
    
    Call SetPlayerGuild(Index, GetPlayerGuild(NameIndex))
    Call SetPlayerGuildAccess(Index, 0)
    Call SendPlayerData(Index)
    Call PlayerMsg(Index, "You have become a member of the Group: " & GetPlayerGuild(NameIndex) & "!", WHITE)
    Call PlayerMsg(NameIndex, PlayerName & " has accepted your Group invitation!", WHITE)
    Call SendGuildMemberHP(NameIndex)
End Sub

Public Sub Packet_SayMessage(ByVal Index As Long, ByVal Message As String)
    ' Changes player's text color in the chatbox based on rank
    Select Case GetPlayerAccess(Index)
        Case 0
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Message, GREEN)
        Case 1
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Message, BLACK)
        Case 2
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Message, BRIGHTBLUE)
        Case 3
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Message, BROWN)
        Case 4
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Message, WHITE)
        Case 5
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Message, YELLOW)
    End Select

    Call MapMsg2(GetPlayerMap(Index), Message, Index)

    Call TextAdd(frmServer.txtText(3), GetPlayerName(Index) & " On Map " & GetPlayerMap(Index) & ": " & Message, True)
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " : " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_GroupMessage(ByVal Index As Long, ByVal Message As String)
    Dim i As Long
    
    If Len(GetPlayerGuild(Index)) <= 0 Then
        Call PlayerMsg(Index, "You are not in a Group!", WHITE)
        Exit Sub
    End If
    
    For i = 1 To MAX_PLAYERS
        If GetPlayerGuild(i) = GetPlayerGuild(Index) Then
            Call PlayerMsg(i, "[" & GetPlayerGuild(Index) & "] " & GetPlayerName(Index) & ": " & Message, BRIGHTBLUE)
        End If
    Next i
            
    Call TextAdd(frmServer.txtText(6), GetPlayerName(Index) & ": " & Message, True)
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_BroadcastMessage(ByVal Index As Long, ByVal Message As String)
    If IsPlayerMuted(Index) Then
        Call PlayerMsg(Index, "You are muted, so you cannot send any global messages!", BRIGHTRED)
        Exit Sub
    End If
    
  ' Changes broadcast message color depending on rank
    Select Case GetPlayerAccess(Index)
        Case 0
            Call GlobalMsg("[GlobalMsg] " & GetPlayerName(Index) & ": " & Message, GREEN)
        Case 1
            Call GlobalMsg("[Moderator] " & GetPlayerName(Index) & ": " & Message, BLACK)
        Case 2
            Call GlobalMsg("[Designer] " & GetPlayerName(Index) & ": " & Message, BRIGHTBLUE)
        Case 3
            Call GlobalMsg("[Developer] " & GetPlayerName(Index) & ": " & Message, BROWN)
        Case 4
            Call GlobalMsg("[Administrator] " & GetPlayerName(Index) & ": " & Message, WHITE)
        Case 5
            Call GlobalMsg("[Creator] " & GetPlayerName(Index) & ": " & Message, YELLOW)
    End Select

    Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & ": " & Message, True)
    Call TextAdd(frmServer.txtText(1), GetPlayerName(Index) & ": " & Message, True)
    Call AddLog(GetPlayerName(Index) & ": " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_GlobalMessage(ByVal Index As Long, ByVal Message As String)
    If IsPlayerMuted(Index) Then
        Call PlayerMsg(Index, "You are muted, so you cannot send any global messages!", BRIGHTRED)
        Exit Sub
    End If

    If GetPlayerAccess(Index) > 0 Then
        Call GlobalMsg("(Global) " & GetPlayerName(Index) & ": " & Message, GlobalColor)

        Call TextAdd(frmServer.txtText(0), "(Global) " & GetPlayerName(Index) & ": " & Message, True)
        Call TextAdd(frmServer.txtText(2), GetPlayerName(Index) & ": " & Message, True)
        Call AddLog("(Global) " & GetPlayerName(Index) & ": " & Message, ADMIN_LOG)
    End If
End Sub

Public Sub Packet_AdminMessage(ByVal Index As Long, ByVal Message As String)
    If GetPlayerAccess(Index) > 0 Then
        Call AdminMsg("[Development Chat] " & GetPlayerName(Index) & ": " & Message, AdminColor)

        Call TextAdd(frmServer.txtText(5), GetPlayerName(Index) & ": " & Message, True)
        Call AddLog("[Development Chat] " & GetPlayerName(Index) & ": " & Message, ADMIN_LOG)
    End If
End Sub

Public Sub Packet_PlayerMessage(ByVal Index As Long, ByVal Name As String, ByVal Message As String)
    If IsPlayerMuted(Index) Then
        Call PlayerMsg(Index, "You are muted, so you cannot send any global messages!", BRIGHTRED)
        Exit Sub
    End If
    
    Dim MsgTo As Long

    If LenB(Name) = 0 Then
        Call PlayerMsg(Index, "You must select a player name to private message.", BRIGHTRED)
        Exit Sub
    End If

    If LenB(Message) = 0 Then
        Call PlayerMsg(Index, "You must send a message to private message another player.", BRIGHTRED)
        Exit Sub
    End If

    MsgTo = FindPlayer(Name)

    If MsgTo = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    Call PlayerMsg(Index, "[PM] To " & GetPlayerName(MsgTo) & ": " & Message, BRIGHTCYAN)
    Call PlayerMsg(MsgTo, "[PM] From " & GetPlayerName(Index) & ": " & Message, BRIGHTCYAN)

    Call TextAdd(frmServer.txtText(4), "[PM] To " & GetPlayerName(MsgTo) & "[PM] From " & GetPlayerName(Index) & ": " & Message, True)
    Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Message & "'", PLAYER_LOG)
End Sub

Public Sub Packet_OtherMessage(ByVal Index As Long, ByVal Name As String, ByVal Message As String)
    If IsPlayerMuted(Index) Then
        Call PlayerMsg(Index, "You are muted, so you cannot send any global messages!", BRIGHTRED)
        Exit Sub
    End If
    
    Dim MsgTo As Long

    MsgTo = FindPlayer(Name)

    If MsgTo = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If
    
    Call OtherMsg(Index, "[PM] To " & GetPlayerName(MsgTo) & ": " & Message, BRIGHTCYAN)
    Call OtherMsg(MsgTo, "[PM] From " & GetPlayerName(Index) & ": " & Message, BRIGHTCYAN)
    
    Call TextAdd(frmServer.txtText(4), "[PM] To " & GetPlayerName(MsgTo) & "[PM] From " & GetPlayerName(Index) & ": " & Message, True)
    Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Message & "'", PLAYER_LOG)
End Sub

Public Sub Packet_PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long, Xpos As Long, Ypos As Long)
    If Player(Index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If

    If Movement <> 1 And Movement <> 2 Then
        Call HackingAttempt(Index, "Invalid Movement")
        Exit Sub
    End If

    If Player(Index).Locked = True Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, Movement, Xpos, Ypos)
End Sub

Public Sub Packet_PlayerDirection(ByVal Index As Long, ByVal Dir As Long)
    If Player(Index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)

    Call SendDataToMapBut(Index, GetPlayerMap(Index), SPackets.Splayerdir & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
End Sub

Public Sub Packet_UseItem(ByVal Index As Long, ByVal InvNum As Long)
    Dim n As Long, ItemNum As Long, EquippedItem As Long, CharNum As Long, SpellID As Long, MinLvl As Long, X As Long, Y As Long, AmmoValue As Long
    Dim ReloadAmount As Long, AmmoCapacity As Long
    Dim CanUseAmmo As Integer, HasRangedWeapon As Integer
    
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If

    If Player(Index).LockedItems Then
        Call PlayerMsg(Index, "You currently cannot use any items.", BRIGHTRED)
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, InvNum)
    
    If ItemNum = 57 And Trim$(GetVar(App.Path & "\Scripts\" & "WhackFame.ini", GetPlayerName(Index), "In")) <> "Yes" Then
        Call PlayerMsg(Index, "You have not earned enough points in Whack-A-Monty to be able to equip this item!", WHITE)
        Exit Sub
    End If
    
    ' Stop players from equipping Power Rush and Damage Dodge if they haven't earned at least 50 points
    If (ItemNum = 85 Or ItemNum = 86) And Val(GetVar(App.Path & "\Scripts\" & "WhackFame.ini", GetPlayerName(Index), "Points")) < 50 Then
        Call PlayerMsg(Index, "You have not earned enough points in Whack-A-Monty to be able to equip this item!", WHITE)
        Exit Sub
    End If
    
    CharNum = Player(Index).CharNum

    ' Stops players from using the item if it's not usable
    If Item(ItemNum).Type >= ITEM_TYPE_WEAPON And Item(ItemNum).Type <= ITEM_TYPE_MUSHROOMBADGE Then
        If ItemIsUsable(Index, InvNum) = False Then
            Exit Sub
        End If
    End If
    
    EquippedItem = 0
    
    ' Find out what kind of item it is
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_WEAPON
            If Item(ItemNum).Stackable <> 2 Then
                Call EquipItem(Index, InvNum, 1)
            Else
                Call EquipItem(Index, InvNum, 4)
            End If
        Case ITEM_TYPE_ARMOR
            Call EquipItem(Index, InvNum, 2)
        Case ITEM_TYPE_HELMET
            Call EquipItem(Index, InvNum, 3)
        Case ITEM_TYPE_SPECIALBADGE
            Call EquipItem(Index, InvNum, 4)
        Case ITEM_TYPE_LEGS
            Call EquipItem(Index, InvNum, 5)
        Case ITEM_TYPE_FLOWERBADGE
            Call EquipItem(Index, InvNum, 6)
        Case ITEM_TYPE_MUSHROOMBADGE
            Call EquipItem(Index, InvNum, 7)
        Case ITEM_TYPE_CHANGEHPFPSP
            Call SendSoundTo(Index, "spm_get_health.wav")
            Call SpellAnim(4, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
              ' Send message
                If FindItemVowels(ItemNum) = True Then
                    Call PlayerMsg(Index, "You used an " & Trim$(Item(ItemNum).Name) & "!", YELLOW)
                Else
                    Call PlayerMsg(Index, "You used a " & Trim$(Item(ItemNum).Name) & "!", YELLOW)
                End If
            Call SetPlayerHP(Index, GetPlayerHP(Index) + Item(ItemNum).Data1)
            Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(ItemNum).Data2)
            Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(ItemNum).Data3)
            
            Call TakeSpecificItem(Index, InvNum, 1)
        
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            
            ' Kill player if the item lowered his/her HP below 1
            If GetPlayerHP(Index) <= 0 Then
                If GetPlayerInBattle(Index) = True Then
                    Call TurnBasedDeath(Index, GetPlayerMap(Index), Player(Index).TargetNPC, Abs(Item(ItemNum).Data1))
                Else
                    Call PlayerDeath(Index)
                End If
            End If
        Case ITEM_TYPE_KEY
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If GetPlayerY(Index) > 0 Then
                        X = GetPlayerX(Index)
                        Y = GetPlayerY(Index) - 1
                    Else
                        Exit Sub
                    End If
    
                Case DIR_DOWN
                    If GetPlayerY(Index) < MAX_MAPY Then
                        X = GetPlayerX(Index)
                        Y = GetPlayerY(Index) + 1
                    Else
                        Exit Sub
                    End If
    
                Case DIR_LEFT
                    If GetPlayerX(Index) > 0 Then
                        X = GetPlayerX(Index) - 1
                        Y = GetPlayerY(Index)
                    Else
                        Exit Sub
                    End If
    
                Case DIR_RIGHT
                    If GetPlayerX(Index) < MAX_MAPX Then
                        X = GetPlayerX(Index) + 1
                        Y = GetPlayerY(Index)
                    Else
                        Exit Sub
                    End If
            End Select
    
            ' Check if a key exists.
            If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY Then
                ' Check if the key they are using matches the map key.
                If ItemNum = Map(GetPlayerMap(Index)).Tile(X, Y).Data1 Then
                    TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                    TempTile(GetPlayerMap(Index)).DoorTimer(X, Y) = GetTickCount
                    
                    Call SendMapKey(GetPlayerMap(Index), X, Y, 1)
                    
                    If Trim$(Map(GetPlayerMap(Index)).Tile(X, Y).String1) <> vbNullString Then
                        Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(X, Y).String1), WHITE)
                    End If

                    Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "key" & END_CHAR)
    
                    ' Check if we are supposed to take away the item.
                    If Map(GetPlayerMap(Index)).Tile(X, Y).Data2 = 1 Then
                        Call TakeItem(Index, ItemNum, 1)
                    End If
                End If
            End If
    
            If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer(X, Y) = GetTickCount
                
                Call SendMapKey(GetPlayerMap(Index), X, Y, 1)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "key" & END_CHAR)
            End If
        Case ITEM_TYPE_SPELL
            SpellID = Item(ItemNum).Data1
    
            If SpellID > 0 Then
                If Spell(SpellID).ClassReq - 1 = GetPlayerClass(Index) Or Spell(SpellID).ClassReq = 0 Then
                    MinLvl = GetSpellReqLevel(SpellID)

                    If MinLvl <= GetPlayerLevel(Index) Then
                        MinLvl = FindOpenSpellSlot(Index)
    
                        If MinLvl > 0 Then
                            If Not HasSpell(Index, SpellID) Then
                                Call SetPlayerSpell(Index, MinLvl, SpellID)
                                Call TakeSpecificItem(Index, InvNum, 0)
                                Call PlayerMsg(Index, "You've learned a new special attack!", WHITE)
                            Else
                                Call PlayerMsg(Index, "You've already learned this special attack!", WHITE)
                            End If
                        Else
                            Call PlayerMsg(Index, "You cannot learn anymore special attacks!", WHITE)
                        End If
                    Else
                        Call PlayerMsg(Index, "You must be level " & MinLvl & " to learn this special attack!", WHITE)
                    End If
                Else
                    Call PlayerMsg(Index, "This special attack can only be learned by " & GetClassName(Spell(SpellID).ClassReq - 1) & "!", WHITE)
                End If
            End If
        Case ITEM_TYPE_SCRIPTED
            Call ScriptedItem(Index, Item(Player(Index).Char(CharNum).NewInv(InvNum).num).Data1)
        Case ITEM_TYPE_AMMO
            AmmoValue = GetPlayerInvItemValue(Index, InvNum)
            CanUseAmmo = 0
            HasRangedWeapon = 0
            
            ' Check to see if we can reload the weapon first
            ItemNum = GetPlayerEquipSlotNum(Index, 1)
            If ItemNum > 0 Then
                If Item(ItemNum).Ammo > -1 Then
                    If Item(ItemNum).Data3 = Item(GetPlayerInvItemNum(Index, InvNum)).Data3 Then
                        AmmoCapacity = Item(ItemNum).Ammo
                        
                        If GetPlayerEquipSlotAmmo(Index, 1) < AmmoCapacity Then
                            ReloadAmount = AmmoCapacity - GetPlayerEquipSlotAmmo(Index, 1)
                            ' If they have enough ammo, take it and refill; otherwise, take whatever ammo they have and refill
                            If CanTake(Index, GetPlayerInvItemNum(Index, InvNum), ReloadAmount) Then
                                Call SetPlayerEquipSlotAmmo(Index, 1, AmmoCapacity)
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), ReloadAmount)
                                Call PlayerMsg(Index, "You've reloaded your weapon with " & ReloadAmount & " ammo!", WHITE)
                                Call SendEquipmentUpdate(Index, 1)
                                Exit Sub
                            Else
                                Call SetPlayerEquipSlotAmmo(Index, 1, GetPlayerEquipSlotAmmo(Index, 1) + AmmoValue)
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), AmmoValue)
                                Call PlayerMsg(Index, "You've reloaded your weapon with " & AmmoValue & " ammo!", WHITE)
                                Call SendEquipmentUpdate(Index, 1)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
                            
            For n = 1 To GetPlayerMaxInv(Index)
                ItemNum = GetPlayerInvItemNum(Index, n)
                ' Makes sure there is an item in the slot
                If ItemNum > 0 Then
                ' Makes sure the weapon needs ammo
                    If Item(ItemNum).Ammo > -1 Then
                        HasRangedWeapon = 1
                        ' Makes sure the ammo is the same as the bow's ammo
                        If Item(ItemNum).Data3 = Item(GetPlayerInvItemNum(Index, InvNum)).Data3 Then
                            AmmoCapacity = Item(ItemNum).Ammo
                            CanUseAmmo = 1
                            ' Checks to see if weapon is full of ammo
                            If GetPlayerInvItemAmmo(Index, n) < AmmoCapacity Then
                                ReloadAmount = AmmoCapacity - GetPlayerInvItemAmmo(Index, n)
                                CanUseAmmo = -1
                                ' If they have enough ammo, take it and refill; otherwise, take whatever ammo they have and refill
                                If CanTake(Index, GetPlayerInvItemNum(Index, InvNum), ReloadAmount) Then
                                    Call SetPlayerInvItemAmmo(Index, n, AmmoCapacity)
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), ReloadAmount)
                                    Call PlayerMsg(Index, "You've reloaded your weapon with " & ReloadAmount & " ammo!", WHITE)
                                    Call SendInventoryUpdate(Index, n)
                                    Exit For
                                Else
                                    Call SetPlayerInvItemAmmo(Index, n, GetPlayerInvItemAmmo(Index, n) + AmmoValue)
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), AmmoValue)
                                    Call PlayerMsg(Index, "You've reloaded your weapon with " & AmmoValue & " ammo!", WHITE)
                                    Call SendInventoryUpdate(Index, n)
                                    Exit For
                                End If
                            End If
                        Else
                            CanUseAmmo = CanUseAmmo
                        End If
                    Else
                        HasRangedWeapon = HasRangedWeapon
                    End If
                End If
            Next n
            
            If CanUseAmmo > 0 Then
                Call PlayerMsg(Index, "All of your weapons that use this type of ammo are already completely filled!", WHITE)
            ElseIf CanUseAmmo = 0 And HasRangedWeapon > 0 Then
                Call PlayerMsg(Index, "You cannot use a different type of ammo to reload your weapon!", WHITE)
            End If
            If HasRangedWeapon = 0 Then
                Call PlayerMsg(Index, "There are no weapons in your inventory that would require reloading!", WHITE)
            End If
        Case ITEM_TYPE_CARD
            If GetVar(App.Path & "\Scripts\" & "Cards.ini", CInt(ItemNum), GetPlayerName(Index)) <> "Has" Then
                Call PutVar(App.Path & "\Scripts\" & "Cards.ini", CInt(ItemNum), GetPlayerName(Index), "Has")
                Call PlayerMsg(Index, "This card has successfully been added into your collection!", WHITE)
            End If
    End Select

    Call SendStats(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)

    Call SendIndexWornEquipment(Index)
End Sub

' This packet seems to me like it's incomplete. [Mellowz]
Public Sub Packet_PlayerMoveMouse(ByVal Index As Long, ByVal Dir As Long)
    If Player(Index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If

    If Player(Index).Locked = True Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    If Player(Index).CastedSpell = YES Then
        If GetTickCount > Player(Index).AttackTimer + 1000 Then
            Player(Index).CastedSpell = NO
        Else
            Call SendPlayerXY(Index)
            Exit Sub
        End If
    End If
End Sub

Public Sub Packet_Warp(ByVal Index As Long, ByVal Dir As Long)
    Select Case Dir
        Case DIR_UP
            If Map(GetPlayerMap(Index)).Up > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), MAX_MAPY)
                Exit Sub
            End If

        Case DIR_DOWN
            If Map(GetPlayerMap(Index)).Down > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                Exit Sub
            End If

        Case DIR_LEFT
            If Map(GetPlayerMap(Index)).Left > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, MAX_MAPX, GetPlayerY(Index))
                Exit Sub
            End If

        Case DIR_RIGHT
            If Map(GetPlayerMap(Index)).Right > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                Exit Sub
            End If
    End Select
End Sub

Public Sub SpinFinish(ByVal Index As Long, ByVal Success As Long)
    ' Play the Spin sound
    Call SendSoundToMap(GetPlayerMap(Index), "m&lss_spin.wav")
    
    If Success = 0 Then
        Player(Index).HookShotX = 0
        Player(Index).HookShotY = 0
        
        Call SpellAnim(2, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Exit Sub
    End If

    Call PlayerMsg(Index, "You use the spin ability to reach your destination.", WHITE)
    
    Call SpellAnim(2, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    Call SetPlayerX(Index, Player(Index).HookShotX)
    Call SetPlayerY(Index, Player(Index).HookShotY)

    Player(Index).HookShotX = 0
    Player(Index).HookShotY = 0

    Call SendPlayerXY(Index)
End Sub

Public Sub Packet_Attack(ByVal Index As Long)
    Dim i As Long, Damage As Long, ItemNum As Long
    Dim X As Byte, Y As Byte, tX As Byte, tY As Byte
    
    If Player(Index).LockedAttack Then
        Exit Sub
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Index, GetPlayerMap(Index), SPackets.Sattack & SEP_CHAR & Index & END_CHAR)
    
    If GetPlayerInBattle(Index) = True Then
        Call TurnBasedBattle(Index, GetPlayerTargetNpc(Index))
        Exit Sub
    End If
    
    ItemNum = GetPlayerEquipSlotNum(Index, 1)
    
    ' Check if the player is in the Dodgebill minigame
    If GetPlayerMap(Index) = 188 Then
        ' Check for the Bullet Bill item
        If HasItem(Index, 186) >= 1 Then
            ' Make sure no weapon is equipped
            If ItemNum <= 0 Then
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Scheckarrows & SEP_CHAR & Index & SEP_CHAR & Item(186).Data3 & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
                Call TakeItem(Index, 186, 1)
            Else
                Call PlayerMsg(Index, "You cannot throw a bullet bill while you're holding a weapon!", BRIGHTRED)
            End If
        End If
        
        Exit Sub
    End If
    
    ' Check if the player is trying to get the key from the tree
    If GetPlayerMap(Index) = 194 Then
        ' Make sure the player didn't already get it
        If GetVar(App.Path & "\Scripts\" & "Key.ini", GetPlayerName(Index), "Key") <> "Got" Then
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SCRIPTED Then
                ' Make sure the tile is script #25
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 = 25 Then
                    ' Make sure the player is facing up
                    If GetPlayerDir(Index) = DIR_UP Then
                        ' Make sure the player has enough inventory space
                        If GetFreeSlots(Index) > 0 Then
                            Call GiveItem(Index, 227, 1)
                            
                            Call PlayerMsg(Index, "A " & Trim$(Item(227).Name) & " fell out of the tree!", YELLOW)
                            Call PutVar(App.Path & "\Scripts\" & "Key.ini", GetPlayerName(Index), "Key", "Got")
                            Exit Sub
                        Else
                            Call PlayerMsg(Index, "Something fell out of the tree, but you couldn't fit it in your inventory.", WHITE)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If ItemNum > 0 Then
        If Item(ItemNum).Data3 > 0 And Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
            If Item(ItemNum).Stackable = 0 Then
                If GetPlayerEquipSlotAmmo(Index, 1) > -1 Then
                    If GetPlayerEquipSlotAmmo(Index, 1) = 0 Then
                        Call PlayerMsg(Index, "You've ran out of ammo! Refill your weapon before you try to fight again.", WHITE)
                        Exit Sub
                    ElseIf GetPlayerEquipSlotAmmo(Index, 1) > 0 Then
                        If Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_MINIGAME Then
                            Call SetPlayerEquipSlotAmmo(Index, 1, GetPlayerEquipSlotAmmo(Index, 1) - 1)
                            Call SendEquipmentUpdate(Index, 1)
                        End If
                    End If
                End If
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Scheckarrows & SEP_CHAR & Index & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
            End If
            Exit Sub
        End If
    End If
    
    X = GetPlayerX(Index) + IIf(GetPlayerDir(Index) >= 2, (GetPlayerDir(Index) - 2) * 2 - 1, 0)
    Y = GetPlayerY(Index) + IIf(GetPlayerDir(Index) < 2, GetPlayerDir(Index) * 2 - 1, 0)

    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_SWITCH Then
        If (Map(GetPlayerMap(Index)).Tile(X, Y).Data3 And 1) = 1 Then ' Advanced Bit Logic, ask for help before changing this line.
            tX = Map(GetPlayerMap(Index)).Tile(X, Y).Data2 \ 256
            tY = Map(GetPlayerMap(Index)).Tile(X, Y).Data2 Mod 256
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Tile(X, Y).Data1, tX, tY)
            Exit Sub
        End If
    End If
    
    ' Try to attack another player.
    For i = 1 To MAX_PLAYERS
        If i <> Index Then
            If CanAttackPlayer(Index, i) Then
            
                Player(Index).Target = i
                Player(Index).TargetType = TARGET_TYPE_PLAYER
                
                If Not CanPlayerBlockHit(i) Then
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = Int(GetPlayerDamage(Index) - GetPlayerProtection(i))
                        Call SendAttackSound(GetPlayerMap(Index))
                    Else
                        Damage = Int((GetPlayerDamage(Index) - GetPlayerProtection(i)) * 1.5)

                        Call BattleMsg(Index, "Critical hit!", BRIGHTGREEN, 0)
                        Call BattleMsg(i, GetPlayerName(Index) & " got a critical hit on you!", RED, 1)

                        Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "critical" & END_CHAR)
                    End If
                    
                    Damage = DamageUpDamageDown(Index, Damage, i)
                    
                    ' Randomizes damage
                    Damage = Int(Rand(Damage - 2, Damage + 2))
                    
                    If Damage > 0 Then
                        Call OnAttack(Index, Damage)
                    Else
                        Call OnAttack(Index, 0)
                        
                        ' Stops sound from occurring in Steal the Shroom and Dodgebill
                        If GetPlayerMap(Index) <> 33 And GetPlayerMap(Index) <> 188 Then
                            Call PlayerMsg(Index, "Your attack was too weak to harm the player!", WHITE)
                            Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
                        End If
                    End If
                Else
                    Call OnAttack(Index, 0)
                    Call BattleMsg(Index, GetPlayerName(i) & " blocked your attack!", DARKGREY, 0)
                    Call BattleMsg(i, "You blocked " & GetPlayerName(Index) & "'s attack!", BRIGHTBLUE, 1)

                    Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
                End If
                
                Exit Sub
            End If
        End If
    Next i

    ' Try to attack an NPC.
    For i = 1 To MAX_MAP_NPCS
        If CanAttackNpc(Index, i) Then
            ' Get the damage we can do
            Player(Index).TargetNPC = i
            Player(Index).TargetType = TARGET_TYPE_NPC
            
            ' Begins turn-based battle system if possible
            If Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_MINIGAME Then
                ' Always enable turn-based battles in the Poison Cave
                If IsInPoisonCave(Index) = False Then
                    If GetPlayerInBattle(Index) = False And GetPlayerTurnBased(Index) = True Then
                        Call PlayerFirstStrike(Index, i)
                        Exit Sub
                    End If
                Else
                    If GetPlayerInBattle(Index) = False Then
                        Call PlayerFirstStrike(Index, i)
                        Exit Sub
                    End If
                End If
            End If
            
            If Not CanPlayerCriticalHit(Index) Then
                Damage = GetPlayerDamage(Index) - Int(NPC(MapNPC(GetPlayerMap(Index), i).num).DEF)
                Call SendAttackSound(GetPlayerMap(Index))
            Else
                Damage = Int((GetPlayerDamage(Index) - Int(NPC(MapNPC(GetPlayerMap(Index), i).num).DEF)) * 1.5)
                
                ' Stops critical hit message from occurring in Whack-A-Monty
                If GetPlayerMap(Index) <> 72 And GetPlayerMap(Index) <> 73 And GetPlayerMap(Index) <> 74 Then
                    Call BattleMsg(Index, "Critical hit!", BRIGHTGREEN, 0)
                End If
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "critical" & END_CHAR)
            End If
            
            Damage = DamageUpDamageDown(Index, Damage)
            
            ' Randomizes damage
            Damage = Int(Rand(Damage - 2, Damage + 2))
            
            ' Make it so you cannot hit less than 1 if your attack is greater than the target's defense
            If Damage <= 0 And GetPlayerSTR(Index) > Int(NPC(MapNPC(GetPlayerMap(Index), i).num).DEF) Then
                Damage = 1
            End If
            
            If Damage > 0 Then
                Call OnAttack(Index, Damage)
                Call SendDataTo(Index, SPackets.Sblitplayerdmg & SEP_CHAR & Damage & SEP_CHAR & i & END_CHAR)
            Else
                Call OnAttack(Index, 0)
                Call PlayerMsg(Index, "Your attack was too weak to harm the enemy!", WHITE)
                Call SendDataTo(Index, SPackets.Sblitplayerdmg & SEP_CHAR & 0 & SEP_CHAR & i & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
            End If

            Exit Sub
        End If
    Next i
End Sub

Public Sub Packet_UseStatPoint(ByVal Index As Long, ByVal PointType As Integer)
    If PointType < 0 Or PointType > 5 Then
        Call HackingAttempt(Index, "Invalid Point Type")
        Exit Sub
    End If

    If GetPlayerPOINTS(Index) > 0 Then
        Call UsingStatPoints(Index, PointType)
    End If

    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)

    Player(Index).Char(Player(Index).CharNum).MAXHP = GetPlayerMaxHP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXMP = GetPlayerMaxMP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXSP = GetPlayerMaxSP(Index)

    Call SendStats(Index)

    Call SendDataTo(Index, SPackets.Splayerpoints & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)

    If GetPlayerPOINTS(Index) = 0 Then
        Call SendDataTo(Index, SPackets.Sstatpointused & END_CHAR)
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
    End If
End Sub

Public Sub Packet_GetStats(ByVal Index As Long, ByVal Name As String)
    Dim PlayerID As Long, BlockChance As Double, CritChance As Double

    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerMsg(Index, "Account: " & Trim$(Player(PlayerID).Login) & "; Name: " & GetPlayerName(PlayerID), BRIGHTGREEN)

        If GetPlayerAccess(Index) > ADMIN_MONITER Then
            Call PlayerMsg(Index, "Stats for " & GetPlayerName(PlayerID) & ":", BRIGHTGREEN)
            Call PlayerMsg(Index, "Level: " & GetPlayerLevel(PlayerID) & "; EXP: " & GetPlayerExp(PlayerID) & "/" & GetPlayerNextLevel(PlayerID), BRIGHTGREEN)
            Call PlayerMsg(Index, "HP: " & GetPlayerHP(PlayerID) & "/" & GetPlayerMaxHP(PlayerID) & "; FP: " & GetPlayerMP(PlayerID) & "/" & GetPlayerMaxMP(PlayerID) & "; SP: " & GetPlayerSP(PlayerID) & "/" & GetPlayerMaxSP(PlayerID), BRIGHTGREEN)
            Call PlayerMsg(Index, "STR: " & GetPlayerSTR(PlayerID) & "; DEF: " & GetPlayerDEF(PlayerID) & "; STCH: " & GetPlayerStache(PlayerID) & "; SPD: " & GetPlayerSPEED(PlayerID), BRIGHTGREEN)
            
            CritChance = Round(GetPlayerCritHitChance(PlayerID), 2)
            If CritChance < 0 Then
                CritChance = 0
            End If
            If CritChance > 100 Then
                CritChance = 100
            End If

            BlockChance = Round(GetPlayerBlockChance(PlayerID), 2)
            If BlockChance < 0 Then
                BlockChance = 0
            End If
            If BlockChance > 100 Then
                BlockChance = 100
            End If

            Call PlayerMsg(Index, "Critical Chance: " & CritChance & "%; Block Chance: " & BlockChance & "%", BRIGHTGREEN)
        End If
    Else
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_SetPlayerSprite(ByVal Index As Long, ByVal Name As String, ByVal SpriteID As Long)
    Dim PlayerID As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call SetPlayerSprite(PlayerID, SpriteID)
        Call SendPlayerData(PlayerID)
    Else
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_RequestNewMap(ByVal Index As Long, ByVal Dir As Long)
    Dim X As Long, Y As Long
    
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
    
    Y = GetPlayerY(Index)
    X = GetPlayerX(Index)

    Select Case Dir
        Case DIR_UP
            Y = Y - 1
        Case DIR_DOWN
            Y = Y + 1
        Case DIR_LEFT
            X = X - 1
        Case DIR_RIGHT
            X = X + 1
    End Select
    
     Call PlayerMove(Index, Dir, 1, X, Y)
     Call SendPlayerNewXY(Index)
End Sub

Public Sub Packet_WarpMeTo(ByVal Index As Long, ByVal Name As String)
    Dim PlayerID As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerWarp(Index, GetPlayerMap(PlayerID), GetPlayerX(PlayerID), GetPlayerY(PlayerID))
    Else
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_WarpToMe(ByVal Index As Long, ByVal Name As String)
    Dim PlayerID As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerWarp(PlayerID, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    Else
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_MapData(ByVal Index As Long, ByRef MapData() As String)
    Dim MapIndex As Long, MapNum As Long, MapRevision As Long, X As Long, Y As Long, i As Long
    
    ' Check to see if the user is at least a mapper.
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
            
    MapNum = GetPlayerMap(Index)
            
    ' Get revision number before it clears
    MapRevision = Map(MapNum).Revision + 1
            
    MapIndex = 1

    Call ClearMap(MapNum)

    MapNum = CLng(MapData(MapIndex))
    Map(MapNum).Name = MapData(MapIndex + 1)
    Map(MapNum).Revision = MapRevision
    Map(MapNum).Moral = CByte(MapData(MapIndex + 3))
    Map(MapNum).Up = CInt(MapData(MapIndex + 4))
    Map(MapNum).Down = CInt(MapData(MapIndex + 5))
    Map(MapNum).Left = CInt(MapData(MapIndex + 6))
    Map(MapNum).Right = CInt(MapData(MapIndex + 7))
    Map(MapNum).music = MapData(MapIndex + 8)
    Map(MapNum).BootMap = CInt(MapData(MapIndex + 9))
    Map(MapNum).BootX = CByte(MapData(MapIndex + 10))
    Map(MapNum).BootY = CByte(MapData(MapIndex + 11))
    Map(MapNum).Indoors = CByte(MapData(MapIndex + 12))
    Map(MapNum).Weather = CInt(MapData(MapIndex + 13))

    MapIndex = MapIndex + 14

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(MapNum).Tile(X, Y).Ground = CLng(MapData(MapIndex))
            Map(MapNum).Tile(X, Y).Mask = CLng(MapData(MapIndex + 1))
            Map(MapNum).Tile(X, Y).Anim = CLng(MapData(MapIndex + 2))
            Map(MapNum).Tile(X, Y).Mask2 = CLng(MapData(MapIndex + 3))
            Map(MapNum).Tile(X, Y).M2Anim = CLng(MapData(MapIndex + 4))
            Map(MapNum).Tile(X, Y).Fringe = CLng(MapData(MapIndex + 5))
            Map(MapNum).Tile(X, Y).FAnim = CLng(MapData(MapIndex + 6))
            Map(MapNum).Tile(X, Y).Fringe2 = CLng(MapData(MapIndex + 7))
            Map(MapNum).Tile(X, Y).F2Anim = CLng(MapData(MapIndex + 8))
            Map(MapNum).Tile(X, Y).Type = CByte(MapData(MapIndex + 9))
            Map(MapNum).Tile(X, Y).Data1 = CLng(MapData(MapIndex + 10))
            Map(MapNum).Tile(X, Y).Data2 = CLng(MapData(MapIndex + 11))
            Map(MapNum).Tile(X, Y).Data3 = CLng(MapData(MapIndex + 12))
            Map(MapNum).Tile(X, Y).String1 = MapData(MapIndex + 13)
            Map(MapNum).Tile(X, Y).String2 = MapData(MapIndex + 14)
            Map(MapNum).Tile(X, Y).String3 = MapData(MapIndex + 15)
            Map(MapNum).Tile(X, Y).Light = CLng(MapData(MapIndex + 16))
            Map(MapNum).Tile(X, Y).GroundSet = CByte(MapData(MapIndex + 17))
            Map(MapNum).Tile(X, Y).MaskSet = CByte(MapData(MapIndex + 18))
            Map(MapNum).Tile(X, Y).AnimSet = CByte(MapData(MapIndex + 19))
            Map(MapNum).Tile(X, Y).Mask2Set = CByte(MapData(MapIndex + 20))
            Map(MapNum).Tile(X, Y).M2AnimSet = CByte(MapData(MapIndex + 21))
            Map(MapNum).Tile(X, Y).FringeSet = CByte(MapData(MapIndex + 22))
            Map(MapNum).Tile(X, Y).FAnimSet = CByte(MapData(MapIndex + 23))
            Map(MapNum).Tile(X, Y).Fringe2Set = CByte(MapData(MapIndex + 24))
            Map(MapNum).Tile(X, Y).F2AnimSet = CByte(MapData(MapIndex + 25))
            
            MapIndex = MapIndex + 26
        Next X
    Next Y
            
    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).NPC(X) = CInt(MapData(MapIndex))
        Map(MapNum).SpawnX(X) = CByte(MapData(MapIndex + 1))
        Map(MapNum).SpawnY(X) = CByte(MapData(MapIndex + 2))
        MapIndex = MapIndex + 3
        Call ClearMapNpc(X, MapNum)
    Next X
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            QuestionBlock(MapNum, X, Y).Item1 = CLng(MapData(MapIndex))
            QuestionBlock(MapNum, X, Y).Item2 = CLng(MapData(MapIndex + 1))
            QuestionBlock(MapNum, X, Y).Item3 = CLng(MapData(MapIndex + 2))
            QuestionBlock(MapNum, X, Y).Item4 = CLng(MapData(MapIndex + 3))
            QuestionBlock(MapNum, X, Y).Item5 = CLng(MapData(MapIndex + 4))
            QuestionBlock(MapNum, X, Y).Item6 = CLng(MapData(MapIndex + 5))
            QuestionBlock(MapNum, X, Y).Chance1 = CLng(MapData(MapIndex + 6))
            QuestionBlock(MapNum, X, Y).Chance2 = CLng(MapData(MapIndex + 7))
            QuestionBlock(MapNum, X, Y).Chance3 = CLng(MapData(MapIndex + 8))
            QuestionBlock(MapNum, X, Y).Chance4 = CLng(MapData(MapIndex + 9))
            QuestionBlock(MapNum, X, Y).Chance5 = CLng(MapData(MapIndex + 10))
            QuestionBlock(MapNum, X, Y).Chance6 = CLng(MapData(MapIndex + 11))
            QuestionBlock(MapNum, X, Y).Value1 = CLng(MapData(MapIndex + 12))
            QuestionBlock(MapNum, X, Y).Value2 = CLng(MapData(MapIndex + 13))
            QuestionBlock(MapNum, X, Y).Value3 = CLng(MapData(MapIndex + 14))
            QuestionBlock(MapNum, X, Y).Value4 = CLng(MapData(MapIndex + 15))
            QuestionBlock(MapNum, X, Y).Value5 = CLng(MapData(MapIndex + 16))
            QuestionBlock(MapNum, X, Y).Value6 = CLng(MapData(MapIndex + 17))
            
            ' Save the Question Blocks
                If QuestionBlock(MapNum, X, Y).Item1 > 0 Or QuestionBlock(MapNum, X, Y).Item2 > 0 Or QuestionBlock(MapNum, X, Y).Item3 > 0 Or QuestionBlock(MapNum, X, Y).Item4 > 0 Or QuestionBlock(MapNum, X, Y).Item5 > 0 Or QuestionBlock(MapNum, X, Y).Item6 > 0 Then
                    Call SaveQuestionBlocks(MapNum, X, Y)
                End If
            
            MapIndex = MapIndex + 18
        Next X
    Next Y
    
    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, -1, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next i
    
    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    
    ' Mapper is on the map
    PlayersOnMap(MapNum) = YES

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next i

    ' Refresh map for everyone online
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call SendCheckForMap(i)
        End If
    Next i
End Sub

Public Sub Packet_NeedMap(ByVal Index As Long, ByVal NeedMap As String)
    Dim i As Long

    If NeedMap = "yes" Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)
    Call SendDataTo(Index, SPackets.Smapdone & END_CHAR)

    Player(Index).GettingMap = NO

    Call SendPlayerData(Index)
    Call SendPlayerXY(Index)
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendGuildMemberHP(i)
            Call SendIndexWornEquipment(i)
            Call SendWornEquipment(i)
        End If
    Next i
End Sub

Public Sub Packet_MapGetItem(ByVal Index As Long)
    Call PlayerMapGetItem(Index)
End Sub

Public Sub Packet_MapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim ItemNum As Long
    
    If InvNum < 1 Or InvNum > GetPlayerMaxInv(Index) Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If
    
    ' Stop players from dropping items in minigames, with the exception of Dodgebill
    If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_MINIGAME And GetPlayerMap(Index) <> 188 Then
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, InvNum)
    
    ' Prevent hacking
    If ItemIsStackable(ItemNum) = True Then
        If Amount <= 0 Then
            Call PlayerMsg(Index, "You must drop at least 1 of that item!", BRIGHTRED)
            Exit Sub
        End If

        If Amount > GetPlayerInvItemValue(Index, InvNum) Then
            Call PlayerMsg(Index, "You don't have that many to drop!", BRIGHTRED)
            Exit Sub
        End If
    End If

    ' Prevent hacking
    If Item(ItemNum).Type <> ITEM_TYPE_CURRENCY Then
        If Item(ItemNum).Stackable = 1 Then
            If Amount > GetPlayerInvItemValue(Index, InvNum) Then
                Call HackingAttempt(Index, "Item amount modification")
                Exit Sub
            End If
        End If
    End If

    Call PlayerMapDropItem(Index, InvNum, Amount)

    Call SendStats(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
End Sub

Public Sub Packet_MapRespawn(ByVal Index As Long)
    Dim i As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call RespawnMap(GetPlayerMap(Index))
    Call PlayerMsg(Index, "Map respawned.", BLUE)
End Sub

Public Sub Packet_KickPlayer(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    If GetPlayerAccess(Index) < 1 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> Index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(PlayerIndex) & " has been kicked from Super Mario Bros. Online by " & GetPlayerName(Index) & "!", WHITE)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(PlayerIndex) & ".", ADMIN_LOG)
                Call AlertMsg(PlayerIndex, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "You cannot kick someone with higher or equal access to you!", BRIGHTRED)
            End If
        Else
            Call PlayerMsg(Index, "You cannot kick yourself!", WHITE)
        End If
    Else
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
    End If
End Sub

Public Sub Packet_GetMuteList(ByVal Index As Long)
    If GetPlayerAccess(Index) < 1 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Dim FileName As String, MuteListInfo As String
    Dim i As Integer
    Dim NotifiedOfEntries As Boolean

    If Not FileExists("SMBOMuteList.ini") Then
        Call PlayerMsg(Index, "The mute list is empty!", BRIGHTRED)
        Exit Sub
    End If

    FileName = App.Path & "\SMBOMuteList.ini"
    
    For i = 1 To 100
        MuteListInfo = GetVar(FileName, "Mute List", CStr(i))
        
        If MuteListInfo <> "" Then
            Call PlayerMsg(Index, i & ". " & MuteListInfo, WHITE)
            NotifiedOfEntries = True
        End If
    Next

    If NotifiedOfEntries = False Then
        Call PlayerMsg(Index, "The mute list is empty!", WHITE)
        Exit Sub
    End If
End Sub

Public Sub Packet_MutePlayer(ByVal Index As Long, ByVal Name As String)
    If GetPlayerAccess(Index) < 1 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> Index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(Index) Then
                Call MutePlayer(PlayerIndex, Index)
            Else
                Call PlayerMsg(Index, "You cannot mute someone with higher or equal access to you!", BRIGHTRED)
            End If
        Else
            Call PlayerMsg(Index, "You cannot mute yourself!", WHITE)
        End If
    Else
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
    End If
End Sub

Public Sub Packet_UnmutePlayer(ByVal Index As Long, ByVal Name As String)
    If GetPlayerAccess(Index) < 1 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' If the user entered a username
    If Len(Trim$(Name)) >= 3 Then
        Dim PlayerIndex As Long

        PlayerIndex = FindPlayer(Name)

        If PlayerIndex > 0 Then
            Call UnmutePlayer(0, PlayerIndex, Index)
        Else
            Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        End If
    
        Exit Sub
    End If
    
    ' If the user entered a number
    If IsNumeric(Name) = False Then
        Call PlayerMsg(Index, "You've entered an invalid mute list entry number!", WHITE)
        Exit Sub
    End If
    
    If GetVar(App.Path & "\SMBOMuteList.ini", "Mute List", Name) = "" Then
        ' There were no matches for the number entered, so notify the user of this
        Call PlayerMsg(Index, "The username or mute list entry number you entered is invalid!", BRIGHTRED)
        Exit Sub
    End If
    
    ' Unmute the player by number in the list
    Call UnmutePlayer(CInt(Name), 0, Index)
End Sub

Public Sub Packet_GetBanList(ByVal Index As Long)
    If GetPlayerAccess(Index) < 1 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Dim FileName As String, BanListInfo As String
    Dim i As Integer
    Dim NotifiedOfEntries As Boolean

    If Not FileExists("SMBOBanList.ini") Then
        Call PlayerMsg(Index, "The ban list is empty!", BRIGHTRED)
        Exit Sub
    End If

    FileName = App.Path & "\SMBOBanList.ini"
    
    For i = 1 To 100
        BanListInfo = GetVar(FileName, "Ban List", CStr(i))
        
        If BanListInfo <> "" Then
            Call PlayerMsg(Index, i & ". " & BanListInfo, WHITE)
            NotifiedOfEntries = True
        End If
    Next

    If NotifiedOfEntries = False Then
        Call PlayerMsg(Index, "The ban list is empty!", WHITE)
        Exit Sub
    End If
End Sub

Public Sub Packet_BanPlayer(ByVal Index As Long, ByVal Name As String, Optional ByVal PlayerValidate As Boolean = False)
    If GetPlayerAccess(Index) < 1 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    ' Protect against unauthorized banning
    If PlayerValidate = False Then
        Call HackingAttempt(Index, "Unauthorized banning.")
        Exit Sub
    End If

    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> Index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(Index) Then
                Call BanPlayer(PlayerIndex, Index)
            Else
                Call PlayerMsg(Index, "You cannot ban someone with higher or equal access to you!", BRIGHTRED)
            End If
        Else
            Call PlayerMsg(Index, "You cannot ban yourself!", WHITE)
        End If
    Else
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
    End If
End Sub

Public Sub Packet_UnbanPlayer(ByVal Index As Long, ByVal BanListNum As Integer, Optional ByVal PlayerValidate As Boolean = False)
    If GetPlayerAccess(Index) < 1 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    ' Protect against unauthorized unbanning
    If PlayerValidate = False Then
        Call HackingAttempt(Index, "Unauthorized unbanning.")
        Exit Sub
    End If
    
    If GetVar(App.Path & "\SMBOBanList.ini", "Ban List", CStr(BanListNum)) = "" Then
        ' There were no matches for the number entered, so notify the user of this
        Call PlayerMsg(Index, "The ban list entry number you entered is invalid!", BRIGHTRED)
        Exit Sub
    End If
                
    ' Unban the player by number in the list
    Call UnbanPlayer(Index, BanListNum)
End Sub

Public Sub Packet_RequestEditMap(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Seditmap & END_CHAR)
End Sub

Public Sub Packet_RequestEditItem(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Sitemeditor & END_CHAR)
End Sub

Public Sub Packet_EditItem(ByVal Index As Long, ByVal ItemNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If

    Call SendEditItemTo(Index, ItemNum)

    Call AddLog(GetPlayerName(Index) & " editing item #" & ItemNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveItem(ByVal Index As Long, ByRef ItemData() As String)
    Dim ItemNum As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    ItemNum = CLng(ItemData(1))

    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If

    Item(ItemNum).Name = ItemData(2)
    Item(ItemNum).Pic = CLng(ItemData(3))
    Item(ItemNum).Type = CByte(ItemData(4))
    Item(ItemNum).Data1 = CLng(ItemData(5))
    Item(ItemNum).Data2 = CLng(ItemData(6))
    Item(ItemNum).Data3 = CLng(ItemData(7))
    Item(ItemNum).StrReq = CLng(ItemData(8))
    Item(ItemNum).DefReq = CLng(ItemData(9))
    Item(ItemNum).SpeedReq = CLng(ItemData(10))
    Item(ItemNum).MagicReq = CLng(ItemData(11))
    Item(ItemNum).ClassReq = CLng(ItemData(12))
    Item(ItemNum).AccessReq = CByte(ItemData(13))

    Item(ItemNum).addHP = CLng(ItemData(14))
    Item(ItemNum).addMP = CLng(ItemData(15))
    Item(ItemNum).addSP = CLng(ItemData(16))
    Item(ItemNum).AddStr = CLng(ItemData(17))
    Item(ItemNum).AddDef = CLng(ItemData(18))
    Item(ItemNum).AddSpeed = CLng(ItemData(19))
    Item(ItemNum).AddMagi = CLng(ItemData(20))
    Item(ItemNum).AddEXP = CLng(ItemData(21))
    Item(ItemNum).Desc = ItemData(22)
    Item(ItemNum).AttackSpeed = CLng(ItemData(23))
    Item(ItemNum).Price = CLng(ItemData(24))
    Item(ItemNum).Stackable = CByte(ItemData(25))
    Item(ItemNum).Bound = CByte(ItemData(26))
    Item(ItemNum).LevelReq = CLng(ItemData(27))
    Item(ItemNum).HPReq = CLng(ItemData(28))
    Item(ItemNum).FPReq = CLng(ItemData(29))
    Item(ItemNum).Ammo = CLng(ItemData(30))
    Item(ItemNum).AddCritChance = CDbl(ItemData(31))
    Item(ItemNum).AddBlockChance = CDbl(ItemData(32))
    Item(ItemNum).Cookable = CBool(ItemData(33))

    Call SendUpdateItemToAll(ItemNum)
    Call SaveItem(ItemNum)

    Call AddLog(GetPlayerName(Index) & " saved item #" & ItemNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditNPC(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Snpceditor & END_CHAR)
End Sub

Public Sub Packet_EditNPC(ByVal Index As Long, ByVal NpcNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid NPC Index")
        Exit Sub
    End If

    Call SendEditNpcTo(Index, NpcNum)

    Call AddLog(GetPlayerName(Index) & " editing npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveNPC(ByVal Index As Long, ByRef NPCData() As String)
    Dim i As Long, NpcNum As Long, NPCIndex As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    NpcNum = CLng(NPCData(1))

    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid NPC Index")
        Exit Sub
    End If

    NPC(NpcNum).Name = NPCData(2)
    NPC(NpcNum).AttackSay = NPCData(3)
    NPC(NpcNum).Sprite = CLng(NPCData(4))
    NPC(NpcNum).SpawnSecs = CLng(NPCData(5))
    NPC(NpcNum).Behavior = CByte(NPCData(6))
    NPC(NpcNum).Range = CByte(NPCData(7))
    NPC(NpcNum).STR = CLng(NPCData(8))
    NPC(NpcNum).DEF = CLng(NPCData(9))
    NPC(NpcNum).Speed = CLng(NPCData(10))
    NPC(NpcNum).Magi = CLng(NPCData(11))
    NPC(NpcNum).Big = CLng(NPCData(12))
    NPC(NpcNum).MAXHP = CLng(NPCData(13))
    NPC(NpcNum).Exp = CLng(NPCData(14))
    NPC(NpcNum).SpawnTime = CLng(NPCData(15))
    NPC(NpcNum).Element = CLng(NPCData(16))
    NPC(NpcNum).SPRITESIZE = CByte(NPCData(17))

    NPCIndex = 18

    For i = 1 To MAX_NPC_DROPS
        NPC(NpcNum).ItemNPC(i).Chance = CLng(NPCData(NPCIndex))
        NPC(NpcNum).ItemNPC(i).ItemNum = CLng(NPCData(NPCIndex + 1))
        NPC(NpcNum).ItemNPC(i).ItemValue = CLng(NPCData(NPCIndex + 2))
        NPCIndex = NPCIndex + 3
    Next i

    NPC(NpcNum).AttackSay2 = NPCData(48)
    NPC(NpcNum).LEVEL = CLng(NPCData(49))

    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)

    Call AddLog(GetPlayerName(Index) & " saved npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditShop(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Sshopeditor & END_CHAR)
End Sub

Public Sub Packet_EditShop(ByVal Index As Long, ByVal ShopNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If

    Call SendEditShopTo(Index, ShopNum)

    Call AddLog(GetPlayerName(Index) & " editing shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveShop(ByVal Index As Long, ByRef ShopData() As String)
    Dim i As Long, ShopNum As Long, ShopIndex As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    ShopNum = CLng(ShopData(1))

    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If

    Shop(ShopNum).Name = ShopData(2)
    Shop(ShopNum).FixesItems = 0
    Shop(ShopNum).BuysItems = CByte(ShopData(3))
    Shop(ShopNum).ShowInfo = CByte(ShopData(4))
    Shop(ShopNum).CurrencyItem = CInt(ShopData(5))

    ShopIndex = 6

    For i = 1 To MAX_SHOP_ITEMS
        Shop(ShopNum).ShopItem(i).ItemNum = CLng(ShopData(ShopIndex))
        Shop(ShopNum).ShopItem(i).Amount = CDbl(ShopData(ShopIndex + 1))
        Shop(ShopNum).ShopItem(i).Price = CDbl(ShopData(ShopIndex + 2))
        Shop(ShopNum).ShopItem(i).CurrencyItem = CInt(ShopData(ShopIndex + 3))
        ShopIndex = ShopIndex + 4
    Next i

    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)

    Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditSpell(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Sspelleditor & END_CHAR)
End Sub

Public Sub Packet_EditSpell(ByVal Index As Long, ByVal SpellNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If

    Call SendEditSpellTo(Index, SpellNum)

    Call AddLog(GetPlayerName(Index) & " editing spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveSpell(ByVal Index As Long, ByRef SpellData() As String)
    Dim SpellNum As Long
    
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    SpellNum = CLng(SpellData(1))

    If SpellNum < 1 Or SpellNum > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If

    Spell(SpellNum).Name = SpellData(2)
    Spell(SpellNum).ClassReq = CLng(SpellData(3))
    Spell(SpellNum).LevelReq = CLng(SpellData(4))
    Spell(SpellNum).Type = CLng(SpellData(5))
    Spell(SpellNum).Data1 = CLng(SpellData(6))
    Spell(SpellNum).Data2 = CLng(SpellData(7))
    Spell(SpellNum).Data3 = CLng(SpellData(8))
    Spell(SpellNum).MPCost = CLng(SpellData(9))
    Spell(SpellNum).Sound = SpellData(10)
    Spell(SpellNum).Range = CByte(SpellData(11))
    Spell(SpellNum).SpellAnim = CLng(SpellData(12))
    Spell(SpellNum).SpellTime = CLng(SpellData(13))
    Spell(SpellNum).SpellDone = CLng(SpellData(14))
    Spell(SpellNum).AE = CLng(SpellData(15))
    Spell(SpellNum).Big = CLng(SpellData(16))
    Spell(SpellNum).Element = CLng(SpellData(17))
    Spell(SpellNum).Stat = CInt(SpellData(18))
    Spell(SpellNum).StatTime = CLng(SpellData(19))
    Spell(SpellNum).Multiplier = CDbl(SpellData(20))
    Spell(SpellNum).PassiveStat = CInt(SpellData(21))
    Spell(SpellNum).PassiveStatChange = CInt(SpellData(22))
    Spell(SpellNum).UsePassiveStat = CBool(SpellData(23))
    Spell(SpellNum).SelfSpell = CBool(SpellData(24))

    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)

    Call AddLog(GetPlayerName(Index) & " saving spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_ForgetSpell(ByVal Index As Long, ByVal SpellNum As Long)
    If SpellNum < 1 Or SpellNum > MAX_PLAYER_SPELLS Then
        Call HackingAttempt(Index, "Invalid Special Attack Slot")
        Exit Sub
    End If

    With Player(Index).Char(Player(Index).CharNum)
        If .Spell(SpellNum) = 0 Then
            Call PlayerMsg(Index, "There's no special attack here!", WHITE)
        Else
            Call PlayerMsg(Index, "You have forgotten the special attack, " & Trim$(Spell(.Spell(SpellNum)).Name) & "!", WHITE)

            .Spell(SpellNum) = 0

            Call SendSpells(Index)
        End If
    End With
End Sub

Public Sub Packet_RequestEditRecipe(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Srecipeeditor & END_CHAR)
End Sub

Public Sub Packet_EditRecipe(ByVal Index As Long, ByVal RecipeNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If RecipeNum < 0 Or RecipeNum > MAX_RECIPES Then
        Call HackingAttempt(Index, "Invalid Recipe Index")
        Exit Sub
    End If

    Call SendEditRecipeTo(Index, RecipeNum)

    Call AddLog(GetPlayerName(Index) & " editing recipe #" & RecipeNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveRecipe(ByVal Index As Long, ByRef RecipeData() As String)
    Dim RecipeNum As Long
    
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    RecipeNum = CLng(RecipeData(1))

    If RecipeNum < 1 Or RecipeNum > MAX_RECIPES Then
        Call HackingAttempt(Index, "Invalid Recipe Index")
        Exit Sub
    End If

    Recipe(RecipeNum).Ingredient1 = CLng(RecipeData(2))
    Recipe(RecipeNum).Ingredient2 = CLng(RecipeData(3))
    Recipe(RecipeNum).ResultItem = CLng(RecipeData(4))
    Recipe(RecipeNum).Name = RecipeData(5)

    Call SendUpdateRecipeToAll(RecipeNum)
    Call SaveRecipe(RecipeNum)

    Call AddLog(GetPlayerName(Index) & " saving recipe #" & RecipeNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SetAccess(ByVal Index As Long, ByVal Name As String, ByVal AccessLvl As Byte)
    If GetPlayerAccess(Index) <> 5 Then
        Call PlayerMsg(Index, "You do not have the authority to change another player's access!", BRIGHTRED)
        Exit Sub
    End If
    
    If AccessLvl < 0 Or AccessLvl > 5 Then
        Call PlayerMsg(Index, "You have entered an invalid access level.", BRIGHTRED)
        Exit Sub
    End If
    
    Dim PlayerIndex As Long
    
    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If GetPlayerName(Index) <> GetPlayerName(PlayerIndex) Then
            If GetPlayerAccess(Index) > GetPlayerAccess(PlayerIndex) Then
                If GetPlayerAccess(PlayerIndex) > AccessLvl Then
                    Call PlayerMsg(PlayerIndex, "You have been demoted.", AdminColor)
                ElseIf GetPlayerAccess(PlayerIndex) < AccessLvl Then
                    Call PlayerMsg(PlayerIndex, "You have been promoted.", AdminColor)
                End If
            
                Call SetPlayerAccess(PlayerIndex, AccessLvl)
                Call SendPlayerData(PlayerIndex)
    
                Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(PlayerIndex) & "'s access.", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "You cannot change someone's rank if his/her rank is equal to or higher than yours!", BLACK)
            End If
        Else
            Call PlayerMsg(Index, "You can't change your own access.", RED)
        End If
    Else
        Call PlayerMsg(Index, "Player is not online.", WHITE)
    End If
End Sub

Public Sub Packet_OnlineList(ByVal Index As Long)
    Call SendOnlineList
End Sub

Public Sub Packet_SetMOTD(ByVal Index As Long, ByVal MOTD As String)
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call PutVar(App.Path & "\MOTD.ini", "MOTD", "Msg", MOTD)
            
    Call GlobalMsg("MOTD changed to: " & MOTD, BRIGHTCYAN)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & MOTD, ADMIN_LOG)
End Sub

Public Sub Packet_BuyItem(ByVal Index As Long, ByVal ShopIndex As Long, ByVal ItemIndex As Long, ByVal ItemAmount As Long)
    Dim InvItem As Long, ItemNum As Long, Stache As Long
    Dim ItemPrice As Currency
    
    ItemPrice = Shop(ShopIndex).ShopItem(ItemIndex).Price
    ItemNum = Shop(ShopIndex).ShopItem(ItemIndex).ItemNum
    Stache = GetPlayerStache(Index)
    
    ' Prevents players from buying the Experience Point Plus badge in the WAM lobby
    If ItemNum = 57 Then
        If GetVar(App.Path & "\Scripts\" & "WhackFame.ini", GetPlayerName(Index), "In") <> "Yes" Then
            Call PlayerMsg(Index, "You have not earned enough points in Whack-A-Monty to purchase this item!", WHITE)
            Exit Sub
        End If
    End If
    ' Prevents players from buying Super Shrooms until they complete the Favor in Kinopio Village
    If ItemNum = 40 Then
        If GetVar(App.Path & "\Scripts\" & "Quests.ini", GetPlayerName(Index), "ItemQuest6") <> "Done" Then
            Call PlayerMsg(Index, "Shop Owner: Sorry, but we're out of stock on Super Shrooms.", WHITE)
            Exit Sub
        End If
    End If
    
    ' Exclude Daily Event Shop
    If ShopIndex <> 14 And ShopIndex <> 19 And ShopIndex <> 21 And ShopIndex <> 23 Then
        If Stache <= 30 Then
            ItemPrice = Int(ItemPrice * ((100 - Stache) / 100))
        Else
            ItemPrice = Int(ItemPrice * 0.7)
        End If
    End If
        
    ItemPrice = Int(ItemPrice * ItemAmount)
    
    If ShopIndex < 1 Or ShopIndex > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If
    
    If ItemIndex < 1 Or ItemIndex > MAX_SHOP_ITEMS Then
        Call HackingAttempt(Index, "Invalid Shop Item")
        Exit Sub
    End If

    ' Check to see if player's inventory is full
    If ItemIsStackable(ItemNum) = True Then
        InvItem = FindOpenInvSlot(Index, ItemNum)
            
        If InvItem = 0 Then
            Call PlayerMsg(Index, "Your inventory is full!", BRIGHTRED)
            Exit Sub
        End If
    Else
        InvItem = GetFreeSlots(Index)
        
        If ItemAmount > InvItem Then
            Call PlayerMsg(Index, "Your inventory is full!", BRIGHTRED)
            Exit Sub
        End If
    End If
    
    InvItem = 1
    
    ' Adjust how many of the item you receive when buying stackables
    If ItemIsStackable(ItemNum) Then
        ItemAmount = (ItemAmount * Shop(ShopIndex).ShopItem(ItemIndex).Amount)
    End If
    
    ' Check to see if they have enough currency
    If HasItem(Index, Shop(ShopIndex).ShopItem(ItemIndex).CurrencyItem) >= ItemPrice Then
        Call TakeItem(Index, Shop(ShopIndex).ShopItem(ItemIndex).CurrencyItem, ItemPrice)
        Call GiveItem(Index, ItemNum, ItemAmount)
        
        Dim ItemName As String
                
        ItemName = Trim$(Item(ItemNum).Name)
        
        If Right$(ItemName, 1) = "s" Then
            ItemName = Mid$(ItemName, 1, Len(ItemName) - 1)
        End If
        
        If ItemAmount > 1 Then
            Call PlayerMsg(Index, "You bought " & ItemAmount & " " & ItemName & "s!", YELLOW)
        Else
            If FindItemVowels(ItemNum) = True Then
                Call PlayerMsg(Index, "You bought an " & Trim$(Item(ItemNum).Name) & "!", YELLOW)
            Else
                Call PlayerMsg(Index, "You bought a " & Trim$(Item(ItemNum).Name) & "!", YELLOW)
            End If
        End If
    Else
        Call PlayerMsg(Index, "You cannot afford that item!", RED)
    End If
End Sub

Public Sub Packet_SellItem(ByVal Index As Long, ByVal ShopNum As Long, ByVal ItemNum As Long, ByVal ItemSlot As Long, ByVal ItemAmt As Long)
    Dim i As Integer, p As Integer, Q As Integer, InvAmount As Integer
    Dim ItemPrice As Currency
    Dim Stache As Long
  
    ItemPrice = Item(ItemNum).Price
    Stache = GetPlayerStache(Index)
    InvAmount = 1

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        Call PlayerMsg(Index, "You cannot sell currency.", RED)
        Exit Sub
    End If

    If ItemIsStackable(ItemNum) = True Then
        If ItemAmt > GetPlayerInvItemValue(Index, ItemSlot) Then
            Call PlayerMsg(Index, "You don't have enough of that item to sell that many!", RED)
            Exit Sub
        End If
    Else
        If ItemAmt > CountInvItemNum(Index, ItemNum) Then
            Call PlayerMsg(Index, "You don't have enough of that item to sell that many!", RED)
            Exit Sub
        End If
    End If
    
    ' Exclude Daily Event shop and calculate the sell price bonus from Stache
    If ShopNum <> 14 And ShopNum <> 19 And ShopNum <> 21 And ShopNum <> 23 Then
        If Stache <= 30 Then
            ItemPrice = Int(ItemPrice * ((100 + Stache) / 100))
        Else
            ItemPrice = Int(ItemPrice * 1.3)
        End If
    End If
    
    Dim ShouldExit As Boolean
    
    For Q = 1 To MAX_SHOPS
        For i = 1 To MAX_SHOP_ITEMS
            If Shop(Q).ShopItem(i).ItemNum = ItemNum Then
                p = Shop(Q).ShopItem(i).CurrencyItem
                
                ShouldExit = True
                Exit For
            End If
        Next i
        
        If ShouldExit = True Then
            Exit For
        End If
    Next Q
    
    If p <= 0 Then
        ' Give Coins for Mushroom Kingdom items and Beanbean Coins for Beanbean Kingdom items
        If ItemNum < 272 Then
            p = 1
        Else
            p = 271
        End If
    End If
    
    If Item(ItemNum).Price > 0 Then
        Call TakeItem(Index, ItemNum, ItemAmt)
        Call GiveItem(Index, p, ItemPrice * ItemAmt)
        
        Dim ItemName As String
                
        ItemName = Trim$(Item(ItemNum).Name)
        
        If Right$(ItemName, 1) = "s" Then
            ItemName = Mid$(ItemName, 1, Len(ItemName) - 1)
        End If

        If ItemAmt > 1 Then
            Call PlayerMsg(Index, "You sold " & ItemAmt & " " & ItemName & "s for " & (ItemPrice * ItemAmt) & " " & Trim$(Item(p).Name) & "s!", YELLOW)
        Else
            If FindItemVowels(ItemNum) = True Then
                If (ItemPrice * ItemAmt) > 1 Then
                    Call PlayerMsg(Index, "You sold an " & Trim$(Item(ItemNum).Name) & " for " & (ItemPrice * ItemAmt) & " " & Trim$(Item(p).Name) & "s!", YELLOW)
                Else
                    Call PlayerMsg(Index, "You sold an " & Trim$(Item(ItemNum).Name) & " for " & (ItemPrice * ItemAmt) & " " & Trim$(Item(p).Name) & "!", YELLOW)
                End If
            Else
                If (ItemPrice * ItemAmt) > 1 Then
                    Call PlayerMsg(Index, "You sold a " & Trim$(Item(ItemNum).Name) & " for " & (ItemPrice * ItemAmt) & " " & Trim$(Item(p).Name) & "s!", YELLOW)
                Else
                    Call PlayerMsg(Index, "You sold a " & Trim$(Item(ItemNum).Name) & " for " & (ItemPrice * ItemAmt) & " " & Trim$(Item(p).Name) & "!", YELLOW)
                End If
            End If
        End If
    Else
        Call PlayerMsg(Index, "This item cannot be sold.", RED)
    End If
End Sub

Public Sub Packet_Search(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long

    If X < 0 Or X > MAX_MAPX Then
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Exit Sub
    End If
    
    If GetPlayerInBattle(Index) = True Then
        Exit Sub
    End If
    
    ' Check for a player
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        If GetPlayerInBattle(i) = False Then
                            ' Change the target
                            Player(Index).Target = i
                            Player(Index).TargetType = TARGET_TYPE_PLAYER

                            Call PlayerMsg(Index, "Your target is now " & GetPlayerName(i) & ".", YELLOW)

                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' Check for an NPC
    For i = 1 To MAX_MAP_NPCS
        If MapNPC(GetPlayerMap(Index), i).num > 0 Then
            If MapNPC(GetPlayerMap(Index), i).X = X Then
                If MapNPC(GetPlayerMap(Index), i).Y = Y Then
                    If MapNPC(GetPlayerMap(Index), i).InBattle = False Then
                        ' Change the target
                        Player(Index).TargetNPC = i
                        Player(Index).TargetType = TARGET_TYPE_NPC

                        Call PlayerMsg(Index, "Your target is now a " & Trim$(NPC(MapNPC(GetPlayerMap(Index), i).num).Name) & ".", YELLOW)

                        Exit Sub
                    End If
                End If
            End If
        End If
    Next i

    ' Check for an item on the ground.
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(GetPlayerMap(Index), i).num > 0 Then
            If MapItem(GetPlayerMap(Index), i).X = X Then
                If MapItem(GetPlayerMap(Index), i).Y = Y Then
                    Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), i).num).Name) & ".", YELLOW)
                    Exit Sub
                End If
            End If
        End If
    Next i
End Sub

Public Sub Packet_PlayerChat(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
        Exit Sub
    End If

    If PlayerIndex = Index Then
        Call PlayerMsg(Index, "You cannot chat with yourself.", PINK)
        Exit Sub
    End If

    If Player(Index).InChat = 1 Then
        Call PlayerMsg(Index, "You're already in a chat with another player!", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).InChat = 1 Then
        Call PlayerMsg(Index, Name & " is already in a chat with another player!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(Index, "Chat request has been sent to " & GetPlayerName(PlayerIndex) & ".", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " wants you to chat with them. Type /chat to accept, or /chatdecline to decline.", PINK)

    Player(Index).ChatPlayer = PlayerIndex
    Player(PlayerIndex).ChatPlayer = Index
End Sub

Public Sub Packet_AcceptChat(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(Index).ChatPlayer <> PlayerIndex Then
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Sppchatting & SEP_CHAR & GetPlayerName(PlayerIndex) & END_CHAR)
    Call SendDataTo(PlayerIndex, SPackets.Sppchatting & SEP_CHAR & GetPlayerName(Index) & END_CHAR)
End Sub

Public Sub Packet_DenyChat(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(Index).ChatPlayer <> PlayerIndex Then
        Call PlayerMsg(Index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(Index, "Declined chat request.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " declined your request.", PINK)

    Player(Index).ChatPlayer = 0
    Player(Index).InChat = 0

    Player(PlayerIndex).ChatPlayer = 0
    Player(PlayerIndex).InChat = 0
End Sub

Public Sub Packet_QuitChat(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).ChatPlayer

    If PlayerIndex > 0 Then
        Call SendDataTo(Index, SPackets.Sqchat & END_CHAR)
        Call SendDataTo(PlayerIndex, SPackets.Sqchat & END_CHAR)
    
        Player(Index).ChatPlayer = 0
        Player(Index).InChat = 0
    
        Player(PlayerIndex).ChatPlayer = 0
        Player(PlayerIndex).InChat = 0
    End If
End Sub

Public Sub Packet_SendChat(ByVal Index As Long, ByVal Message As String)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> Index Then
        Call PlayerMsg(Index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call SendDataTo(PlayerIndex, SPackets.Ssendchat & SEP_CHAR & Message & SEP_CHAR & GetPlayerName(Index) & END_CHAR)
End Sub

Public Sub Packet_TradeRequest(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long
    
    PlayerIndex = FindPlayer(Name)
    
    ' Player is offline
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    ' Player is yourself
    If PlayerIndex = Index Then
        Call PlayerMsg(Index, "You cannot trade with yourself!", WHITE)
        Exit Sub
    End If

    ' Player is on a different map
    If GetPlayerMap(Index) <> GetPlayerMap(PlayerIndex) Then
        Call PlayerMsg(Index, "You must be on the same map to trade with " & GetPlayerName(PlayerIndex) & "!", WHITE)
        Exit Sub
    End If

    ' You are already in a trade
    If Player(Index).InTrade Then
        Call PlayerMsg(Index, "You're already in a trade with someone else!", WHITE)
        Exit Sub
    End If

    ' Player is already in a trade
    If Player(PlayerIndex).InTrade Then
        Call PlayerMsg(Index, Name & " is already in a trade!", WHITE)
        Exit Sub
    End If
    
    ' Player has already been offered a trade
    If Player(PlayerIndex).TradePlayer <> 0 Then
        Call PlayerMsg(Index, Name & " has already been offered a trade request!", WHITE)
        Exit Sub
    End If
    
    ' Requires you to be at most 3 spaces away to be able to trade with another player
    If GetPlayerY(Index) = GetPlayerY(PlayerIndex) Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) + 1 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) + 2 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) + 3 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) - 1 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) - 2 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) - 3 Then
        If GetPlayerX(Index) = GetPlayerX(PlayerIndex) And GetPlayerY(Index) < GetPlayerY(PlayerIndex) + 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) + 1 And GetPlayerY(Index) < GetPlayerY(PlayerIndex) + 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) + 2 And GetPlayerY(Index) < GetPlayerY(PlayerIndex) + 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) + 3 And GetPlayerY(Index) > GetPlayerY(PlayerIndex) - 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) - 1 And GetPlayerY(Index) > GetPlayerY(PlayerIndex) - 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) - 2 And GetPlayerY(Index) > GetPlayerY(PlayerIndex) - 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) - 3 And GetPlayerY(Index) > GetPlayerY(PlayerIndex) - 4 Then
            If GetPlayerInBattle(Index) = True Then
                Call PlayerMsg(Index, "You cannot trade with another player while you're in battle!", WHITE)
                Exit Sub
            End If
            If GetPlayerInBattle(PlayerIndex) = True Then
                Call PlayerMsg(Index, "You cannot trade with a player that is in a battle!", WHITE)
                Exit Sub
            End If
              
            ' Set that the players are trading with each other
            Player(Index).TradePlayer = PlayerIndex
            Player(PlayerIndex).TradePlayer = Index
            
            Call SendDataTo(PlayerIndex, SPackets.Straderequest & SEP_CHAR & GetPlayerName(Index) & END_CHAR)
        Else
            Call PlayerMsg(Index, "You must be at most 3 spaces away from the other player to trade!", WHITE)
        End If
    Else
        Call PlayerMsg(Index, "You must be at most 3 spaces away from the other player to trade!", WHITE)
    End If
End Sub

Public Sub Packet_AcceptTrade(ByVal Index As Long)
    Dim PlayerIndex As Long
    Dim i As Long

    PlayerIndex = Player(Index).TradePlayer
    
    ' Player is not online
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to trade with anyone!", WHITE)
        Exit Sub
    End If

    ' Player is not on the same map
    If GetPlayerMap(Index) <> GetPlayerMap(PlayerIndex) Then
        Call PlayerMsg(Index, "You must be on the same map to trade with " & GetPlayerName(PlayerIndex) & "!", WHITE)
        Exit Sub
    End If
    
    ' Requires you to be at most 3 spaces away to be able to trade with another player
    If GetPlayerY(Index) = GetPlayerY(PlayerIndex) Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) + 1 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) + 2 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) + 3 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) - 1 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) - 2 Or GetPlayerY(Index) = GetPlayerY(PlayerIndex) - 3 Then
        If GetPlayerX(Index) = GetPlayerX(PlayerIndex) And GetPlayerY(Index) < GetPlayerY(PlayerIndex) + 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) + 1 And GetPlayerY(Index) < GetPlayerY(PlayerIndex) + 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) + 2 And GetPlayerY(Index) < GetPlayerY(PlayerIndex) + 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) + 3 And GetPlayerY(Index) > GetPlayerY(PlayerIndex) - 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) - 1 And GetPlayerY(Index) > GetPlayerY(PlayerIndex) - 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) - 2 And GetPlayerY(Index) > GetPlayerY(PlayerIndex) - 4 Or GetPlayerX(Index) = GetPlayerX(PlayerIndex) - 3 And GetPlayerY(Index) > GetPlayerY(PlayerIndex) - 4 Then
        Else
            Player(Index).TradePlayer = 0
            Player(PlayerIndex).TradePlayer = 0
            
            Call PlayerMsg(Index, "You must be at most 3 spaces away from the other player to trade!", WHITE)
            Exit Sub
        End If
    Else
        Player(Index).TradePlayer = 0
        Player(PlayerIndex).TradePlayer = 0
        
        Call PlayerMsg(Index, "You must be at most 3 spaces away from the other player to trade!", WHITE)
        Exit Sub
    End If
    
    ' Clear out any previous trade settings
    For i = 1 To MAX_PLAYER_TRADES
        Player(Index).Trades(i).InvNum = 0
        Player(Index).Trades(i).InvName = vbNullString
        Player(Index).Trades(i).InvVal = 0

        Player(PlayerIndex).Trades(i).InvNum = 0
        Player(PlayerIndex).Trades(i).InvName = vbNullString
        Player(PlayerIndex).Trades(i).InvVal = 0
    Next i

    ' Set that the players are trading
    Player(Index).InTrade = True
    Player(PlayerIndex).InTrade = True
    
    ' Notify the players that they will begin trading
    Call PlayerMsg(Index, "You are trading with " & GetPlayerName(PlayerIndex) & "!", WHITE)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " accepted your trade request!", WHITE)

    ' Start the trade
    Call SendDataTo(Index, SPackets.Sstarttrade & SEP_CHAR & GetPlayerName(PlayerIndex) & END_CHAR)
    Call SendDataTo(PlayerIndex, SPackets.Sstarttrade & SEP_CHAR & GetPlayerName(Index) & END_CHAR)
End Sub

Public Sub Packet_DeclineTrade(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).TradePlayer
    
    ' Player is not online
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to trade with anyone!", WHITE)
        Exit Sub
    End If

    Call PlayerMsg(Index, "You declined " & GetPlayerName(PlayerIndex) & "'s trade request.", WHITE)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " declined your trade request.", WHITE)

    Player(Index).TradePlayer = 0
    Player(Index).InTrade = False

    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False
End Sub

Public Sub Packet_StopTrading(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).TradePlayer

    ' Player is not online
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to trade with anyone!", WHITE)
        Exit Sub
    End If

    Call PlayerMsg(Index, "You stopped trading.", WHITE)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " stopped trading with you!", WHITE)

    Player(Index).TradeOk = 0
    Player(Index).TradePlayer = 0
    Player(Index).InTrade = False

    Player(PlayerIndex).TradeOk = 0
    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False

    Call SendDataTo(Index, SPackets.Sstoptrading & END_CHAR)
    Call SendDataTo(PlayerIndex, SPackets.Sstoptrading & END_CHAR)
End Sub

Public Sub Packet_UpdateTradeOffers(ByVal Index As Long, ByVal TradeOfferNum As Long, ByVal ItemName As String, ByVal InvNum As Long, ByVal Amount As Long)
    Dim PlayerIndex As Long, ItemNum As Long
    
    PlayerIndex = Player(Index).TradePlayer
    
    ' Set values
    Player(Index).Trades(TradeOfferNum).InvName = ItemName
    Player(Index).Trades(TradeOfferNum).InvNum = InvNum
    Player(Index).Trades(TradeOfferNum).InvVal = Amount
 
    If InvNum > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, InvNum)
    End If
    
    ' Set that both players didn't accept the trade since changes were made
    Player(Index).TradeOk = 0
    Player(PlayerIndex).TradeOk = 0
    
    ' Send the trade update
    Call SendDataTo(PlayerIndex, SPackets.Supdatetradeoffers & SEP_CHAR & TradeOfferNum & SEP_CHAR & InvNum & SEP_CHAR & ItemName & SEP_CHAR & Amount & SEP_CHAR & ItemNum & END_CHAR)
    
    ' Update the trade message
    Call SendDataTo(Index, SPackets.Strademessage & SEP_CHAR & Index & SEP_CHAR & PlayerIndex & SEP_CHAR & 0 & END_CHAR)
    Call SendDataTo(PlayerIndex, SPackets.Strademessage & SEP_CHAR & Index & SEP_CHAR & PlayerIndex & SEP_CHAR & 0 & END_CHAR)
End Sub

Public Sub Packet_CompleteTrade(ByVal Index As Long)
    Dim i As Long, PlayerIndex As Long, ItemNum As Long, PlayerStackableItem(1 To MAX_PLAYER_TRADES) As Long, OtherStackableItem(1 To MAX_PLAYER_TRADES) As Long, Items(1 To MAX_PLAYER_TRADES) As Long, OtherItems(1 To MAX_PLAYER_TRADES) As Long
    Dim PlayerInvSlots As Integer, OtherPlayerInvSlots As Integer, PlayerTradeSlots As Integer, OtherPlayerTradeSlots As Integer, PlayerTotalTradeSlots As Integer, OtherPlayerTotalTradeSlots As Integer
    
    PlayerIndex = Player(Index).TradePlayer
    
    If Player(Index).TradeOk = 0 Then
        Player(Index).TradeOk = 1
        Call SendDataTo(PlayerIndex, SPackets.Strademessage & SEP_CHAR & PlayerIndex & SEP_CHAR & Index & SEP_CHAR & 1 & END_CHAR)
        Call SendDataTo(Index, SPackets.Strademessage & SEP_CHAR & PlayerIndex & SEP_CHAR & Index & SEP_CHAR & 1 & END_CHAR)
    End If
    
    ' Start swapping items when both players have confirmed the trade
    If Player(Index).TradeOk = 1 And Player(PlayerIndex).TradeOk = 1 Then
        ' Find out how many inventory slots you can hold
        PlayerInvSlots = GetFreeSlots(Index)
        
        ' Store the stackable items in the trade
        For i = 1 To MAX_PLAYER_TRADES
            If Player(Index).Trades(i).InvNum > 0 Then
                PlayerTotalTradeSlots = PlayerTotalTradeSlots + 1
                ItemNum = GetPlayerInvItemNum(Index, Player(Index).Trades(i).InvNum)
                
                If ItemNum > 0 Then
                    If ItemIsStackable(ItemNum) = True Then
                        PlayerStackableItem(i) = ItemNum
                    Else
                        ' Store how many unstackable items you're trading
                        PlayerTradeSlots = PlayerTradeSlots + 1
                    End If
                End If
            End If
        Next i
        
        ' Find out how many inventory slots the other player can hold
        OtherPlayerInvSlots = GetFreeSlots(PlayerIndex)
        
        ' Store the stackable items in the trade
        For i = 1 To MAX_PLAYER_TRADES
            If Player(PlayerIndex).Trades(i).InvNum > 0 Then
                OtherPlayerTotalTradeSlots = OtherPlayerTotalTradeSlots + 1
                ItemNum = GetPlayerInvItemNum(PlayerIndex, Player(PlayerIndex).Trades(i).InvNum)
                
                If ItemNum > 0 Then
                    If ItemIsStackable(ItemNum) = True Then
                        OtherStackableItem(i) = ItemNum
                    Else
                        ' Store how many unstackable items the other player is trading
                        OtherPlayerTradeSlots = OtherPlayerTradeSlots + 1
                    End If
                End If
            End If
        Next i
        
        ' Add traded stackables to players' inventory capacity
        For i = 1 To MAX_PLAYER_TRADES
            ' Add to your inventory capacity
            If OtherStackableItem(i) > 0 Then
                ' Check if you have the item
                If FindOpenInvSlot(Index, OtherStackableItem(i)) > 0 Then
                    PlayerInvSlots = PlayerInvSlots + 1
                End If
            End If
            ' Add to other players' inventory capacity
            If PlayerStackableItem(i) > 0 Then
                ' Check if the other player has the item
                If FindOpenInvSlot(PlayerIndex, PlayerStackableItem(i)) > 0 Then
                    OtherPlayerInvSlots = OtherPlayerInvSlots + 1
                End If
            End If
        Next i
        
        ' Since we're taking away the items before we trade them, add the number of non-stackable items to each players' inventory capacity
        PlayerInvSlots = PlayerInvSlots + PlayerTradeSlots
        OtherPlayerInvSlots = OtherPlayerInvSlots + OtherPlayerTradeSlots
        
        ' Only trade if both players can hold enough items for the trade
        If PlayerInvSlots >= OtherPlayerTotalTradeSlots And OtherPlayerInvSlots >= PlayerTotalTradeSlots Then
            ' Take away items from both players
            For i = 1 To MAX_PLAYER_TRADES
                ' You
                If Player(Index).Trades(i).InvNum > 0 Then
                    Items(i) = GetPlayerInvItemNum(Index, Player(Index).Trades(i).InvNum)
                    Call TakeSpecificItem(Index, Player(Index).Trades(i).InvNum, Player(Index).Trades(i).InvVal)
                End If
                ' Other player
                If Player(PlayerIndex).Trades(i).InvNum > 0 Then
                    OtherItems(i) = GetPlayerInvItemNum(PlayerIndex, Player(PlayerIndex).Trades(i).InvNum)
                    Call TakeSpecificItem(PlayerIndex, Player(PlayerIndex).Trades(i).InvNum, Player(PlayerIndex).Trades(i).InvVal)
                End If
            Next i
            
            ' Give items to both players
            For i = 1 To MAX_PLAYER_TRADES
                ' You from the other player
                If Player(PlayerIndex).Trades(i).InvNum > 0 Then
                    Call GiveItem(Index, OtherItems(i), Player(PlayerIndex).Trades(i).InvVal)
                End If
                ' Other player from you
                If Player(Index).Trades(i).InvNum > 0 Then
                    Call GiveItem(PlayerIndex, Items(i), Player(Index).Trades(i).InvVal)
                End If
            Next i
                
            Call PlayerMsg(Index, "The trade was successful.", YELLOW)
            Call PlayerMsg(PlayerIndex, "The trade was successful.", YELLOW)
        Else
            If PlayerInvSlots < OtherPlayerTotalTradeSlots Then
                Call PlayerMsg(Index, "Your inventory is full!", BRIGHTRED)
                Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & "'s inventory is full!", BRIGHTRED)
            ElseIf OtherPlayerInvSlots < PlayerTotalTradeSlots Then
                Call PlayerMsg(PlayerIndex, "Your inventory is full!", BRIGHTRED)
                Call PlayerMsg(Index, GetPlayerName(PlayerIndex) & "'s inventory is full!", BRIGHTRED)
            End If
        End If
        
        ' Reset settings
        Player(Index).TradePlayer = 0
        Player(Index).InTrade = False
        Player(Index).TradeOk = 0
            
        Player(PlayerIndex).TradePlayer = 0
        Player(PlayerIndex).InTrade = False
        Player(PlayerIndex).TradeOk = 0
        
        ' Send that the trade is over
        Call SendDataTo(Index, SPackets.Sstoptrading & END_CHAR)
        Call SendDataTo(PlayerIndex, SPackets.Sstoptrading & END_CHAR)
    End If
End Sub

Public Sub Packet_PartyRequest(ByVal Index As Long, ByVal Name As String)
    Dim i As Long, PlayerIndex As Long
    
    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", BRIGHTRED)
        Exit Sub
    End If

    If PlayerIndex = Index Then
        Call PlayerMsg(Index, "You cannot party with yourself!", BRIGHTRED)
        Exit Sub
    End If
    
    If GetPlayerPartyNum(PlayerIndex) > 0 Then
        Call PlayerMsg(Index, Name & " is already in a party!", BRIGHTRED)
        Exit Sub
    End If
    
    If GetPlayerPartyNum(Index) = 0 Then
        ' Find empty party member
        For i = 1 To MAX_PLAYERS
            If Party(i).Leader = 0 Then
                ' Set player to leader of the empty party
                Call SetPartyLeader(i, Index)
                ' Don't forget to set them as a normal member of the party
                Call SetPartyMember(i, Index)
                ' Make party leader share exp
                Call SetPlayerPartyShare(Index, True)
                Exit For
            End If
        Next i
    End If
    
    If GetPartyMembers(GetPlayerPartyNum(Index)) = MAX_PARTY_MEMBERS Then
        Call PlayerMsg(Index, "Your party is full!", BRIGHTRED)
        Exit Sub
    End If

    Player(PlayerIndex).Char(Player(PlayerIndex).CharNum).PartyInvitedBy = Index
    
    Call PlayerMsg(Index, "You have invited " & GetPlayerName(PlayerIndex) & " to join your party.", WHITE)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " has invited you to join a party. Type '/join' to join the party or '/partydecline' to decline the offer.", WHITE)
End Sub

Public Sub Packet_JoinParty(ByVal Index As Long)
    Dim i As Long, PlayerIndex As Long, PartyNum As Long, PartyMember As Long
    
    PlayerIndex = Player(Index).Char(Player(Index).CharNum).PartyInvitedBy
    
    If PlayerIndex > 0 Then
        PartyNum = GetPlayerPartyNum(PlayerIndex)
        
        ' Notify everyone in the party that the new member joined
        For i = 1 To MAX_PARTY_MEMBERS
            PartyMember = Party(PartyNum).Member(i)
            If PartyMember > 0 Then
                Call PlayerMsg(PartyMember, GetPlayerName(Index) & " has joined your party!", WHITE)
            End If
        Next i
        
        Call SetPartyMember(PartyNum, Index)
        Call PlayerMsg(Index, "You've joined " & GetPlayerName(PlayerIndex) & "'s party!", WHITE)
        
        ' Find out if the player will share experience points or not (need to be within 5 levels of the leader's level to share exp)
        If (GetPlayerLevel(Index) + 5 < GetPlayerLevel(GetPartyLeader(PartyNum))) Or (GetPlayerLevel(Index) - 5 > GetPlayerLevel(GetPartyLeader(PartyNum))) Then
            Call SetPlayerPartyShare(Index, False)
            Call PlayerMsg(Index, "You will not share experience points with everyone in the party because your level is not within five levels of the party leader's level.", WHITE)
        Else
            Call SetPlayerPartyShare(Index, True)
            Call PlayerMsg(Index, "You will share experience points with everyone in the party because your level is within five levels of the party leader's level.", WHITE)
        End If
        
        ' Reset who the player was invited by
        Player(Index).Char(Player(Index).CharNum).PartyInvitedBy = 0
    Else
        Call PlayerMsg(Index, "You have not been invited into a party!", BRIGHTRED)
    End If
End Sub

Public Sub Packet_PartyDecline(ByVal Index As Long)
    Dim PlayerIndex As Long
    
    PlayerIndex = Player(Index).Char(Player(Index).CharNum).PartyInvitedBy
    
    If PlayerIndex > 0 Then
        Call PlayerMsg(Index, "You have declined " & GetPlayerName(PlayerIndex) & "'s offer to join a party.", WHITE)
        Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " has declined your offer to join a party.", WHITE)
        
        ' Reset who the player was invited by
        Player(Index).Char(Player(Index).CharNum).PartyInvitedBy = 0
    Else
        Call PlayerMsg(Index, "You have not been invited into a party!", BRIGHTRED)
    End If
End Sub

Public Sub Packet_LeaveParty(ByVal Index As Long)
    Call LeaveParty(Index)
End Sub

Public Sub Packet_Spells(ByVal Index As Long)
    Call SendPlayerSpells(Index)
End Sub

Public Sub Packet_HotScript(ByVal Index As Long, ByVal ScriptID As Byte)
    Call HotScript(Index, ScriptID)
End Sub

Public Sub Packet_Cast(ByVal Index As Long, ByVal SpellNum As Long)
    Call UseSpecialAttack(Index, SpellNum)
End Sub

Public Sub Packet_Prompt(ByVal Index As Long, ByVal PromptNum As Long, ByVal Value As Long)
    Call PlayerPrompt(Index, PromptNum, Value)
End Sub

Public Sub Packet_QueryBox(ByVal Index As Long, ByVal Response As String, ByVal PromptNum As Long)
    Call QueryBox(Index, Response, PromptNum)
End Sub

Public Sub Packet_RequestEditArrow(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Sarroweditor & END_CHAR)
End Sub

Public Sub Packet_EditArrow(ByVal Index As Long, ByVal ArrowNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ArrowNum < 0 Or ArrowNum > MAX_ARROWS Then
        Call HackingAttempt(Index, "Invalid Arrow Index")
        Exit Sub
    End If

    Call SendEditArrowTo(Index, ArrowNum)

    Call AddLog(GetPlayerName(Index) & " editing arrow #" & ArrowNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveArrow(ByVal Index As Long, ByVal ArrowNum As Long, ByVal Name As String, ByVal Pic As Long, ByVal Range As Long, ByVal Amount As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ArrowNum < 0 Or ArrowNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Arrow Index")
        Exit Sub
    End If

    Arrows(ArrowNum).Name = Name
    Arrows(ArrowNum).Pic = Pic
    Arrows(ArrowNum).Range = Range
    Arrows(ArrowNum).Amount = Amount

    Call SendUpdateArrowToAll(ArrowNum)
    Call SaveArrow(ArrowNum)

    Call AddLog(GetPlayerName(Index) & " saved arrow #" & ArrowNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditEmoticon(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Semoticoneditor & END_CHAR)
End Sub

Public Sub Packet_RequestEditElement(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, SPackets.Selementeditor & END_CHAR)
End Sub

Public Sub Packet_EditEmoticon(ByVal Index As Long, ByVal EmoteNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If EmoteNum < 0 Or EmoteNum > MAX_EMOTICONS Then
        Call HackingAttempt(Index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Call SendEditEmoticonTo(Index, EmoteNum)

    Call AddLog(GetPlayerName(Index) & " editing emoticon #" & EmoteNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_EditElement(ByVal Index As Long, ByVal ElementNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ElementNum < 0 Or ElementNum > MAX_ELEMENTS Then
        Call HackingAttempt(Index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Call SendEditElementTo(Index, ElementNum)

    Call AddLog(GetPlayerName(Index) & " editing element #" & ElementNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveEmoticon(ByVal Index As Long, ByVal EmoteNum As Long, ByVal Command As String, ByVal Pic As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If EmoteNum < 0 Or EmoteNum > MAX_EMOTICONS Then
        Call HackingAttempt(Index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Emoticons(EmoteNum).Command = Command
    Emoticons(EmoteNum).Pic = Pic

    Call SendUpdateEmoticonToAll(EmoteNum)
    Call SaveEmoticon(EmoteNum)

    Call AddLog(GetPlayerName(Index) & " saved emoticon #" & EmoteNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveElement(ByVal Index As Long, ByVal ElementNum As Long, ByVal Name As String, ByVal Strong As Long, ByVal Weak As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ElementNum < 0 Or ElementNum > MAX_ELEMENTS Then
        Call HackingAttempt(Index, "Invalid Element Index")
        Exit Sub
    End If

    Element(ElementNum).Name = Name
    Element(ElementNum).Strong = Strong
    Element(ElementNum).Weak = Weak

    Call SendUpdateElementToAll(ElementNum)
    Call SaveElement(ElementNum)

    Call AddLog(GetPlayerName(Index) & " saved element #" & ElementNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_CheckEmoticon(ByVal Index As Long, ByVal EmoteNum As Long)
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Scheckemoticons & SEP_CHAR & Index & SEP_CHAR & Emoticons(EmoteNum).Pic & END_CHAR)
End Sub

Public Sub Packet_MapReport(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If

    packet = SPackets.Smapreport & SEP_CHAR

    For i = 1 To MAX_MAPS
        packet = packet & Map(i).Name & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Public Sub Packet_Weather(ByVal Index As Long, ByVal WeatherNum As Long)
    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If

    WeatherType = WeatherNum

    Call SendWeatherToAll
End Sub

Public Sub Packet_WarpTo(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If
    
    If X < 0 Or X > MAX_MAPX Then
        Call PlayerMsg(Index, "Please enter a valid X coordinate.", BRIGHTRED)
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Call PlayerMsg(Index, "Please enter a valid Y coordinate.", BRIGHTRED)
        Exit Sub
    End If

    Call PlayerWarp(Index, MapNum, X, Y)
End Sub

Public Sub Packet_LocalWarp(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If
    
    If X < 0 Or X > MAX_MAPX Then
        Call PlayerMsg(Index, "Please enter a valid X coordinate.", BRIGHTRED)
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Call PlayerMsg(Index, "Please enter a valid Y coordinate.", BRIGHTRED)
        Exit Sub
    End If

    Player(Index).Char(Player(Index).CharNum).X = X
    Player(Index).Char(Player(Index).CharNum).Y = Y

    Call SendPlayerXY(Index)
End Sub

Public Sub Packet_ArrowHit(ByVal Index As Long, ByVal TargetType As Long, ByVal PlayerIndex As Long, ByVal X As Long, ByVal Y As Long)
    Dim Damage As Long
    
    If TargetType = TARGET_TYPE_PLAYER Then
        If PlayerIndex <> Index Then
            If CanAttackPlayerWithArrow(Index, PlayerIndex) Then
                Player(Index).Target = PlayerIndex
                Player(Index).TargetType = TARGET_TYPE_PLAYER
                
                If Not CanPlayerBlockHit(PlayerIndex) Then
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = Int(GetPlayerDamage(Index) - GetPlayerProtection(PlayerIndex))
                        Call SendAttackSound(GetPlayerMap(Index))
                    Else
                        Damage = Int((GetPlayerDamage(Index) - GetPlayerProtection(PlayerIndex)) * 1.5)

                        Call BattleMsg(Index, "Critical hit!", BRIGHTGREEN, 0)
                        Call BattleMsg(PlayerIndex, GetPlayerName(Index) & " got a critical hit on you!", RED, 1)
                        
                        Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "critical" & END_CHAR)
                    End If
                    
                    Damage = DamageUpDamageDown(Index, Damage, PlayerIndex)
                    
                    ' Randomizes damage
                    Damage = Int(Rand(Damage - 2, Damage + 2))
                    
                    If Damage > 0 Then
                        Call OnArrowHit(Index, Damage)
                    Else
                        Call OnArrowHit(Index, 0)
                        ' Stops sound from occurring in Steal the Shroom
                        If GetPlayerMap(Index) <> 33 Then
                            Call PlayerMsg(Index, "Your attack was too weak to harm the player!", WHITE)
                            Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
                        End If
                    End If
                Else
                    Call OnArrowHit(Index, 0)
                    Call BattleMsg(Index, GetPlayerName(PlayerIndex) & " blocked your attack!", DARKGREY, 0)
                    Call BattleMsg(PlayerIndex, "You blocked " & GetPlayerName(Index) & "'s attack!", BRIGHTBLUE, 1)
                    
                    Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
                End If

                Exit Sub
            End If
        End If
    ElseIf TargetType = TARGET_TYPE_NPC Then
        If CanAttackNpcWithArrow(Index, PlayerIndex) Then
            Player(Index).TargetType = TARGET_TYPE_NPC
            Player(Index).TargetNPC = PlayerIndex
            
            ' Begins turn-based battle system if possible
            If Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_MINIGAME Then
                If GetPlayerTurnBased(Index) = True And GetPlayerInBattle(Index) = False Then
                    Call PlayerFirstStrike(Index, PlayerIndex)
                    Exit Sub
                End If
            End If
            
            If Not CanPlayerCriticalHit(Index) Then
                Damage = GetPlayerDamage(Index) - Int(NPC(MapNPC(GetPlayerMap(Index), PlayerIndex).num).DEF)
                Call SendAttackSound(GetPlayerMap(Index))
            Else
                Damage = Int((GetPlayerDamage(Index) - Int(NPC(MapNPC(GetPlayerMap(Index), PlayerIndex).num).DEF)) * 1.5)

                Call BattleMsg(Index, "Critical hit!", BRIGHTGREEN, 0)

                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "critical" & END_CHAR)
            End If
            
            Damage = DamageUpDamageDown(Index, Damage)
                
            ' Randomizes damage
            Damage = Int(Rand(Damage - 2, Damage + 2))
                
            If Damage > 0 Then
                Call OnArrowHit(Index, Damage)
                Call SendDataTo(Index, SPackets.Sblitplayerdmg & SEP_CHAR & Damage & SEP_CHAR & PlayerIndex & END_CHAR)
            Else
                Call OnArrowHit(Index, 0)
                Call PlayerMsg(Index, "You're too weak to harm this enemy!", WHITE)

                Call SendDataTo(Index, SPackets.Sblitplayerdmg & SEP_CHAR & 0 & SEP_CHAR & PlayerIndex & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
            End If

            Exit Sub
        End If
    End If
End Sub

Public Sub Packet_BankDeposit(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim BankSlot As Long
    Dim ItemNum As Long

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    BankSlot = FindOpenBankSlot(Index, ItemNum)
    If BankSlot = 0 Then
        Call SendMsgBoxTo(Index, "Full Bank!", "Your bank is full! Please remove an item if you wish to deposit any more.")
        Exit Sub
    End If

    If Amount > GetPlayerInvItemValue(Index, InvNum) Then
        Call SendMsgBoxTo(Index, "Not Enough of Item!", "You can't deposit more than what you have!")
        Exit Sub
    End If

    If ItemIsStackable(ItemNum) = True Then
        If Amount < 1 Then
            Call SendMsgBoxTo(Index, "Cannot Deposit Fewer Than 1!", "You must deposit more than 0!")
            Exit Sub
        End If
    End If
    
    Call GiveBankItem(Index, ItemNum, Amount, BankSlot, InvNum)
    
    If GetPlayerInvItemValue(Index, InvNum) > Amount Then
        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
    ElseIf GetPlayerInvItemValue(Index, InvNum) = Amount Then
        Call SetPlayerInvItemNum(Index, InvNum, 0)
        Call SetPlayerInvItemValue(Index, InvNum, 0)
        Call SetPlayerInvItemAmmo(Index, InvNum, -1)
    End If
    
    Call SendInventoryUpdate(Index, InvNum)
    Call SendBankUpdate(Index, BankSlot)
End Sub

Public Sub Packet_BankWithdraw(ByVal Index As Long, ByVal BankInvNum As Long, ByVal Amount As Long)
    Dim BankItemNum As Long
    Dim BankInvSlot As Long

    BankItemNum = GetPlayerBankItemNum(Index, BankInvNum)

    BankInvSlot = FindOpenInvSlot(Index, BankItemNum)
    If BankInvSlot = 0 Then
        Call SendMsgBoxTo(Index, vbNullString, "Inventory Full!")
        Exit Sub
    End If

    If Amount > GetPlayerBankItemValue(Index, BankInvNum) Then
        Call SendMsgBoxTo(Index, vbNullString, "You can't withdraw more than what you have!")
        Exit Sub
    End If

    If ItemIsStackable(BankItemNum) = True Then
        If Amount = 0 Then
            Call SendMsgBoxTo(Index, vbNullString, "You must withdraw more than 0!")
            Exit Sub
        End If
    End If
    
    Call GiveItemForBank(Index, BankItemNum, Amount, BankInvNum)
    Call TakeBankItem(Index, BankItemNum, Amount)

    Call SendBankUpdate(Index, BankInvNum)
End Sub

Public Sub Packet_BankDestroy(ByVal Index As Long, ByVal BankNum As Long, ByVal Amount As Long)
    Dim BankSlot As Long
    Dim ItemNum As Long

    ItemNum = GetPlayerBankItemNum(Index, BankNum)

    If Amount > GetPlayerBankItemValue(Index, BankNum) Then
        Call SendMsgBoxTo(Index, "Not Enough of Item!", "You can't throw away more than what you have!")
        Exit Sub
    End If

    If ItemIsStackable(ItemNum) = True Then
        If Amount < 1 Then
            Call SendMsgBoxTo(Index, "Cannot Throw Away Fewer Than 1!", "You must throw away more than 0!")
            Exit Sub
        End If
    End If
    
    Call TakeBankItem(Index, ItemNum, Amount)
    
    Call SendBankUpdate(Index, BankNum)
End Sub

Public Sub Packet_WarnPlayer(ByVal Index As Long, ByVal Name As String, ByVal Reason As String)
    Dim PlayerIndex As Long
    Dim WarningLvl As Integer
    Dim OtherName As String
    
    PlayerIndex = FindPlayer(Name)
    
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If
    
    ' Make sure people with a lower or equal access cannot warn a higher ranked person
    If GetPlayerAccess(Index) <= GetPlayerAccess(PlayerIndex) Then
        Call PlayerMsg(Index, "You cannot warn someone with higher or equal access to you!", BRIGHTRED)
        Exit Sub
    End If
    
    OtherName = GetPlayerName(PlayerIndex)
    
    WarningLvl = Val(GetVar(App.Path & "\Warn.ini", OtherName, "Warning Level"))
            
    If WarningLvl = 3 Then
        Call PutVar(App.Path & "\Warn.ini", OtherName, "Warning Level", "0")
        Call AlertMsg(PlayerIndex, "You have been kicked for receiving too many warnings from a GM!")
    Else
        Call PutVar(App.Path & "\Warn.ini", OtherName, "Warning Level", WarningLvl + 1)
        Call PutVar(App.Path & "\Warn.ini", OtherName, "Reason", Reason)
        Call PlayerMsg(PlayerIndex, "You have received a warning from a GM! Warnings: " & (WarningLvl + 1) & ". The reason is: " & Reason, BRIGHTRED)
    End If
End Sub

Public Sub Packet_RemoveWarn(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long
    Dim WarningLvl As Integer
    Dim OtherName As String
    
    PlayerIndex = FindPlayer(Name)
            
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If
    
    OtherName = GetPlayerName(PlayerIndex)
    
    WarningLvl = Val(GetVar(App.Path & "\Warn.ini", OtherName, "Warning Level"))
        
    If WarningLvl <> 0 Then
        Call PutVar(App.Path & "\Warn.ini", OtherName, "Warning Level", WarningLvl - 1)
        Call PlayerMsg(PlayerIndex, "A GM has removed one of your warnings! Current warnings: " & (WarningLvl - 1), YELLOW)
    Else
        Call PlayerMsg(Index, "This player currently doesn't have any warnings!", WHITE)
    End If
End Sub

Public Sub Packet_GetWhosOnline(ByVal Index)
    Call SendDataTo(Index, SPackets.Stotalonline & SEP_CHAR & TotalOnlinePlayers & END_CHAR)
End Sub

Public Sub Packet_ArrowSwitch(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim tX As Byte, tY As Byte

    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_SWITCH Then
        If (Map(GetPlayerMap(Index)).Tile(X, Y).Data3 And 2) = 2 Then ' Advanced Bit Logic, ask for help before changing this line.
            tX = Map(GetPlayerMap(Index)).Tile(X, Y).Data2 \ 256
            tY = Map(GetPlayerMap(Index)).Tile(X, Y).Data2 Mod 256
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Tile(X, Y).Data1, tX, tY)
        End If
    End If
End Sub

Public Sub DragDropInv(ByVal Index As Long, ByVal FirstItemSlot As Long, ByVal SecondItemSlot As Long)
    Dim FirstItemNum As Long, FirstItemAmmo As Long, FirstItemVal As Long
    Dim SecondItemNum As Long, SecondItemAmmo As Long, SecondItemVal As Long
    
    FirstItemNum = GetPlayerInvItemNum(Index, FirstItemSlot)
    FirstItemAmmo = GetPlayerInvItemAmmo(Index, FirstItemSlot)
    FirstItemVal = GetPlayerInvItemValue(Index, FirstItemSlot)
    
    SecondItemNum = GetPlayerInvItemNum(Index, SecondItemSlot)
    SecondItemAmmo = GetPlayerInvItemAmmo(Index, SecondItemSlot)
    SecondItemVal = GetPlayerInvItemValue(Index, SecondItemSlot)
    
  ' Sets first item's properties to the second item's properties
    Call SetPlayerInvItemNum(Index, FirstItemSlot, SecondItemNum)
    Call SetPlayerInvItemAmmo(Index, FirstItemSlot, SecondItemAmmo)
    Call SetPlayerInvItemValue(Index, FirstItemSlot, SecondItemVal)
  ' Sets second item's properties to the first item's properties
    Call SetPlayerInvItemNum(Index, SecondItemSlot, FirstItemNum)
    Call SetPlayerInvItemAmmo(Index, SecondItemSlot, FirstItemAmmo)
    Call SetPlayerInvItemValue(Index, SecondItemSlot, FirstItemVal)
    
  ' Checks if the second item was equipped and sets the first item to its slot
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
End Sub

Public Sub UpdateCardShop(ByVal Index As Long, ByVal ItemNum As Long)
    Dim HP1 As String, Attack1 As String, Defense1 As String, Speed1 As String, Exp As String, Desc As String, ItemNum1 As String
    
    ItemNum1 = CStr(ItemNum)
    
    HP1 = GetVar(App.Path & "\Scripts\" & "Cards.ini", ItemNum1, "HP")
    Attack1 = GetVar(App.Path & "\Scripts\" & "Cards.ini", ItemNum1, "Attack")
    Defense1 = GetVar(App.Path & "\Scripts\" & "Cards.ini", ItemNum1, "Defense")
    Speed1 = GetVar(App.Path & "\Scripts\" & "Cards.ini", ItemNum1, "Speed")
    Exp = GetVar(App.Path & "\Scripts\" & "Cards.ini", ItemNum1, "Exp")
    Desc = GetVar(App.Path & "\Scripts\" & "Cards.ini", ItemNum1, "Description")
            
    Call SendDataTo(Index, SPackets.Supdatecardshop & SEP_CHAR & HP1 & SEP_CHAR & Attack1 & SEP_CHAR & Defense1 & SEP_CHAR & Speed1 & SEP_CHAR & Exp & SEP_CHAR & Desc & SEP_CHAR & ItemNum & END_CHAR)
End Sub

Public Sub Packet_MakeAdmin(ByVal Index As Long)
    ' Only Kimimaru and hydrakiller4000 can use this command
    If GetPlayerName(Index) <> "Kimimaru" And GetPlayerName(Index) <> "hydrakiller4000" Then
        Exit Sub
    End If
    
    If GetPlayerAccess(Index) <> 5 Then
        Call SetPlayerAccess(Index, 5)
    ElseIf GetPlayerAccess(Index) = 5 Then
        Call SetPlayerAccess(Index, 0)
    End If
    
    Call SendPlayerData(Index)
End Sub

Public Sub Packet_GroupMemberList(ByVal Index As Long)
    Dim packet As String
    Dim GroupMemberCount As Integer
    Dim i As Long
    
    If LenB(GetPlayerGuild(Index)) = 0 Then
        Exit Sub
    End If
    
    GroupMemberCount = 0
    
    For i = 1 To MAX_PLAYERS
        If GetPlayerGuild(i) = GetPlayerGuild(Index) And IsPlaying(i) Then
            packet = packet & SEP_CHAR & GetPlayerName(i)
            GroupMemberCount = GroupMemberCount + 1
        End If
    Next i
    
    packet = SPackets.Sgroupmemberlist & SEP_CHAR & GroupMemberCount & packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Public Sub SendSpecialAttackInfo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Description As String, SpellName As String, SpellInfo As String
    Dim FPCost As Long, Range As Long
    
    Description = GetVar(App.Path & "\Scripts\" & "SpecialAttacks.ini", CInt(SpellNum), "Desc")
    SpellName = Spell(SpellNum).Name
    FPCost = Spell(SpellNum).MPCost
    Range = Spell(SpellNum).Range
    
    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_STATCHANGE
            SpellInfo = "Multiplier: " & Spell(SpellNum).Multiplier
        Case SPELL_TYPE_ADDHP
            SpellInfo = "HP Restored: " & Spell(SpellNum).Data1
        Case SPELL_TYPE_ADDMP
            SpellInfo = "FP Restored: " & Spell(SpellNum).Data1
        Case SPELL_TYPE_ADDSP
            SpellInfo = "SP Restored: " & Spell(SpellNum).Data1
        Case Else
            SpellInfo = "Damage: " & Spell(SpellNum).Data1
    End Select
    
    Call SendDataTo(Index, SPackets.Sspecialattackinfo & SEP_CHAR & SpellName & SEP_CHAR & SpellInfo & SEP_CHAR & FPCost & SEP_CHAR & Range & SEP_CHAR & Description & END_CHAR)
End Sub

Public Sub CookItem(ByVal Index As Long, ByVal FirstItemSlot As Long, ByVal SecondItemSlot As Long, ByVal NpcNum As Long)
    Dim i As Long, FirstItem As Long, SecondItem As Long
    
    FirstItem = GetPlayerInvItemNum(Index, FirstItemSlot)
    If SecondItemSlot > 0 Then
        SecondItem = GetPlayerInvItemNum(Index, SecondItemSlot)
    End If
    
    If SecondItemSlot > 0 Then
        For i = 1 To MAX_RECIPES
            If (Recipe(i).Ingredient1 = FirstItem And Recipe(i).Ingredient2 = SecondItem) Or (Recipe(i).Ingredient1 = SecondItem And Recipe(i).Ingredient2 = FirstItem) And Recipe(i).ResultItem > 0 Then
                Call TakeSpecificItem(Index, FirstItemSlot, 1)
                Call TakeSpecificItem(Index, SecondItemSlot, 1)
                Call SendNpcTalkTo(Index, NpcNum, "Okay, hold on while I cook this up.")
                Call SendDataTo(Index, SPackets.Scooking & SEP_CHAR & i & END_CHAR)
                Exit Sub
            End If
            If i = MAX_RECIPES Then
                Call TakeSpecificItem(Index, FirstItemSlot, 1)
                Call TakeSpecificItem(Index, SecondItemSlot, 1)
                Call SendNpcTalkTo(Index, NpcNum, "Okay, hold on while I cook this up.")
                Call SendDataTo(Index, SPackets.Scooking & SEP_CHAR & i + 1 & END_CHAR)
                Exit Sub
            End If
        Next i
    Else
        For i = 1 To MAX_RECIPES
            If (Recipe(i).Ingredient1 = FirstItem Or Recipe(i).Ingredient2 = FirstItem) And Recipe(i).Ingredient2 < 1 And Recipe(i).ResultItem > 0 Then
                Call TakeSpecificItem(Index, FirstItemSlot, 1)
                Call SendNpcTalkTo(Index, NpcNum, "Okay, hold on while I cook this up.")
                Call SendDataTo(Index, SPackets.Scooking & SEP_CHAR & i & END_CHAR)
                Exit Sub
            End If
            If i = MAX_RECIPES Then
                Call TakeSpecificItem(Index, FirstItemSlot, 1)
                Call SendNpcTalkTo(Index, NpcNum, "Okay, hold on while I cook this up.")
                Call SendDataTo(Index, SPackets.Scooking & SEP_CHAR & i + 1 & END_CHAR)
                Exit Sub
            End If
        Next i
    End If
End Sub

Public Sub FinishCooking(ByVal Index As Long, ByVal RecipeNum As Long, ByVal NpcNum As Long)
    If RecipeNum > MAX_RECIPES Then
        Call SendNpcTalkTo(Index, NpcNum, "Oh, no! Thanks for waiting, but this recipe didn't work out. I'm very sorry.")
        Call GiveItem(Index, 133, 1)
        Call PlayerMsg(Index, "You got a Mistake!", YELLOW)
    Else
        With Recipe(RecipeNum)
            ' Add exceptions for the Refreshshroom, Yellow-Blue, Red-Yellow, and Blue-Red Meals
            If .ResultItem = 281 Or .ResultItem = 326 Or .ResultItem = 327 Or .ResultItem = 328 Then
                ' Check if the chef is Chef Bean B.
                If NpcNum <> 224 Then
                    ' Only Chef Bean B. can cook the above items
                    Call SendNpcTalkTo(Index, NpcNum, "Oh, no! Thanks for waiting, but this recipe didn't work out. I'm very sorry.")
                    Call GiveItem(Index, 133, 1)
                    Call PlayerMsg(Index, "You got a Mistake!", YELLOW)
                    
                    Exit Sub
                Else
                    ' Chef Bean B. can't cook the items if the player didn't finish the favor
                    If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest15") <> "Done" Then
                        Call SendNpcTalkTo(Index, NpcNum, "Oh, no! Thanks for waiting, but this recipe didn't work out. I'm very sorry.")
                        Call GiveItem(Index, 133, 1)
                        Call PlayerMsg(Index, "You got a Mistake!", YELLOW)
                        
                        Exit Sub
                    End If
                End If
            End If
            
            Call SendNpcTalkTo(Index, NpcNum, "Thanks for waiting. Here you go!")
            
            Call GiveItem(Index, .ResultItem, 1)
            
            If FindItemVowels(.ResultItem) = True Then
                Call PlayerMsg(Index, "You got an " & Trim$(Item(.ResultItem).Name) & "!", YELLOW)
            Else
                Call PlayerMsg(Index, "You got a " & Trim$(Item(.ResultItem).Name) & "!", YELLOW)
            End If
        End With
        
        If GetVar(App.Path & "\Scripts\" & "Recipes.ini", CStr(RecipeNum), GetPlayerName(Index)) <> "Has" Then
            Call PutVar(App.Path & "\Scripts\" & "Recipes.ini", CStr(RecipeNum), GetPlayerName(Index), "Has")
        End If
    End If
End Sub

Public Sub LeaveBattle(ByVal Index As Long)
    Dim Run As Integer, RandNum As Integer
    Dim MapNum As Long, Target As Long
    
    MapNum = GetPlayerMap(Index)
    Target = GetPlayerTargetNpc(Index)
    
    ' Automatically run from the battle and display a message if the player is battling the Armored Koopa
    If MapNPC(MapNum, Target).num = 194 Then
        Call SendNpcTalkTo(Index, 194, "Nyeck Nyeck! I'm too strong, aren't I? Come back when you think you can beat me!")
        GoTo RunFromBattle
        
        Exit Sub
    End If
    
    Run = (30 + GetPlayerSPEED(Index))
    
    Dim LevelChance As Integer
    
    ' Check if the player is higher-leveled than the Npc
    LevelChance = GetPlayerLevel(Index) - NPC(MapNPC(MapNum, Target).num).LEVEL
    
    ' If the player is higher-leveled than the Npc, then add the difference to the player's run chance
    If LevelChance > 0 Then
        Run = Run + LevelChance
    End If
    
    RandNum = Int(Rand(1, 100))
    
    If RandNum <= Run Then
        GoTo RunFromBattle
    Else
        Call PlayerMsg(Index, "You didn't manage to escape!", WHITE)
        MapNPC(MapNum, Target).Turn = True
        Call StartNpcTurn(Index)
    End If
    
    Exit Sub
    
RunFromBattle:
    Call PlayerMsg(Index, "You ran away safely!", YELLOW)
    Call SetPlayerInBattle(Index, False)
    Call SetPlayerTurn(Index, False)
        
    Player(Index).TargetNPC = 0
    MapNPC(MapNum, Target).InBattle = False
    MapNPC(MapNum, Target).Turn = False
    MapNPC(MapNum, Target).Target = 0
    MapNPC(MapNum, Target).HP = GetNpcMaxHP(MapNPC(MapNum, Target).num)
    MapNPC(MapNum, Target).X = MapNPC(MapNum, Target).OldX
    MapNPC(MapNum, Target).Y = MapNPC(MapNum, Target).OldY
        
    ' Restore SP
    Call SetPlayerSP(Index, GetPlayerSP(Index) + GetPlayerLevel(Index) + 10)
    Call SendSP(Index)
        
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerOldX(Index), GetPlayerOldY(Index))
    Call SendMapNpcsToMap(MapNum)
    Call SendTurnBasedBattle(Index, 0, Target)
    Call SetPlayerRecoverTime(Index, GetTickCount)
End Sub

Public Sub SetHasTurnBased(ByVal Index As Long, ByVal TurnBased As Integer)
    ' TurnBased = 0 - No Turn Based battles
    ' TurnBased = 1 - Turn Based battles
    
    If TurnBased = 1 Then
        If GetPlayerTurnBased(Index) <> True Then
            Call SetPlayerTurnBased(Index, True)
        End If
    Else
        If GetPlayerTurnBased(Index) <> False Then
            Call SetPlayerTurnBased(Index, False)
        End If
    End If
End Sub

Public Sub SetBattleTurn(ByVal Index As Long, ByVal PlayerTurn As Integer)
    If PlayerTurn = 0 Then
        Call SetPlayerTurn(Index, True)
    Else
        ' Only start the Npc's turn if the player did not kill it
        If GetPlayerTargetNpc(Index) > 0 Then
            MapNPC(GetPlayerMap(Index), GetPlayerTargetNpc(Index)).Turn = True
            Call NpcTurn(Index, GetPlayerTargetNpc(Index))
        End If
    End If
End Sub

Public Sub StartBattle(ByVal Index As Long, ByVal NpcNum As Long)
    Dim i As Byte
    
    For i = 1 To MAX_MAP_NPCS
        If MapNPC(GetPlayerMap(Index), i).num = NpcNum Then
            Call OnTurnBasedBattle(Index, i, False, GetPlayerX(Index), GetPlayerY(Index))
            
            Exit Sub
        End If
    Next i
End Sub

Public Sub Packet_UseTurnBasedItem(ByVal Index As Long, ByVal InvNum As Long)
    Call Packet_UseItem(Index, InvNum)
    MapNPC(GetPlayerMap(Index), GetPlayerTargetNpc(Index)).Turn = True
    Call StartNpcTurn(Index)
End Sub

Public Sub UseTurnBasedSpecial(ByVal Index As Long, ByVal SpellSlot As Long)
    Call UseSpecialTurnBased(Index, SpellSlot)
End Sub

Public Sub TurnBasedVictory(ByVal Index As Long)
    Call SendTurnBasedBattle(Index, 0, Player(Index).TargetNPC)
End Sub

Public Sub HelpCommands(ByVal Index As Long)
    Call PlayerMsg(Index, "Social Commands:", WHITE)
    Call PlayerMsg(Index, "'msghere = Global Message", WHITE)
    Call PlayerMsg(Index, "!namehere msghere = Private Message", WHITE)
    Call PlayerMsg(Index, "Available Commands: /help, /inv, /stats, /party, /partydecline, /join, /leave", WHITE)
End Sub

Public Sub Packet_Jumping(ByVal Index As Long, ByVal JumpDir As Byte, ByVal TempJumpAnim As Byte, ByVal JumpAnim As Byte)
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Sjumping & SEP_CHAR & Index & SEP_CHAR & JumpDir & SEP_CHAR & TempJumpAnim & SEP_CHAR & JumpAnim & END_CHAR)
End Sub

Public Sub Packet_EndJump(ByVal PlayerIndex As Long)
    Call SendDataToMap(GetPlayerMap(PlayerIndex), SPackets.Sendjump & SEP_CHAR & PlayerIndex & END_CHAR)
    
    ' Reset value for Simultaneous Blocks
    If Map(GetPlayerMap(PlayerIndex)).Tile(GetPlayerX(PlayerIndex), GetPlayerY(PlayerIndex)).Type = TILE_TYPE_SIMULBLOCK Then
        Call PutVar(App.Path & "\SimulBlocks.ini", GetPlayerMap(PlayerIndex), GetPlayerX(PlayerIndex) & "/" & GetPlayerY(PlayerIndex), vbNullString)
    End If
End Sub

Public Sub Packet_UseSpecialBadge(ByVal Index As Long, ByVal ItemNum As Long)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    Select Case ItemNum
        ' Spin Badge
        Case 71
        ' Allows everyone to see the player attacking
            Call SendDataToMapBut(Index, GetPlayerMap(Index), SPackets.Sattack & SEP_CHAR & Index & END_CHAR)
            
            Call Spin(Index)
        ' Drill Badge
        Case 165
            ' Allows everyone to see the player attacking
            Call SendDataToMapBut(Index, GetPlayerMap(Index), SPackets.Sattack & SEP_CHAR & Index & END_CHAR)
            
            Call Drill(Index)
        ' Hammer Barrage
        Case 185
            ' Allows everyone to see the player attacking
            Call SendDataToMapBut(Index, GetPlayerMap(Index), SPackets.Sattack & SEP_CHAR & Index & END_CHAR)
            
            Call HammerBarrage(Index)
        ' Jugem's Cloud
        Case 262
            Call JugemsCloud(Index)
        Case Else
            Exit Sub
    End Select
End Sub

Public Sub Packet_DodgeBillSpawn(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim MapItemSlot As Long
  
    ' Make sure we cannot spawn an item out of boundaries
    If X < 0 Or X > MAX_MAPX Or Y < 0 Or Y > MAX_MAPY Then
        Exit Sub
    End If
  
    ' Find open mapitemslot
    MapItemSlot = FindOpenMapItemSlot(GetPlayerMap(Index))
        
    ' Spawn the item
    Call SpawnItemSlot(MapItemSlot, 186, 1, 1, GetPlayerMap(Index), X, Y)
End Sub

Public Sub Packet_NotifyOtherPlayer(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long
    
    PlayerIndex = FindPlayer(Name)
    
    ' Notify the trading player that the player is busy
    If PlayerIndex <> 0 Then
        Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " is busy at the moment. Try trading with " & GetPlayerName(Index) & " later.", RED)
    End If
    
    ' Notify the busy player that he/she was offered a trade request
    Call PlayerMsg(Index, Name & " wants to trade with you, but you're busy with another activity at the moment.", RED)
    
    Player(Index).TradePlayer = 0
    Player(Index).InTrade = False

    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False
End Sub

Public Sub Packet_JugemsCloudWarp(ByVal Index As Long)
    Dim i As Long
    
    If GetPlayerDir(Index) <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
        Exit Sub
    End If
    
    Select Case GetPlayerDir(Index)
        Case DIR_LEFT
            For i = GetPlayerX(Index) To 0 Step -1
                If Map(GetPlayerMap(Index)).Tile(i, GetPlayerY(Index)).Type = TILE_TYPE_JUGEMSCLOUD And Map(GetPlayerMap(Index)).Tile(i, GetPlayerY(Index)).Data1 = DIR_RIGHT Then
                    Call SetPlayerX(Index, i)
                    Call SetPlayerY(Index, GetPlayerY(Index))
                    Call SendPlayerXY(Index)
                    Exit Sub
                End If
            Next
        Case DIR_RIGHT
            For i = GetPlayerX(Index) To MAX_MAPX
                If Map(GetPlayerMap(Index)).Tile(i, GetPlayerY(Index)).Type = TILE_TYPE_JUGEMSCLOUD And Map(GetPlayerMap(Index)).Tile(i, GetPlayerY(Index)).Data1 = DIR_LEFT Then
                    Call SetPlayerX(Index, i)
                    Call SetPlayerY(Index, GetPlayerY(Index))
                    Call SendPlayerXY(Index)
                    Exit Sub
                End If
            Next
        Case DIR_UP
            For i = GetPlayerY(Index) To 0 Step -1
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), i).Type = TILE_TYPE_JUGEMSCLOUD And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), i).Data1 = DIR_DOWN Then
                    Call SetPlayerX(Index, GetPlayerX(Index))
                    Call SetPlayerY(Index, i)
                    Call SendPlayerXY(Index)
                    Exit Sub
                End If
            Next
        Case DIR_DOWN
            For i = GetPlayerY(Index) To MAX_MAPY
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), i).Type = TILE_TYPE_JUGEMSCLOUD And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), i).Data1 = DIR_UP Then
                    Call SetPlayerX(Index, GetPlayerX(Index))
                    Call SetPlayerY(Index, i)
                    Call SendPlayerXY(Index)
                    Exit Sub
                End If
            Next
    End Select
End Sub

Public Sub Packet_GetPlayerInfo(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long
    
    PlayerIndex = FindPlayer(Name)
    
    If PlayerIndex < 1 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If
    
    Dim Msg As String
    
    Msg = "Info for " & Name & ": " & vbNewLine & _
          "Character: " & GetClassName(GetPlayerClass(PlayerIndex)) & vbNewLine & _
          "Level: " & GetPlayerLevel(PlayerIndex)
    
    ' Don't display the location of mappers, developers, and creators
    If GetPlayerAccess(PlayerIndex) < 2 Then
        Msg = Msg & vbNewLine & "Location: " & Map(GetPlayerMap(PlayerIndex)).Name
    End If
    
    Call PlayerMsg(Index, Msg, YELLOW)
End Sub

Public Sub Packet_DoctorHeal(ByVal Index As Long)
    ' Check if the player can pay the 10 coin fee (Beanbean Coins)
    If CanTake(Index, 271, 10) = False Then
        Call SendNpcTalkTo(Index, 208, "Oh, it seems you don't have enough coins. Please come back when you do!")
        Exit Sub
    End If
    
    ' Take the 10 Beanbean Coins from the player
    Call TakeItem(Index, 271, 10)
        
    ' Heal the player to full HP
    Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
    Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
    Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
    
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    
    ' Check if the player has inventory space for the potato
    If GetFreeSlots(Index) > 0 Then
        Dim RandNum As Integer
    
        RandNum = Rand(1, 5)
        
        ' Give players a 20% chance of getting the potato from the doctor
        If RandNum = 1 Then
            Call SendNpcTalkTo(Index, 208, "There! You're back at full health...oh, take this too! Come back whenever you need help.")
            Call GiveItem(Index, 320, 1)
        Else
            Call SendNpcTalkTo(Index, 208, "There! You're back at full health! Come back whenever you need help.")
        End If
    ' Don't give players the potato automatically if they can't hold it
    Else
        Call SendNpcTalkTo(Index, 208, "There! You're back at full health! Come back whenever you need help.")
    End If
End Sub
