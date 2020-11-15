Attribute VB_Name = "modDatabase"
Option Explicit

' ---------------------------------------------------------------------------------------
' Procedure : GetVar
' Purpose   :  Reads a variable from an INI file
' ---------------------------------------------------------------------------------------
Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    On Error GoTo GetVar_Error

    szReturn = vbNullString

    sSpaces = Space(5000)

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), file)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

    On Error GoTo 0
    Exit Function

GetVar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetVar of Module modDatabase"
End Function

' ---------------------------------------------------------------------------------------
' Procedure : PutVar
' Purpose   : Writes a file to an INI file
' ---------------------------------------------------------------------------------------
Sub PutVar(file As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error

    Call WritePrivateProfileString(Header, Var, Value, file)

    On Error GoTo 0
    Exit Sub

PutVar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"
End Sub

Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
ErrorHandler:
' if an error occurs, this function returns False
End Function

Function FolderExists(inPath As String) As Boolean
    If LenB(Dir(inPath, vbDirectory)) = 0 Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function

Sub LoadExperience()
    On Error GoTo ExpErr
    Dim FileName As String
    Dim i As Integer

    Call CheckExperience

    FileName = App.Path & "\Experience.ini"

    For i = 1 To MAX_LEVEL
        temp = i / MAX_LEVEL * 100
        Call SetStatus("Loading Experience... " & temp & "%")
        Experience(i) = GetVar(FileName, "EXPERIENCE", "Exp" & i)
    Next i
    Exit Sub

ExpErr:
    Call MsgBox("Error loading EXP for level " & i & ". Make sure Experience.ini has the correct variables! ERR: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    Call DestroyServer
End Sub

Sub CheckExperience()
    If Not FileExists("Experience.ini") Then
        Dim i As Integer

        For i = 1 To MAX_LEVEL
            temp = i / MAX_LEVEL * 100
            Call SetStatus("Saving Experience... " & temp & "%")
            Call PutVar(App.Path & "\Experience.ini", "EXPERIENCE", "Exp" & i, i * 1500)
        Next i
    End If
End Sub

Sub ClearExperience()
    Dim i As Integer

    For i = 1 To MAX_LEVEL
        Experience(i) = 0
    Next i
End Sub

Sub LoadEmoticon()
    Dim FileName As String
    Dim i As Integer

    Call CheckEmoticon

    FileName = App.Path & "\Emoticons.ini"

    For i = 0 To MAX_EMOTICONS
        temp = i / MAX_EMOTICONS * 100
        Call SetStatus("Loading Emoticons... " & temp & "%")
        Emoticons(i).Pic = GetVar(FileName, "EMOTICONS", "Emoticon" & i)
        Emoticons(i).Command = GetVar(FileName, "EMOTICONS", "EmoticonC" & i)
    Next i
End Sub

Sub SaveEmoticon(ByVal EmoNum As Long)
    Dim FileName As String

    FileName = App.Path & "\Emoticons.ini"

    Call PutVar(FileName, "EMOTICONS", "EmoticonC" & EmoNum, Trim$(Emoticons(EmoNum).Command))
    Call PutVar(FileName, "EMOTICONS", "Emoticon" & EmoNum, Val(Emoticons(EmoNum).Pic))
End Sub

Sub CheckEmoticon()
    If Not FileExists("Emoticons.ini") Then
        Dim i As Integer

        For i = 0 To MAX_EMOTICONS
            temp = i / MAX_LEVEL * 100
            Call SetStatus("Saving emoticons... " & temp & "%")
            Call PutVar(App.Path & "\Emoticons.ini", "EMOTICONS", "Emoticon" & i, 0)
            Call PutVar(App.Path & "\Emoticons.ini", "EMOTICONS", "EmoticonC" & i, vbNullString)
        Next i
    End If
End Sub

Sub ClearEmoticon()
    Dim i As Integer

    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = vbNullString
    Next i
End Sub

Sub LoadElements()
    On Error GoTo ElementErr
    Dim FileName As String
    Dim i As Integer

    Call CheckElements

    FileName = App.Path & "\Elements.ini"

    For i = 0 To MAX_ELEMENTS
        temp = i / MAX_ELEMENTS * 100
        Call SetStatus("Loading elements... " & temp & "%")
        Element(i).Name = GetVar(FileName, "ELEMENTS", "ElementName" & i)
        Element(i).Strong = Val(GetVar(FileName, "ELEMENTS", "ElementStrong" & i))
        Element(i).Weak = Val(GetVar(FileName, "ELEMENTS", "ElementWeak" & i))
    Next i
    Exit Sub

ElementErr:
    Call MsgBox("Error loading element " & i & ". Make sure all the variables in Elements.ini are correct!", vbCritical)
    Call DestroyServer
    End
End Sub

Sub CheckElements()
    If Not FileExists("Elements.ini") Then
        Dim i As Integer

        For i = 0 To MAX_ELEMENTS
            temp = i / MAX_ELEMENTS * 100
            Call SetStatus("Saving elements... " & temp & "%")
            Call PutVar(App.Path & "\Elements.ini", "ELEMENTS", "ElementName" & i, vbNullString)
            Call PutVar(App.Path & "\Elements.ini", "ELEMENTS", "ElementStrong" & i, 0)
            Call PutVar(App.Path & "\Elements.ini", "ELEMENTS", "ElementWeak" & i, 0)
        Next i
    End If
End Sub

Sub SaveElement(ByVal ElementNum As Long)
    Dim FileName As String

    FileName = App.Path & "\Elements.ini"

    Call PutVar(FileName, "ELEMENTS", "ElementName" & ElementNum, Trim$(Element(ElementNum).Name))
    Call PutVar(FileName, "ELEMENTS", "ElementStrong" & ElementNum, Val(Element(ElementNum).Strong))
    Call PutVar(FileName, "ELEMENTS", "ElementWeak" & ElementNum, Val(Element(ElementNum).Weak))
End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim FileName As String
    Dim f As Long 'File
    Dim i As Integer
    Dim Value As String
    
    On Error Resume Next

    ' Save login information first
    FileName = App.Path & "\SMBOAccounts\" & Trim$(Player(Index).Login) & "_Info.ini"

    ' Saves any changes to passwords
    Call PutVar(FileName, "ACCESS", "Password", Trim$(Player(Index).Password))

    Call PutVar(FileName, "ACCESS", "Login", Trim$(Player(Index).Login))

    ' Make the directory
    If LCase$(Dir(App.Path & "\SMBOAccounts\" & Trim$(Player(Index).Login), vbDirectory)) <> LCase$(Trim$(Player(Index).Login)) Then
        Call MkDir(App.Path & "\SMBOAccounts\" & Trim$(Player(Index).Login))
    End If

    ' Now save their characters
    For i = 1 To MAX_CHARS
        FileName = App.Path & "\SMBOAccounts\" & Trim$(Player(Index).Login) & "\Char" & i & ".dat"

        ' Save the character
        f = FreeFile
        Open FileName For Binary As #f
        Put #f, , Player(Index).Char(i)
        Close #f
    Next i
End Sub

Function ConvertV000(FileName As String) As PlayerRec
    'Dim OldRec As V000PlayerRec
    'Dim NewRec As PlayerRec
    'Dim f As Long
    'Dim n As Integer

    'f = FreeFile
    'Open FileName For Binary As #f
    '    Get #f, , OldRec
    'Close #f

    ' General
    'NewRec.Name = OldRec.Name
    'NewRec.Guild = OldRec.Guild
    'NewRec.GuildAccess = OldRec.GuildAccess
    'NewRec.Sex = OldRec.Sex
    'NewRec.Class = OldRec.Class
    'NewRec.Sprite = OldRec.Sprite
    'NewRec.LEVEL = OldRec.LEVEL
    'NewRec.Exp = OldRec.Exp
    'NewRec.Access = OldRec.Access
    'NewRec.PK = OldRec.PK

    ' Vitals
    'NewRec.HP = OldRec.HP
    'NewRec.MP = OldRec.MP
    'NewRec.SP = OldRec.SP

    ' Stats
    'NewRec.STR = OldRec.STR
    'NewRec.DEF = OldRec.DEF
    'NewRec.Speed = OldRec.Speed
    'NewRec.Magi = OldRec.Magi
    'NewRec.POINTS = OldRec.POINTS

    ' Worn equipment
    'NewRec.ArmorSlot = OldRec.ArmorSlot
    'NewRec.WeaponSlot = OldRec.WeaponSlot
    'NewRec.HelmetSlot = OldRec.HelmetSlot
    'NewRec.ShieldSlot = OldRec.ShieldSlot
    'NewRec.LegsSlot = OldRec.LegsSlot
    'NewRec.RingSlot = OldRec.RingSlot
    'NewRec.NecklaceSlot = OldRec.NecklaceSlot

    ' Inventory
    'For n = 1 To MAX_INV
    '    NewRec.Inv(n).num = OldRec.Inv(n).num
    '    NewRec.Inv(n).Value = OldRec.Inv(n).Value
    '    NewRec.Inv(n).Ammo = OldRec.Inv(n).Ammo
    'Next n
    'For n = 1 To MAX_PLAYER_SPELLS
    '    NewRec.Spell(n) = OldRec.Spell(n)
    'Next n
    'For n = 1 To MAX_BANK
    '    NewRec.Bank(n).num = OldRec.Bank(n).num
    '    NewRec.Bank(n).Value = OldRec.Bank(n).Value
    '    NewRec.Bank(n).Ammo = OldRec.Bank(n).Ammo
    'Next n

    ' Position
    'NewRec.Map = OldRec.Map
    'NewRec.X = OldRec.X
    'NewRec.Y = OldRec.Y
    'NewRec.Dir = OldRec.Dir

    'NewRec.TargetNPC = OldRec.TargetNPC

    'NewRec.Head = OldRec.Head
    'NewRec.Body = OldRec.Body
    'NewRec.Leg = OldRec.Leg

    'NewRec.PAPERDOLL = OldRec.PAPERDOLL

    'NewRec.MAXHP = OldRec.MAXHP
    'NewRec.MAXMP = OldRec.MAXMP
    'NewRec.MAXSP = OldRec.MAXSP


    ' *** add new fields ***

    ' version info

    'NewRec.Vflag = 128
    'NewRec.Ver = 2
    'NewRec.SubVer = 8
    'NewRec.Rel = 0

    'ConvertV000 = NewRec
 End Function

Public Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim f As Long
    Dim i As Integer
    Dim FileName As String
    Dim WasConverted As Boolean

    On Error GoTo PlayerErr

    Call ClearPlayer(Index)

    ' Load the account settings
    FileName = App.Path & "\SMBOAccounts\" & Trim$(Name) & "_Info.ini"

    Player(Index).Login = Name
    Player(Index).Password = GetVar(FileName, "ACCESS", "Password")
    Player(Index).Email = vbNullString

    ' Load the .dat
    For i = 1 To MAX_CHARS
        FileName = App.Path & "\SMBOAccounts\" & Trim$(Player(Index).Login) & "\Char" & i & ".dat"

        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Player(Index).Char(i)
        Close #f
        'If Player(Index).Char(i).Vflag <> 128 Then
        '    Player(Index).Char(i) = ConvertV000(FileName)
        'End If
        
        Player(Index).Char(i).PartyNum = 0
        
        ' Check if the player's inventory was converted to the new system
        If Player(Index).Char(i).InvConverted = 0 Then
            Player(Index).Char(i).MaxInv = 24
        
            ReDim Player(Index).Char(i).NewInv(1 To Player(Index).Char(i).MaxInv) As NewPlayerInvRec
            
            Dim j As Integer
            
            ' Convert the player's inventory to the new system
            For j = 1 To Player(Index).Char(i).MaxInv
                Player(Index).Char(i).NewInv(j).num = Player(Index).Char(i).Inv(j).num
                Player(Index).Char(i).NewInv(j).Value = Player(Index).Char(i).Inv(j).Value
                Player(Index).Char(i).NewInv(j).Ammo = Player(Index).Char(i).Inv(j).Ammo
            Next
            
            ' State that the player's inventory was converted
            Player(Index).Char(i).InvConverted = 1
            
            ' Indicate that some data was modified
            WasConverted = True
        End If
            
        ReDim Preserve Player(Index).Char(i).NewInv(1 To Player(Index).Char(i).MaxInv) As NewPlayerInvRec
        
        ' Check if the player's bank was converted to the new system
        If Player(Index).Char(i).BankConverted = 0 Then
            Dim p As Integer
            
            ReDim Player(Index).Char(i).NewBank(1 To MAX_BANK) As NewBankRec
            
            ' Convert the player's bank from 50 slots to 100 slots
            For p = 1 To 50
                Player(Index).Char(i).NewBank(p).num = Player(Index).Char(i).Bank(p).num
                Player(Index).Char(i).NewBank(p).Value = Player(Index).Char(i).Bank(p).Value
                Player(Index).Char(i).NewBank(p).Ammo = Player(Index).Char(i).Bank(p).Ammo
            Next
            
            ' State that the player's bank was converted
            Player(Index).Char(i).BankConverted = 1
            
            ' Indicate that some data was modified
            WasConverted = True
        End If
        
        ReDim Preserve Player(Index).Char(i).NewBank(1 To MAX_BANK) As NewBankRec
    Next i
    
    ' Save the player's data if any of his/her characters were switched over to the new inventory system
    If WasConverted = True Then
        Call SavePlayer(Index)
    End If
    
    Exit Sub

PlayerErr:
    ' If these errors occur, it most likely means the player hasn't updated his/her client
    'Call MsgBox("Couldn't load index " & Index & " for " & Name & "!", vbCritical)
    'Call DestroyServer
End Sub

Function AccountExists(ByVal Name As String) As Boolean
    If FileExists("\SMBOAccounts\" & Trim$(Name) & "_Info.ini") Then
        AccountExists = True
    Else
        AccountExists = False
    End If
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim$(Player(Index).Char(CharNum).Name) <> vbNullString Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim RightPassword As String

    PasswordOK = False

    If AccountExists(Name) Then
        RightPassword = GetVar(App.Path & "\SMBOAccounts\" & Trim$(Name) & "_Info.ini", "ACCESS", "Password")

        If Trim$(Password) = Trim$(RightPassword) Then
            PasswordOK = True
        End If
    End If
End Function

Public Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long

    Player(Index).Login = Name
    Player(Index).Password = Password

    For i = 1 To MAX_CHARS
        Call ClearChar(Index, i)
    Next i

    Call SavePlayer(Index)
    
    Call ClearPlayer(Index)
End Sub

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
    Dim f As Long
    
    With Player(Index).Char(CharNum)
        If Trim$(.Name) = vbNullString Then
            Player(Index).CharNum = CharNum
    
            .Name = Name
            .Sex = Sex
            .Class = ClassNum
            .Sprite = ClassData(ClassNum).MaleSprite
    
            .LEVEL = 1
    
            .STR = ClassData(ClassNum).STR
            .DEF = ClassData(ClassNum).DEF
            .Speed = ClassData(ClassNum).Speed
            .Magi = ClassData(ClassNum).Magi
    
            .Map = ClassData(ClassNum).Map
            .X = ClassData(ClassNum).X
            .Y = ClassData(ClassNum).Y
    
            .HP = GetPlayerMaxHP(Index)
            .MP = GetPlayerMaxMP(Index)
            .SP = GetPlayerMaxSP(Index)
    
            .MAXHP = GetPlayerMaxHP(Index)
            .MAXMP = GetPlayerMaxMP(Index)
            .MAXSP = GetPlayerMaxSP(Index)
    
            .Head = 0
            .Body = 0
            .Leg = 0
            .TempSprite = ClassData(ClassNum).MaleSprite
            
            ' Clear out equipment
            For f = 1 To 7
                Call SetPlayerEquipSlotNum(Index, f, 0)
                Call SetPlayerEquipSlotValue(Index, f, 0)
                Call SetPlayerEquipSlotAmmo(Index, f, -1)
            Next f
            
        ' version info
            .Vflag = 128
            .Ver = 2
            .SubVer = 8
            .Rel = 0
            
            .PAPERDOLL = 1
            .Dir = DIR_DOWN
    
            ' Append name to file
            f = FreeFile
            Open App.Path & "\SMBOAccounts\CharList.txt" For Append As #f
            Print #f, Name
            Close #f
    
            Call SavePlayer(Index)
    
            Exit Sub
        End If
    End With
End Sub

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
    Dim CharName As String, StringIndex As String
    Dim i As Long
    
    CharName = Trim$(Player(Index).Char(CharNum).Name)
    
    ' **** Deletes stored INI data for character ****
    
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
    For i = 94 To MAX_ITEMS
        StringIndex = CStr(i)
            
        If GetVar(App.Path & "\Scripts\" & "Cards.ini", StringIndex, CharName) = "Has" Then
            Call PutVar(App.Path & "\Scripts\" & "Cards.ini", StringIndex, CharName, vbNullString)
        End If
    Next i
    
    ' Recipes
    For i = 1 To MAX_RECIPES
        If GetVar(App.Path & "\Scripts\" & "Recipes.ini", StringIndex, CharName) = "Has" Then
            Call PutVar(App.Path & "\Scripts\" & "Recipes.ini", StringIndex, CharName, vbNullString)
        End If
    Next i
    
    Call DeleteName(CharName)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim S As String

    FindChar = False

    f = FreeFile
    Open App.Path & "\SMBOAccounts\CharList.txt" For Input As #f
    Do While Not EOF(f)
        Input #f, S

        If Trim$(LCase$(S)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If
    Loop
    Close #f
End Function

Sub SaveAllPlayersOnline()
    Dim i As Integer

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next i
End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim i As Long

    Call ClearClasses

    For i = 0 To MAX_CLASSES

        On Error Resume Next ' used if next line tries to divide by 0

        temp = i / MAX_CLASSES * 100

        On Error GoTo ClassErr

        Call SetStatus("Loading classes... " & temp & "%")
        
        FileName = App.Path & "\SMBOClasses\Class" & i & ".ini"

        ClassData(i).Name = GetVar(FileName, "CLASS", "Name")
        ClassData(i).MaleSprite = GetVar(FileName, "CLASS", "MaleSprite")
        ClassData(i).FemaleSprite = GetVar(FileName, "CLASS", "FemaleSprite")
        ClassData(i).Desc = GetVar(FileName, "CLASS", "Desc")
        ClassData(i).STR = CLng(GetVar(FileName, "CLASS", "STR"))
        ClassData(i).DEF = CLng(GetVar(FileName, "CLASS", "DEF"))
        ClassData(i).Speed = CLng(GetVar(FileName, "CLASS", "SPEED"))
        ClassData(i).Magi = CLng(GetVar(FileName, "CLASS", "MAGI"))
        ClassData(i).Map = CLng(GetVar(FileName, "CLASS", "MAP"))
        ClassData(i).X = CByte(GetVar(FileName, "CLASS", "X"))
        ClassData(i).Y = CByte(GetVar(FileName, "CLASS", "Y"))
        ClassData(i).Locked = CLng(GetVar(FileName, "CLASS", "Locked"))
    Next i
    Exit Sub

ClassErr:
    Call MsgBox("Error loading class " & i & ". Check that all the variables in your class files exist!")
    Call DestroyServer
    End
End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim i As Long

    For i = 0 To MAX_CLASSES
        On Error Resume Next ' if MAX_CLASSES is 0
    
        temp = i / MAX_CLASSES * 100
        Call SetStatus("Saving classes... " & temp & "%")
        
        FileName = App.Path & "\SMBOClasses\Class" & i & ".ini"
        
        If Not FileExists("Classes\Class" & i & ".ini") Then
            Call PutVar(FileName, "CLASS", "Name", Trim$(ClassData(i).Name))
            Call PutVar(FileName, "CLASS", "MaleSprite", CStr(ClassData(i).MaleSprite))
            Call PutVar(FileName, "CLASS", "FemaleSprite", CStr(ClassData(i).FemaleSprite))
            Call PutVar(FileName, "CLASS", "STR", CStr(ClassData(i).STR))
            Call PutVar(FileName, "CLASS", "DEF", CStr(ClassData(i).DEF))
            Call PutVar(FileName, "CLASS", "SPEED", CStr(ClassData(i).Speed))
            Call PutVar(FileName, "CLASS", "MAGI", CStr(ClassData(i).Magi))
            Call PutVar(FileName, "CLASS", "MAP", CStr(ClassData(i).Map))
            Call PutVar(FileName, "CLASS", "X", CStr(ClassData(i).X))
            Call PutVar(FileName, "CLASS", "Y", CStr(ClassData(i).Y))
            Call PutVar(FileName, "CLASS", "Locked", CStr(ClassData(i).Locked))
        End If
    Next i
End Sub

Sub SaveItems()
    Dim i As Long

    Call SetStatus("Saving items... ")
    For i = 1 To MAX_ITEMS
        If Not FileExists("items\item" & i & ".dat") Then
            temp = i / MAX_ITEMS * 100
            Call SetStatus("Saving items... " & temp & "%")
            Call SaveItem(i)
        End If
    Next i
End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim FileName As String
    Dim f  As Long
    FileName = App.Path & "\items\item" & ItemNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Item(ItemNum)
    Close #f
End Sub

Function NewRecConvert(FileName As String) As ItemRec
    Dim OldRec As NewItemRec
    Dim NewRec As ItemRec
    Dim f As Long
    Dim n As Integer
    
     f = FreeFile
    Open FileName For Binary As #f
        Get #f, , OldRec
    Close #f
    
 ' Old Rec stuff
    NewRec.Name = OldRec.Name
    NewRec.Desc = OldRec.Desc
    NewRec.Pic = OldRec.Pic
    NewRec.Type = OldRec.Type
    NewRec.Data1 = OldRec.Data1
    NewRec.Data2 = OldRec.Data2
    NewRec.Data3 = OldRec.Data3
    NewRec.StrReq = OldRec.StrReq
    NewRec.DefReq = OldRec.DefReq
    NewRec.SpeedReq = OldRec.SpeedReq
    NewRec.MagicReq = OldRec.MagicReq
    NewRec.ClassReq = OldRec.ClassReq
    NewRec.AccessReq = OldRec.AccessReq
    NewRec.addHP = OldRec.addHP
    NewRec.addMP = OldRec.addMP
    NewRec.addSP = OldRec.addSP
    NewRec.AddStr = OldRec.AddStr
    NewRec.AddMagi = OldRec.AddMagi
    NewRec.AddSpeed = OldRec.AddSpeed
    NewRec.AddEXP = OldRec.AddEXP
    NewRec.AttackSpeed = OldRec.AttackSpeed
    NewRec.Price = OldRec.Price
    NewRec.Stackable = OldRec.Stackable
    NewRec.Bound = OldRec.Bound
    NewRec.LevelReq = OldRec.LevelReq
    NewRec.TwoHanded = OldRec.TwoHanded
    
    NewRecConvert = NewRec
End Function

Sub LoadItems()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckItems

    For i = 1 To MAX_ITEMS
        temp = i / MAX_ITEMS * 100
        Call SetStatus("Loading items... " & temp & "%")

        FileName = App.Path & "\Items\Item" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Item(i)
        Close #f
    Next i
End Sub

Sub CheckItems()
    Call SaveItems
End Sub

Function ConvertMapItem(ByVal Index As Long) As MapItemRec
    'Dim OldRec As NewMapItemRec
    'Dim NewRec As MapItemRec
    
    'NewRec.num = OldRec.num
    'NewRec.Value = OldRec.Value
    
    'NewRec.X = OldRec.X
    'NewRec.Y = OldRec.Y
    'NewRec.Ammo = OldRec.Ammo
    
    'ConvertMapItem = NewRec
End Function

Sub SaveShops()
    Dim i As Long

    Call SetStatus("Saving shops... ")
    For i = 1 To MAX_SHOPS
        If Not FileExists("shops\shop" & i & ".dat") Then
            temp = i / MAX_SHOPS * 100
            Call SetStatus("Saving shops... " & temp & "%")
            Call SaveShop(i)
        End If
    Next i
End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\shops\shop" & ShopNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim i As Long, f As Long

    Call CheckShops

    For i = 1 To MAX_SHOPS
        temp = i / MAX_SHOPS * 100
        Call SetStatus("Loading shops... " & temp & "%")
        FileName = App.Path & "\shops\shop" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Shop(i)
        Close #f

    Next i
End Sub

Sub CheckShops()
    Call SaveShops
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\spells\spells" & SpellNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
    Dim i As Long

    Call SetStatus("Saving spells... ")
    For i = 1 To MAX_SPELLS
        If Not FileExists("spells\spells" & i & ".dat") Then
            temp = i / MAX_SPELLS * 100
            Call SetStatus("Saving spells... " & temp & "%")
            Call SaveSpell(i)
        End If
    Next i
End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckSpells

    For i = 1 To MAX_SPELLS
        temp = i / MAX_SPELLS * 100
        Call SetStatus("Loading spells... " & temp & "%")

        FileName = App.Path & "\spells\spells" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Spell(i)
        Close #f

    Next i
End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub SaveRecipe(ByVal RecipeNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\Recipes\recipes" & RecipeNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Recipe(RecipeNum)
    Close #f
End Sub

Sub SaveRecipes()
    Dim i As Long

    Call SetStatus("Saving recipes... ")
    For i = 1 To MAX_RECIPES
        If Not FileExists("Recipes\recipes" & i & ".dat") Then
            temp = i / MAX_RECIPES * 100
            Call SetStatus("Saving recipes... " & temp & "%")
            Call SaveRecipe(i)
        End If
    Next i
End Sub

Sub LoadRecipes()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckRecipes

    For i = 1 To MAX_RECIPES
        temp = i / MAX_RECIPES * 100
        Call SetStatus("Loading recipes... " & temp & "%")

        FileName = App.Path & "\Recipes\recipes" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Recipe(i)
        Close #f

    Next i
End Sub

Sub CheckRecipes()
    Call SaveRecipes
End Sub

Sub SaveNpcs()
    Dim i As Long

    Call SetStatus("Saving npcs... ")

    For i = 1 To MAX_NPCS
        If Not FileExists("npcs\npc" & i & ".dat") Then
            temp = i / MAX_NPCS * 100
            Call SetStatus("Saving npcs... " & temp & "%")
            Call SaveNpc(i)
        End If
    Next i
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim FileName As String
    Dim f As Long
    FileName = App.Path & "\npcs\npc" & NpcNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , NPC(NpcNum)
    Close #f
End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim i As Integer
    Dim f As Long

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        temp = i / MAX_NPCS * 100
        Call SetStatus("Loading npcs... " & temp & "%")
        FileName = App.Path & "\npcs\npc" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , NPC(i)
        Close #f
    Next i
End Sub

Sub CheckNpcs()
    Call SaveNpcs
End Sub

Sub SaveQuestionBlocks(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    With QuestionBlock(MapNum, X, Y)
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Item1", CStr(.Item1))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Item2", CStr(.Item2))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Item3", CStr(.Item3))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Item4", CStr(.Item4))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Item5", CStr(.Item5))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Item6", CStr(.Item6))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Chance1", CStr(.Chance1))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Chance2", CStr(.Chance2))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Chance3", CStr(.Chance3))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Chance4", CStr(.Chance4))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Chance5", CStr(.Chance5))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Chance6", CStr(.Chance6))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Value1", CStr(.Value1))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Value2", CStr(.Value2))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Value3", CStr(.Value3))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Value4", CStr(.Value4))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Value5", CStr(.Value5))
        Call PutVar(App.Path & "\QuestionBlockData.ini", "Map: " & MapNum & "/X: " & X & "/Y: " & Y, "Value6", CStr(.Value6))
    End With
End Sub

Sub LoadQuestionBlocks()
    Dim i As Long, X As Long, Y As Long
    
    For i = 1 To MAX_MAPS
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                ' The ? Block exists
                If GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Item1") <> vbNullString Then
                    With QuestionBlock(i, X, Y)
                        .Item1 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Item1"))
                        .Item2 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Item2"))
                        .Item3 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Item3"))
                        .Item4 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Item4"))
                        .Item5 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Item5"))
                        .Item6 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Item6"))
                        .Chance1 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Chance1"))
                        .Chance2 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Chance2"))
                        .Chance3 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Chance3"))
                        .Chance4 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Chance4"))
                        .Chance5 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Chance5"))
                        .Chance6 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Chance6"))
                        .Value1 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Value1"))
                        .Value2 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Value2"))
                        .Value3 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Value3"))
                        .Value4 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Value4"))
                        .Value5 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Value5"))
                        .Value6 = CLng(GetVar(App.Path & "\QuestionBlockData.ini", "Map: " & i & "/X: " & X & "/Y: " & Y, "Value6"))
                    End With
                End If
            Next X
        Next Y
    Next i
End Sub

Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map(MapNum)
    Close #f
      
End Sub

Sub LoadMaps()
    Dim FileName As String
    Dim i As Long
    Dim f As Integer

    Call CheckMaps

    For i = 1 To MAX_MAPS
        temp = i / MAX_MAPS * 100
        Call SetStatus("Loading maps... " & temp & "%")
        FileName = App.Path & "\maps\map" & i & ".dat"

        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Map(i)
        Close #f
    Next i

End Sub

Sub CheckMaps()
    Dim FileName As String
    Dim i As Long

    Call ClearMaps

    For i = 1 To MAX_MAPS
        FileName = "maps\map" & i & ".dat"

        ' Check to see if map exists. If it doesn't, create it.
        If Not FileExists(FileName) Then
            temp = i / MAX_MAPS * 100
            Call SetStatus("Saving maps... " & temp & "%")
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub BanPlayer(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim FileName As String, NumString As String
    Dim i As Integer

    FileName = App.Path & "\SMBOBanList.ini"

    For i = 1 To 100
        NumString = CStr(i)
        
        If GetVar(FileName, "Ban List", NumString) = "" Then
            ' Add the player to the ban list INI file
            Call PutVar(FileName, "Ban List", NumString, GetPlayerIP(BanPlayerIndex) & " - " & Trim$(Player(BanPlayerIndex).Login) & "/" & GetPlayerName(BanPlayerIndex))
            Exit For
        End If
    Next
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from Super Mario Bros. Online by " & GetPlayerName(BannedByIndex) & "!", WHITE)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub UnbanPlayer(ByVal Index As Long, ByVal BanListNum As Integer)
    Dim FileName As String

    FileName = App.Path & "\SMBOBanList.ini"
    
    Call PutVar(FileName, "Ban List", CStr(BanListNum), "")
        
    If Index > 0 Then
        Call PlayerMsg(Index, "You have successfully removed entry #" & BanListNum & " from the ban list!", YELLOW)
    End If
End Sub

Public Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim FileName As String
    Dim FileID As Long

    If ServerLog Then
        FileName = App.Path & "\" & FN

        If FileExists(FN) Then
            FileID = FreeFile
            Open FileName For Output As #FileID
            Print #FileID, Time & ": " & Text
            Close #FileID
        End If
    End If
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long, f2 As Long
    Dim S As String

    Call FileCopy(App.Path & "\SMBOAccounts\CharList.txt", App.Path & "\SMBOAccounts\chartemp.txt")

    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\SMBOAccounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\SMBOAccounts\CharList.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, S
        If Trim$(LCase$(S)) <> Trim$(LCase$(Name)) Then
            Print #f2, S
        End If
    Loop

    Close #f1
    Close #f2

    Call Kill(App.Path & "\SMBOAccounts\chartemp.txt")
End Sub

Sub BanByServer(ByVal Index As Long, ByVal Reason As String)
    Dim FileName As String, NumString As String
    Dim i As Integer
    
    If IsPlaying(Index) Then
        FileName = App.Path & "\SMBOBanList.ini"

        For i = 1 To 100
            NumString = CStr(i)
        
            If GetVar(FileName, "Ban List", NumString) = "" Then
                ' Add the player to the ban list INI file
                Call PutVar(FileName, "Ban List", NumString, GetPlayerIP(Index) & " - " & Trim$(Player(Index).Login) & "/" & GetPlayerName(Index))
                Exit For
            End If
        Next

        If LenB(Reason) = 0 Then
            Reason = "No reason given."
        End If
            
        Call GlobalMsg(GetPlayerName(Index) & " has been banned by the server! Reason: (" & Reason & ")", WHITE)
        Call AddLog("The server has banned " & GetPlayerName(Index) & ". Reason: (" & Reason & ")", ADMIN_LOG)
        Call AlertMsg(Index, "You have been banned by the server! Reason: (" & Reason & ")")
    End If
End Sub

Sub SaveLogs()
    Dim FileName As String, CurDate As String, CurTime As String
    Dim FileID As Integer

    On Error Resume Next

    If Not FolderExists(App.Path & "\Logs") Then
        Call MkDir(App.Path & "\Logs")
    End If

    'CurDate = Date
    CurDate = Replace(Date, "/", "-")
    CurTime = Replace(Time, ":", "-")

    If Not FolderExists(App.Path & "\Logs\" & CurDate) Then
        Call MkDir(App.Path & "\Logs\" & CurDate)
    End If

    Call MkDir(App.Path & "\Logs\" & CurDate & "/" & CurTime)
    FileID = FreeFile

    FileName = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Main.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(0).Text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Broadcast.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(1).Text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Global.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(2).Text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Map.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(3).Text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Private.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(4).Text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Admin.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(5).Text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Emote.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(6).Text
    Close #FileID
End Sub

Sub LoadArrows()
    Dim FileName As String
    Dim i As Long

    Call CheckArrows

    FileName = App.Path & "\Arrows.ini"

    For i = 1 To MAX_ARROWS
        temp = i / MAX_ARROWS * 100
        Call SetStatus("Loading Arrows... " & temp & "%")
        Arrows(i).Name = GetVar(FileName, "Arrow" & i, "ArrowName")
        Arrows(i).Pic = GetVar(FileName, "Arrow" & i, "ArrowPic")
        Arrows(i).Range = GetVar(FileName, "Arrow" & i, "ArrowRange")
        Arrows(i).Amount = GetVar(FileName, "Arrow" & i, "ArrowAmount")

    Next i
End Sub

Sub CheckArrows()
    If Not FileExists("Arrows.ini") Then
        Dim i As Long

        For i = 1 To MAX_ARROWS
            temp = i / MAX_ARROWS * 100
            Call SetStatus("Saving arrows... " & temp & "%")

            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowName", vbNullString)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowPic", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowRange", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowAmount", 0)
        Next i
    End If
End Sub

Sub ClearArrows()
    Dim i As Long

    For i = 1 To MAX_ARROWS
        Arrows(i).Name = vbNullString
        Arrows(i).Pic = 0
        Arrows(i).Range = 0
        Arrows(i).Amount = 0
    Next i
End Sub

Sub SaveArrow(ByVal ArrowNum As Long)
    Dim FileName As String

    FileName = App.Path & "\Arrows.ini"

    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowName", Trim$(Arrows(ArrowNum).Name))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowAmount", Val(Arrows(ArrowNum).Amount))
End Sub

Sub ClearTempTile()
    Dim i As Long, Y As Long, X As Long

    For i = 1 To MAX_MAPS
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                TempTile(i).DoorTimer(X, Y) = 0
                TempTile(i).DoorOpen(X, Y) = NO
            Next X
        Next Y
    Next i
End Sub

Sub ClearClasses()
    Dim i As Long

    For i = 0 To MAX_CLASSES
        With ClassData(i)
            .Name = vbNullString
            .AdvanceFrom = 0
            .LevelReq = 0
            .Type = 1
            .STR = 0
            .DEF = 0
            .Speed = 0
            .Magi = 0
            .FemaleSprite = 0
            .MaleSprite = 0
            .Desc = vbNullString
            .Map = 0
            .X = 0
            .Y = 0
        End With
    Next i
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long, n As Long
    Dim FileName As String
    
    FileName = App.Path & "\SMBOAccounts\" & Trim$(Player(Index).Login) & "_Info.ini"
    
    With Player(Index)
        .Login = vbNullString
        .Password = vbNullString
        .Locked = False
        .Mute = False
        .LockedSpells = False
        .LockedItems = False
        .LockedAttack = False
    
        ' Temporary vars
        .Buffer = vbNullString
        .IncBuffer = vbNullString
        .CharNum = 0
        .InGame = False
        .AttackTimer = 0
        .DataTimer = 0
        .DataBytes = 0
        .DataPackets = 0
        .Target = 0
        .TargetType = 0
        .CastedSpell = NO
        .GettingMap = NO
        .Emoticon = -1
        .InTrade = False
        .TradePlayer = 0
        .TradeOk = 0
        
        For i = 1 To MAX_PLAYER_TRADES
            .Trades(i).InvName = vbNullString
            .Trades(i).InvNum = 0
            .Trades(i).InvVal = 0
        Next i
        
        .ChatPlayer = 0
    End With
    
    For i = 1 To MAX_CHARS
        With Player(Index).Char(i)
            .Name = vbNullString
            .Class = 0
            .LEVEL = 0
            .Sprite = 0
            .Exp = 0
            .Access = 0
            .PK = NO
            .POINTS = 0
            .Guild = vbNullString
    
            .HP = 0
            .MP = 0
            .SP = 0
    
            .MAXHP = 0
            .MAXMP = 0
            .MAXSP = 0
    
            .STR = 0
            .DEF = 0
            .Speed = 0
            .Magi = 0
            
            For n = 1 To MAX_INV
                .Inv(n).num = 0
                .Inv(n).Value = 0
                .Inv(n).Ammo = -1
            Next n
            
            For n = 1 To 50
                .Bank(n).num = 0
                .Bank(n).Value = 0
                .Bank(n).Ammo = -1
            Next n
            
            For n = 1 To MAX_PLAYER_SPELLS
                .Spell(n) = 0
            Next n
    
            .ArmorSlot = 0
            .WeaponSlot = 0
            .HelmetSlot = 0
            .ShieldSlot = 0
            .LegsSlot = 0
            .RingSlot = 0
            .NecklaceSlot = 0
    
            .Map = 0
            .X = 0
            .Y = 0
            .Dir = 0
            .InBattle = False
            .Turn = False
            .OldX = 0
            .OldY = 0
            .HasTurnBased = False
            .RecoverTime = 0
            .PartyNum = 0
            .PartyInvitedBy = 0
            .Height = 0
            
            For n = 1 To 7
                .Equipment(n).num = 0
                .Equipment(n).Value = 0
                .Equipment(n).Ammo = -1
            Next n
            
            .TempSprite = 0
            
            If .MaxInv < 1 Then
                .MaxInv = 24
            End If
            
            ReDim .NewInv(1 To .MaxInv) As NewPlayerInvRec
            
            For n = 1 To .MaxInv
                .NewInv(n).num = 0
                .NewInv(n).Value = 0
                .NewInv(n).Ammo = -1
            Next
            
            .InvConverted = 0
            
            ReDim .NewBank(1 To MAX_BANK) As NewBankRec
            
            For n = 1 To MAX_BANK
                .NewBank(n).num = 0
                .NewBank(n).Value = 0
                .NewBank(n).Ammo = -1
            Next
            
            .BankConverted = 0
        End With
    Next i
End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
    Dim n As Long
    
    With Player(Index).Char(CharNum)
        ' version info
        .Vflag = 128
        .Ver = 2
        .SubVer = 8
        .Rel = 0
        
        .Name = vbNullString
        .Class = 0
        .Sprite = 0
        .LEVEL = 0
        .Exp = 0
        .Access = 0
        .PK = NO
        .POINTS = 0
        .Guild = vbNullString
    
        .HP = 0
        .MP = 0
        .SP = 0
    
        .MAXHP = 0
        .MAXMP = 0
        .MAXSP = 0
    
        .STR = 0
        .DEF = 0
        .Speed = 0
        .Magi = 0
    
        For n = 1 To MAX_INV
            .Inv(n).num = 0
            .Inv(n).Value = 0
            .Inv(n).Ammo = -1
        Next n
        
        For n = 1 To 50
            .Bank(n).num = 0
            .Bank(n).Value = 0
            .Bank(n).Ammo = -1
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            .Spell(n) = 0
        Next n
        
        If .MaxInv < 1 Then
            .MaxInv = 24
        End If
    
        ReDim .NewInv(1 To .MaxInv) As NewPlayerInvRec
        
        For n = 1 To .MaxInv
            .NewInv(n).num = 0
            .NewInv(n).Value = 0
            .NewInv(n).Ammo = -1
        Next n
        
        ReDim .NewBank(1 To MAX_BANK) As NewBankRec
        
        For n = 1 To MAX_BANK
            .NewBank(n).num = 0
            .NewBank(n).Value = 0
            .NewBank(n).Ammo = -1
        Next n
        
        .ArmorSlot = 0
        .WeaponSlot = 0
        .HelmetSlot = 0
        .ShieldSlot = 0
        .LegsSlot = 0
        .RingSlot = 0
        .NecklaceSlot = 0
    
        .Map = 0
        .X = 0
        .Y = 0
        .Dir = 0
        .InBattle = False
        .Turn = False
        .OldX = 0
        .OldY = 0
        .HasTurnBased = False
        .RecoverTime = 0
    End With
End Sub

Sub ClearItem(ByVal Index As Long)
    With Item(Index)
        .Name = vbNullString
        .Desc = vbNullString
    
        .Type = 0
        .Data1 = 0
        .Data2 = 0
        .Data3 = 0
        .StrReq = 0
        .DefReq = 0
        .SpeedReq = 0
        .MagicReq = 0
        .ClassReq = -1
        .AccessReq = 0
        .LevelReq = 0
    
        .addHP = 0
        .addMP = 0
        .addSP = 0
        .AddStr = 0
        .AddDef = 0
        .AddMagi = 0
        .AddSpeed = 0
        .AddEXP = 0
        .AttackSpeed = 1000
        .Price = 0
        .Stackable = 0
        .Bound = 0
    End With
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearNpc(ByVal Index As Long)
    Dim i As Long
    
    With NPC(Index)
        .Name = vbNullString
        .AttackSay = vbNullString
        .Sprite = 0
        .SpawnSecs = 0
        .Behavior = 0
        .Range = 0
        .STR = 0
        .DEF = 0
        .Speed = 0
        .Magi = 0
        .Big = 0
        .MAXHP = 0
        .Exp = 0
        .SpawnTime = 0
        .Element = 0
    
        For i = 1 To MAX_NPC_DROPS
            .ItemNPC(i).Chance = 0
            .ItemNPC(i).ItemNum = 0
            .ItemNPC(i).ItemValue = 0
        Next i
    
        .AttackSay2 = vbNullString
        .LEVEL = 0
    End With
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    MapItem(MapNum, Index).num = 0
    MapItem(MapNum, Index).Value = 0
    MapItem(MapNum, Index).X = 0
    MapItem(MapNum, Index).Y = 0
    MapItem(MapNum, Index).Ammo = -1
    Call SendDataToMap(MapNum, SPackets.Sspawnitem & SEP_CHAR & Index & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & -1 & END_CHAR)
End Sub

Sub ClearMapItems()
    Dim X As Long, Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next X
    Next Y
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    With MapNPC(MapNum, Index)
        .num = 0
        .Target = 0
        .HP = 0
        .MP = 0
        .SP = 0
        .X = 0
        .Y = 0
        .Dir = 0
    
        ' Server use only
        .SpawnWait = 0
        .AttackTimer = 0
        .InBattle = False
        .Turn = False
        .OldX = 0
        .OldY = 0
    End With
End Sub

Sub ClearMapNpcs()
    Dim X As Long, Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next X
    Next Y
End Sub

Function NewNpcRecConvert(FileName As String) As NpcRec
    Dim OldRec As NewNpcRec
    Dim NewRec As NpcRec
    Dim f As Long, n As Long
    
     f = FreeFile
    Open FileName For Binary As #f
        Get #f, , OldRec
    Close #f
    
    ' Old Rec stuff
    NewRec.Name = OldRec.Name
    NewRec.AttackSay = OldRec.AttackSay
    NewRec.Sprite = OldRec.Sprite
    NewRec.SpawnSecs = OldRec.SpawnSecs
    NewRec.Behavior = OldRec.Behavior
    NewRec.Range = OldRec.Range
    NewRec.STR = OldRec.STR
    NewRec.DEF = OldRec.DEF
    NewRec.Speed = OldRec.Speed
    NewRec.Magi = OldRec.Magi
    NewRec.Big = OldRec.Big
    NewRec.MAXHP = OldRec.MAXHP
    NewRec.Exp = OldRec.Exp
    NewRec.SpawnTime = OldRec.SpawnTime
    
    For n = 1 To MAX_NPC_DROPS
        NewRec.ItemNPC(n) = OldRec.ItemNPC(n)
    Next n
    
    NewRec.Element = OldRec.Element
    NewRec.SPRITESIZE = OldRec.SPRITESIZE
    
    NewNpcRecConvert = NewRec
End Function

Sub ClearMap(ByVal MapNum As Long)
    Dim X As Long, Y As Long
    
    With Map(MapNum)
        .Name = vbNullString
        .Revision = 0
        .Moral = 0
        .Up = 0
        .Down = 0
        .Left = 0
        .Right = 0
        .Indoors = 0
        .Weather = 0
    
        For X = 1 To MAX_MAP_NPCS
            .NPC(X) = 0
        Next X
    End With

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map(MapNum).Tile(X, Y)
                .Ground = 0
                .Mask = 0
                .Anim = 0
                .Mask2 = 0
                .M2Anim = 0
                .Fringe = 0
                .FAnim = 0
                .Fringe2 = 0
                .F2Anim = 0
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
                .String1 = vbNullString
                .String2 = vbNullString
                .String3 = vbNullString
                .Light = 0
                .GroundSet = 0
                .MaskSet = 0
                .AnimSet = 0
                .Mask2Set = 0
                .M2AnimSet = 0
                .FringeSet = 0
                .FAnimSet = 0
                .Fringe2Set = 0
                .F2AnimSet = 0
            End With
            
            With QuestionBlock(MapNum, X, Y)
                .Item1 = 0
                .Item2 = 0
                .Item3 = 0
                .Item4 = 0
                .Item5 = 0
                .Item6 = 0
                .Chance1 = 0
                .Chance2 = 0
                .Chance3 = 0
                .Chance4 = 0
                .Chance5 = 0
                .Chance6 = 0
                .Value1 = 0
                .Value2 = 0
                .Value3 = 0
                .Value4 = 0
                .Value5 = 0
                .Value6 = 0
            End With
        Next X
    Next Y

    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO

    ' Reset the map cache array for this map.
    MapCache(MapNum) = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i
End Sub

Sub ClearShop(ByVal Index As Long)
    Dim i As Long
    
    With Shop(Index)
        .Name = vbNullString
        .CurrencyItem = 1
        .FixesItems = 0
        .ShowInfo = 0
        
        For i = 1 To MAX_SHOP_ITEMS
            .ShopItem(i).ItemNum = 0
            .ShopItem(i).Amount = 0
            .ShopItem(i).Price = 0
        Next i
    End With
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
End Sub

Sub ClearSpell(ByVal Index As Long)
    With Spell(Index)
        .Name = vbNullString
        .ClassReq = 0
        .LevelReq = 0
        .Type = 0
        .Data1 = 0
        .Data2 = 0
        .Data3 = 0
        .MPCost = 0
        .Sound = 0
        .Range = 0
    
        .SpellAnim = 0
        .SpellTime = 40
        .SpellDone = 1
    
        .AE = 0
        .Big = 0
    
        .Element = 0
    End With
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i
End Sub

Sub ClearRecipe(ByVal Index As Long)
    Recipe(Index).Ingredient1 = 0
    Recipe(Index).Ingredient2 = 0
    Recipe(Index).ResultItem = 0
    Recipe(Index).Name = vbNullString
End Sub

Sub ClearRecipes()
    Dim i As Long

    For i = 1 To MAX_RECIPES
        Call ClearRecipe(i)
    Next i
End Sub

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).GuildAccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal GuildAccess As Long)
    Player(Index).Char(Player(Index).CharNum).GuildAccess = GuildAccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClassData(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    If Index > 0 And Index <= MAX_PLAYERS Then
        Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
    End If
End Sub

Function GetPlayerTempSprite(ByVal Index As Long) As Long
    GetPlayerTempSprite = Player(Index).Char(Player(Index).CharNum).TempSprite
End Function

Sub SetPlayerTempSprite(ByVal Index As Long, ByVal Sprite As Long)
    If Index > 0 And Index <= MAX_PLAYERS Then
        Player(Index).Char(Player(Index).CharNum).TempSprite = Sprite
    End If
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).LEVEL
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal LEVEL As Long)
    Player(Index).Char(Player(Index).CharNum).LEVEL = LEVEL
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
    
    If GetPlayerHP > GetPlayerMaxHP(Index) Then
        GetPlayerHP = GetPlayerMaxHP(Index)
    End If
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    With Player(Index).Char(Player(Index).CharNum)
        .HP = HP

        If .HP > GetPlayerMaxHP(Index) Then
            .HP = GetPlayerMaxHP(Index)
        ElseIf .HP < 0 Then
            .HP = 0
        End If
    End With
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
    
    If GetPlayerMP > GetPlayerMaxMP(Index) Then
        GetPlayerMP = GetPlayerMaxMP(Index)
    End If
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    With Player(Index).Char(Player(Index).CharNum)
        .MP = MP

        If .MP > GetPlayerMaxMP(Index) Then
            .MP = GetPlayerMaxMP(Index)
        ElseIf .MP < 0 Then
            .MP = 0
        End If
    End With
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    With Player(Index).Char(Player(Index).CharNum)
        .SP = SP

        If .SP > GetPlayerMaxSP(Index) Then
            .SP = GetPlayerMaxSP(Index)
        ElseIf .SP < 0 Then
            .SP = 0
        End If
    End With
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    Dim i As Long, CharNum As Long, Add As Long, LvlUp As Long, ItemNum As Long
    
    LvlUp = Val(GetVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "HP")) * 5
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).addHP
        End If
    Next i

    CharNum = Player(Index).CharNum
    
    If GetPlayerClass(Index) = 1 Or GetPlayerClass(Index) = 3 Or GetPlayerClass(Index) = 5 Then
        GetPlayerMaxHP = 15 + Add + LvlUp
    Else
        GetPlayerMaxHP = 10 + Add + LvlUp
    End If
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    Dim i As Long, CharNum As Long, Add As Long, LvlUp As Long, ItemNum As Long
    
    LvlUp = Val(GetVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "FP")) * 5
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).addMP
        End If
    Next i

    CharNum = Player(Index).CharNum
    
    If GetPlayerClass(Index) = 4 Then
        GetPlayerMaxMP = 10 + Add + LvlUp
    Else
        GetPlayerMaxMP = 5 + Add + LvlUp
    End If
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    Dim i As Long, Add As Long, CharNum As Long, ItemNum As Long
    
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).addSP
        End If
    Next i

    CharNum = Player(Index).CharNum
    
    If GetPlayerClass(Index) = 0 Then
        GetPlayerMaxSP = 35 + (GetPlayerLevel(Index) * addSP.LEVEL) + (GetPlayerSPEED(Index) * addSP.Speed) + Add
    Else
        GetPlayerMaxSP = 30 + (GetPlayerLevel(Index) * addSP.LEVEL) + (GetPlayerSPEED(Index) * addSP.Speed) + Add
    End If
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(ClassData(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    If ClassNum = 1 Or ClassNum = 3 Or ClassNum = 5 Then
        GetClassMaxHP = 15
    Else
        GetClassMaxHP = 10
    End If
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    If ClassNum <> 4 Then
        GetClassMaxMP = 5
    Else
        GetClassMaxMP = 10
    End If
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    If ClassNum <> 0 Then
        GetClassMaxSP = 30 + addSP.LEVEL + (ClassData(ClassNum).Speed * addSP.Speed)
    Else
        GetClassMaxSP = 35 + addSP.LEVEL + (ClassData(ClassNum).Speed * addSP.Speed)
    End If
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = ClassData(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = ClassData(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = ClassData(ClassNum).Speed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = ClassData(ClassNum).Magi
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    Dim i As Long, Add As Long, ItemNum As Long
    
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).AddStr
        End If
    Next i
    
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR + Add
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    Dim i As Long, Add As Long, ItemNum As Long
    
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).AddDef
        End If
    Next i
    
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF + Add
    
    If GetPlayerDEF < 0 Then
        GetPlayerDEF = 0
    End If
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    Dim i As Long, Add As Long, ItemNum As Long
    
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).AddSpeed
        End If
    Next i
    
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed + Add
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Function GetPlayerStache(ByVal Index As Long) As Long
    Dim i As Long, Add As Long, ItemNum As Long
    
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).AddMagi
        End If
    Next i
    
    GetPlayerStache = Player(Index).Char(Player(Index).CharNum).Magi + Add
End Function

Sub SetPlayerStache(ByVal Index As Long, ByVal Stache As Long)
    Player(Index).Char(Player(Index).CharNum).Magi = Stache
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Sub SetPlayerMaxSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).MAXSTR = STR
End Sub

Sub SetPlayerMaxDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).MAXDEF = DEF
End Sub

Sub SetPlayerMaxSpeed(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).MAXSpeed = Speed
End Sub

Sub SetPlayerMaxStache(ByVal Index As Long, ByVal Stache As Long)
    Player(Index).Char(Player(Index).CharNum).MAXStache = Stache
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    If Index > 0 And Index <= MAX_PLAYERS Then
        GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
    End If
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(Player(Index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char(Player(Index).CharNum).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerMaxInv(ByVal Index As Long) As Integer
    GetPlayerMaxInv = Player(Index).Char(Player(Index).CharNum).MaxInv
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If InvSlot > 0 Then
        GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).NewInv(InvSlot).num
    End If
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).NewInv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).NewInv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).NewInv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Byte)
    Call SendDataTo(Index, SPackets.Sdamagedisplay & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub MapBattleMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Byte)
    Call SendDataToMap(MapNum, SPackets.Sdamagedisplay & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).Char(Player(Index).CharNum).NewBank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).NewBank(BankSlot).num = ItemNum
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Char(Player(Index).CharNum).NewBank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).NewBank(BankSlot).Value = ItemValue
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetSpellReqLevel(ByVal SpellNum As Long) As Long
    GetSpellReqLevel = Spell(SpellNum).LevelReq
End Function

Function GetPlayerTargetNpc(ByVal Index As Long) As Long
    GetPlayerTargetNpc = Player(Index).TargetNPC
End Function

Function GetPlayerInvItemAmmo(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemAmmo = Player(Index).Char(Player(Index).CharNum).NewInv(InvSlot).Ammo
End Function

Sub SetPlayerInvItemAmmo(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemAmmo As Long)
    Player(Index).Char(Player(Index).CharNum).NewInv(InvSlot).Ammo = ItemAmmo
End Sub

Function GetPlayerBankItemAmmo(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemAmmo = Player(Index).Char(Player(Index).CharNum).NewBank(BankSlot).Ammo
End Function

Sub SetPlayerBankItemAmmo(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemAmmo As Long)
    Player(Index).Char(Player(Index).CharNum).NewBank(BankSlot).Ammo = ItemAmmo
End Sub

Sub SetPlayerTeam(ByVal Index As Long, ByVal Team As Byte, ByVal FilePath As String, ByVal MaxValue As Byte)
  Dim PlayerNum As String, TeamNameBlue As String, TeamNameRed As String, RedMsg As String, BlueMsg As String
  Dim i As Byte, TeamNumberRed As Byte, TeamNumberBlue As Byte

    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If

    ' Red = Team 0; Blue = Team 1
    
    ' Get the correct team name based on the minigame being played
    If FilePath = HideNSneakPath Then
        TeamNameRed = "Hiders"
        TeamNameBlue = "Seekers"
        
        RedMsg = "You are a Hider!"
        BlueMsg = "You are a Seeker!"
    Else
        TeamNameRed = "Red"
        TeamNameBlue = "Blue"
        
        RedMsg = "You are on the Red team!"
        BlueMsg = "You are on the Blue team!"
    End If
    
    ' Sets number of team members for the Blue Team
    If GetVar(FilePath, "Team", TeamNameBlue) = vbNullString Then
        TeamNumberBlue = 0
    Else
        TeamNumberBlue = CByte(GetVar(FilePath, "Team", TeamNameBlue))
    End If

    ' Sets number of team members for the Red Team
    If GetVar(FilePath, "Team", TeamNameRed) = vbNullString Then
        TeamNumberRed = 0
    Else
        TeamNumberRed = CByte(GetVar(FilePath, "Team", TeamNameRed))
    End If
    
    If Team = 0 And TeamNumberRed < MaxValue Then
        For i = 1 To MaxValue
            PlayerNum = CStr(i)
            
            If GetVar(FilePath, TeamNameRed, PlayerNum) = vbNullString Then
                Call PutVar(FilePath, TeamNameRed, PlayerNum, GetPlayerName(Index))
                Call PutVar(FilePath, "Team", TeamNameRed, (TeamNumberRed + 1))
                Call PlayerMsg(Index, RedMsg, WHITE)
                Exit Sub
            End If
        Next i
    ElseIf Team = 1 And TeamNumberBlue < MaxValue Then
        For i = 1 To MaxValue
            PlayerNum = CStr(i)
            
            If GetVar(FilePath, TeamNameBlue, PlayerNum) = vbNullString Then
                Call PutVar(FilePath, TeamNameBlue, PlayerNum, GetPlayerName(Index))
                Call PutVar(FilePath, "Team", TeamNameBlue, (TeamNumberBlue + 1))
                Call PlayerMsg(Index, BlueMsg, WHITE)
                Exit Sub
            End If
        Next i
    End If

    ' Handle the random doors
    If Team = 0 And TeamNumberRed >= MaxValue And TeamNumberBlue < MaxValue Then
        For i = 1 To MaxValue
            PlayerNum = CStr(i)
            
            If GetVar(FilePath, TeamNameBlue, PlayerNum) = vbNullString Then
                Call PutVar(FilePath, TeamNameBlue, PlayerNum, GetPlayerName(Index))
                Call PutVar(FilePath, "Team", TeamNameBlue, (TeamNumberBlue + 1))
                Call PlayerMsg(Index, BlueMsg, WHITE)
                Exit For
            End If
        Next i
    ElseIf Team = 1 And TeamNumberBlue >= MaxValue And TeamNumberRed < MaxValue Then
        For i = 1 To MaxValue
            PlayerNum = CStr(i)
        
            If GetVar(FilePath, TeamNameRed, PlayerNum) = vbNullString Then
                Call PutVar(FilePath, TeamNameRed, PlayerNum, GetPlayerName(Index))
                Call PutVar(FilePath, "Team", TeamNameRed, (TeamNumberRed + 1))
                Call PlayerMsg(Index, RedMsg, WHITE)
                Exit For
            End If
        Next i
    End If
End Sub

Function FindPlayerMinigame(ByVal Index As Long) As String
    Select Case GetPlayerMap(Index)
        ' STS
        Case 33
            FindPlayerMinigame = STSPath
        ' Dodgebill
        Case 188
            FindPlayerMinigame = DodgeBillPath
        ' Hide n' Sneak
        Case 271, 272, 273
            FindPlayerMinigame = HideNSneakPath
    End Select
End Function

Function GetPlayerTeam(ByVal Index As Long, ByVal FilePath As String, ByVal MaxValue As Byte) As Byte
    Dim PlayerNum As String, TeamNameRed As String, TeamNameBlue As String
    Dim i As Byte
  
    ' Get the correct team name based on the minigame being played
    If FilePath = HideNSneakPath Then
        TeamNameRed = "Hider"
        TeamNameBlue = "Seeker"
    Else
        TeamNameRed = "Red"
        TeamNameBlue = "Blue"
    End If
  
    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        GetPlayerTeam = 2
        Exit Function
    End If
    
    For i = 1 To MaxValue
        PlayerNum = CStr(i)
        
        If GetVar(FilePath, TeamNameBlue, PlayerNum) = GetPlayerName(Index) Then
            GetPlayerTeam = 1
            Exit Function
        ElseIf GetVar(FilePath, TeamNameRed, PlayerNum) = GetPlayerName(Index) Then
            GetPlayerTeam = 0
            Exit Function
        End If
    Next i
  
    GetPlayerTeam = 2
End Function

Function GetNpcLevel(ByVal NpcNum As Long) As Long
    GetNpcLevel = NPC(NpcNum).LEVEL
End Function

Function ItemIsStackable(ByVal ItemNum As Long) As Boolean
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        ItemIsStackable = True
    Else
        ItemIsStackable = False
    End If
End Function

Function GetPlayerCritHitChance(ByVal Index As Long) As Double
    Dim i As Byte
    Dim ItemNum As Long
    Dim Add As Double
    
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).AddCritChance
        End If
    Next i
    
    Player(Index).Char(Player(Index).CharNum).CritHitChance = (GetPlayerStache(Index) / 2.5)
    
    GetPlayerCritHitChance = Player(Index).Char(Player(Index).CharNum).CritHitChance + Add
End Function

Sub SetPlayerCritHitChance(ByVal Index As Long, ByVal CritHitChance As Double)
    Player(Index).Char(Player(Index).CharNum).CritHitChance = CritHitChance
End Sub

Function GetPlayerBlockChance(ByVal Index As Long) As Double
    Dim i As Byte
    Dim ItemNum As Long
    Dim Add As Double
    
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            Add = Add + Item(ItemNum).AddBlockChance
        End If
    Next i
    
    Player(Index).Char(Player(Index).CharNum).BlockChance = (GetPlayerSPEED(Index) / 2.5)
    
    GetPlayerBlockChance = Player(Index).Char(Player(Index).CharNum).BlockChance + Add
End Function

Sub SetPlayerBlockChance(ByVal Index As Long, ByVal BlockChance As Double)
    Player(Index).Char(Player(Index).CharNum).BlockChance = BlockChance
End Sub

Function GetPlayerInBattle(ByVal Index As Long) As Boolean
    GetPlayerInBattle = Player(Index).Char(Player(Index).CharNum).InBattle
End Function

Sub SetPlayerInBattle(ByVal Index As Long, ByVal Battle As Boolean)
    Player(Index).Char(Player(Index).CharNum).InBattle = Battle
End Sub

Function GetPlayerTurn(ByVal Index As Long) As Boolean
    GetPlayerTurn = Player(Index).Char(Player(Index).CharNum).Turn
End Function

Sub SetPlayerTurn(ByVal Index As Long, ByVal Turn As Boolean)
    Player(Index).Char(Player(Index).CharNum).Turn = Turn
End Sub

Function GetPlayerOldX(ByVal Index As Long) As Integer
    GetPlayerOldX = Player(Index).Char(Player(Index).CharNum).OldX
End Function

Sub SetPlayerOldX(ByVal Index As Long, ByVal X As Integer)
    Player(Index).Char(Player(Index).CharNum).OldX = X
End Sub

Function GetPlayerOldY(ByVal Index As Long) As Integer
    GetPlayerOldY = Player(Index).Char(Player(Index).CharNum).OldY
End Function

Sub SetPlayerOldY(ByVal Index As Long, ByVal Y As Integer)
    Player(Index).Char(Player(Index).CharNum).OldY = Y
End Sub

Function GetPlayerTurnBased(ByVal Index As Long) As Boolean
    GetPlayerTurnBased = Player(Index).Char(Player(Index).CharNum).HasTurnBased
End Function

Sub SetPlayerTurnBased(ByVal Index As Long, ByVal TurnBased As Boolean)
    Player(Index).Char(Player(Index).CharNum).HasTurnBased = TurnBased
End Sub

Function GetPlayerRecoverTime(ByVal Index As Long) As Long
    GetPlayerRecoverTime = Player(Index).Char(Player(Index).CharNum).RecoverTime
End Function

Sub SetPlayerRecoverTime(ByVal Index As Long, ByVal RecoverTime As Long)
    Player(Index).Char(Player(Index).CharNum).RecoverTime = RecoverTime
End Sub

Function GetPlayerEquipSlotNum(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotNum = Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).num
End Function

Function GetPlayerEquipSlotValue(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotValue = Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).Value
End Function

Function GetPlayerEquipSlotAmmo(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotAmmo = Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).Ammo
End Function

Sub SetPlayerEquipSlotNum(ByVal Index As Long, ByVal EquipSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).num = ItemNum
End Sub

Sub SetPlayerEquipSlotValue(ByVal Index As Long, ByVal EquipSlot As Long, ByVal Value As Long)
    Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).Value = Value
End Sub

Sub SetPlayerEquipSlotAmmo(ByVal Index As Long, ByVal EquipSlot As Long, ByVal Ammo As Long)
    Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).Ammo = Ammo
End Sub

Function GetPartyMember(ByVal PartyNum As Long, ByVal Member As Long) As Long
    GetPartyMember = Party(PartyNum).Member(Member)
End Function

Sub SetPartyMember(ByVal PartyNum As Long, ByVal Member As Long)
    Dim i As Byte

    If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Then Exit Sub
    
    For i = 1 To MAX_PARTY_MEMBERS
        If Party(PartyNum).Member(i) = 0 Then
            Party(PartyNum).Member(i) = Member
            Call SetPlayerPartyNum(Member, PartyNum)
            Exit For
        End If
    Next i
End Sub

Sub RemovePartyMember(ByVal PartyNum As Long, ByVal Member As Long)
    Dim i As Byte
    
    If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Then Exit Sub
    
    For i = 1 To MAX_PARTY_MEMBERS
        If Party(PartyNum).Member(i) = Member Then
            Party(PartyNum).Member(i) = 0
            Call SetPlayerPartyNum(Member, 0)
            Exit For
        End If
    Next i
End Sub

Function GetPartyLeader(ByVal PartyNum As Long) As Long
    GetPartyLeader = Party(PartyNum).Leader
End Function

Sub SetPartyLeader(ByVal PartyNum As Long, ByVal Leader As Long)
    Party(PartyNum).Leader = Leader
End Sub

Function GetPartyMembers(ByVal PartyNum As Long) As Long
    Dim i As Long
    
    GetPartyMembers = 0
    
    If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Then Exit Function
    
    For i = 1 To MAX_PARTY_MEMBERS
        If Party(PartyNum).Member(i) > 0 Then
            GetPartyMembers = GetPartyMembers + 1
        End If
    Next i
End Function

Function GetPlayerPartyNum(ByVal Index As Long) As Long
    GetPlayerPartyNum = Player(Index).Char(Player(Index).CharNum).PartyNum
End Function

Sub SetPlayerPartyNum(ByVal Index As Long, ByVal PartyNum As Long)
    Player(Index).Char(Player(Index).CharNum).PartyNum = PartyNum
End Sub

Function GetPartyShareCount(ByVal Index As Long) As Byte
    Dim i As Byte
    Dim PartyNum As Long
    
    PartyNum = GetPlayerPartyNum(Index)
    
    GetPartyShareCount = 0
    
    If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Then Exit Function
    
    For i = 1 To MAX_PARTY_MEMBERS
        If Party(PartyNum).ShareExp(i) = True Then
            GetPartyShareCount = GetPartyShareCount + 1
        End If
    Next i
End Function

Function GetPlayerPartyShare(ByVal Index As Long) As Boolean
    Dim i As Byte
    Dim PartyNum As Long
    
    PartyNum = GetPlayerPartyNum(Index)
    
    If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Then Exit Function
    
    For i = 1 To MAX_PARTY_MEMBERS
        If Party(PartyNum).Member(i) = Index Then
            GetPlayerPartyShare = Party(PartyNum).ShareExp(i)
        End If
    Next i
End Function

Sub SetPlayerPartyShare(ByVal Index As Long, ByVal Share As Boolean)
    Dim i As Byte
    Dim PartyNum As Long
    
    PartyNum = GetPlayerPartyNum(Index)
    
    If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Then Exit Sub
    
    For i = 1 To MAX_PARTY_MEMBERS
        If Party(PartyNum).Member(i) = Index Then
            Party(PartyNum).ShareExp(i) = Share
            Exit For
        End If
    Next i
End Sub

Function GetPlayerAttackSpeed(ByVal Index As Long) As Long
    Dim i As Integer
    
    GetPlayerAttackSpeed = 1000
    
    For i = 1 To 7
        If GetPlayerEquipSlotNum(Index, i) > 0 Then
            GetPlayerAttackSpeed = GetPlayerAttackSpeed - (1000 - Item(GetPlayerEquipSlotNum(Index, i)).AttackSpeed)
        End If
    Next i
    
    GetPlayerAttackSpeed = GetPlayerAttackSpeed - (GetPlayerSPEED(Index) * 3)
End Function
