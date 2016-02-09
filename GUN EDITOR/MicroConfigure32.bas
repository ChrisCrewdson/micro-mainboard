[setup.main.Window]
    ' Micro MT Configuration Utility
    ' Jim Robertson
    ' 3 Jun 2015
    ' www.lasertagparts.com
    ' Source File MicroConfigure32.bas


    'BEGIN CODE FOR #main

    Global  MAXCOMPORTS             ' How many com ports can we hold in our drop down list?
    MAXCOMPORTS = 30
    Global  HIGHESTCOMPORT          ' Highest numerical com port to check for
    HIGHESTCOMPORT = 30

    PaceDelay = 100      'sets the delay time between transmitted bytes

    'SET INITIAL PARAMETER VALUES
    ValTeamID      = 0      'Team ID = Red
    ValPlayerID    = 49     'Player ID should be 40 ("Micro")
    ValDamage      = 5      '(Table Pointer)
    ValRounds      = 40     '40 Rounds
    ValClips       = 20     '20 Clips
    ValFireSel     = 2      'Full Auto
    ValBurst       = 3      '3-round Burst (If Fire Select = Burst)
    ValCyclic      = 6      'Cyclic = 550 RPM (Table)
    ValRelDelay    = 3      'Reload Delay = 3 Seconds
    ValHealth      = 36     'Health = 100 (Table Pointer, 36 = 100%)
    ValRespDelay   = 6
    ValArmor       = 0
    ValHitDelay    = 4
    ValMaxRespawn  = 0
    ValFriendFire  = 4      '4=On, 0=Off
    ValVolume      = 0
    ValHitledEn    = 8

    ValBright      = 125

    'SET INITIAL INDEX POSITIONS FOR COMBO BOXES
    IndexTeamID     = 1
    IndexDamage     = 6
    IndexFireSel    = 3
    IndexBurst      = 1
    IndexCyclic     = 7
    IndexHitDelay   = 5



    nomainwin
    WindowWidth = 670
    WindowHeight = 310
    UpperLeftX=int((DisplayWidth-WindowWidth)/2)
    UpperLeftY=int((DisplayHeight-WindowHeight)/2)

    'GLOBAL VARIABLES
    Global  OpenedCOMPort$          ' Holds the name of the com port we currently have opened, if any
    OpenedCOMPort$ = ""

    'DIM ALL ARRAYS
    DIM COMPort$(MAXCOMPORTS)       ' Holds strings for COM port numbers (COM1, COM2, etc.)
    DIM NewCOMPort$(MAXCOMPORTS)    ' Holds strings for COM port numbers (COM1, COM2, etc.) inside GetCOMPort routine

    DIM damage$(20)
    DIM cyclic$(20)
    DIM hitdelay$(30)

    firemode$(1) = "Semi-Auto"
    firemode$(2) = "Burst"
    firemode$(3) = "Full-Auto"

    teamID$(1) = "Red"
    teamID$(2) = "Blue"
    teamID$(3) = "Yellow"
    teamID$(4) = "Green"

    damage$(1) = "1"
    damage$(2) = "2"
    damage$(3) = "4"
    damage$(4) = "5"
    damage$(5) = "7"
    damage$(6) = "10"
    damage$(7) = "15"
    damage$(8) = "17"
    damage$(9) = "20"
    damage$(10) = "25"
    damage$(11) = "30"
    damage$(12) = "35"
    damage$(13) = "40"
    damage$(14) = "50"
    damage$(15) = "75"
    damage$(16) = "100"

    cyclic$(1) = "250 RPM"
    cyclic$(2) = "300 RPM"
    cyclic$(3) = "350 RPM"
    cyclic$(4) = "400 RPM"
    cyclic$(5) = "450 RPM"
    cyclic$(6) = "500 RPM"
    cyclic$(7) = "550 RPM"
    cyclic$(8) = "600 RPM"
    cyclic$(9) = "650 RPM"
    cyclic$(10) = "700 RPM"
    cyclic$(11) = "750 RPM"
    cyclic$(12) = "800 RPM"

    burst$(1) = "3"
    burst$(2) = "4"
    burst$(3) = "5"
    burst$(4) = "6"

    hitdelay$(1) = ".00 Sec"
    hitdelay$(2) = ".25 Sec"
    hitdelay$(3) = ".50 Sec"
    hitdelay$(4) = ".75 Sec"
    hitdelay$(5) = "01 Sec"
    hitdelay$(6) = "02 Sec"
    hitdelay$(7) = "03 Sec"
    hitdelay$(8) = "04 Sec"
    hitdelay$(9) = "05 Sec"
    hitdelay$(10) = "06 Sec"
    hitdelay$(11) = "07 Sec"
    hitdelay$(12) = "08 Sec"
    hitdelay$(13) = "09 Sec"
    hitdelay$(14) = "10 Sec"
    hitdelay$(15) = "11 Sec"
    hitdelay$(16) = "12 Sec"
    hitdelay$(17) = "13 Sec"
    hitdelay$(18) = "14 Sec"
    hitdelay$(19) = "15 Sec"
    hitdelay$(20) = "16 Sec"
    hitdelay$(21) = "17 Sec"
    hitdelay$(22) = "18 Sec"
    hitdelay$(23) = "19 Sec"
    hitdelay$(24) = "20 Sec"



    'BEGIN GUI OBLECTS CODE

    ComboboxColor$ = "white"
    combobox #main.firemodecombo, firemode$(), [firemodeDoubleClick],   20,  27, 100, 100
    combobox #main.burstcombo, burst$(), [burstDoubleClick],   260,  27, 60, 100
    combobox #main.COMPort, COMPort$(), [portselectDoubleClick],  460, 192,  95, 100
    combobox #main.teamIDcombo, teamID$(), [teamIDDoubleClick],  460,  27, 100, 100
    combobox #main.damagecombo, damage$(), [damageDoubleClick], 20, 227, 60, 100
    combobox #main.cycliccombo, cyclic$(), [cyclicDoubleClick],   20,  67, 100, 100
    combobox #main.hitdelaycombo, hitdelay$(), [hitdelayDoubleClick],  260, 107,  75, 100

    TextboxColor$ = "white"
    textbox #main.rounds,  20, 107,  60,  25
    textbox #main.clips,  20, 147,  60,  25
    textbox #main.reload,  20, 187,  60,  25
    textbox #main.player, 460,  67,  60,  25
    textbox #main.health, 260,  147,  60,  25
    textbox #main.volume, 260, 227, 60, 25
    textbox #main.bright, 460, 147, 60, 25

    button #main.sendbutton,"Send To Gun",[sendbuttonClick], UL, 460, 227, 105,  25

    statictext #main.statictext01, "Rnds/Clip",  95, 112,  96,  20
    statictext #main.statictext02, "Clips",  95, 152,  30,  20
    statictext #main.statictext16, "Player ID", 535,  72,  71,  20
    statictext #main.statictext04, "Reload Delay",  95, 192,  84,  20
    statictext #main.statictext05, "Hit Delay", 350, 112,  55,  20
    statictext #main.statictext06, "Firing Mode", 135,  32,  71,  20
    statictext #main.statictext07, "Burst", 335,  32,  30,  20
    statictext #main.statictext09, "Health", 335,  152,  39,  20
    statictext #main.statictext10, "Damage",  95, 232,  53,  20
    statictext #main.statictext11, "Team ID", 575,  32,  52,  20
    statictext #main.statictext13, "Firing Rate", 135,  72, 106,  20
    statictext #main.statictext14, "COM Port", 575, 197,  57,  20
    statictext #main.statictext17, "Brightness", 535, 152,  100,  20
    statictext #main.statictext18, "Volume", 335, 232,  57,  20

    checkbox #main.FriendFire, "Friendly Fire", [FriendFireSet], [FriendFireReset], 460, 107,  99,  25
    checkbox #main.Hitled, "Hit LED Enable", [HitledSet], [HitledReset], 260, 187,  110,  25

    'END GUI OBJECTS CODE

    open "MilesTag - uMT Editor V32" for window as #main
    print #main, "font ms_sans_serif 10"
    print #main, "trapclose [quit.main]"

    ' Populate the drop down list of available COM ports
    gosub [GetCOMPorts]


    print #main.firemodecombo, "selectindex "; IndexFireSel
    print #main.teamIDcombo, "selectindex "; IndexTeamID
    'print #main.carriercombo, "selectindex "; IndexCarrier
    print #main.damagecombo, "selectindex "; IndexDamage
    print #main.cycliccombo, "selectindex "; IndexCyclic
    print #main.hitdelaycombo, "selectindex "; IndexHitDelay
    print #main.burstcombo, "selectindex "; IndexBurst

    print #main.player, ValPlayerID
    print #main.rounds, ValRounds
    print #main.clips, ValClips
    print #main.reload, ValRelDelay
    print #main.health, ValHealth
    print #main.volume, ValVolume
    print #main.bright, ValBright

    print #main.FriendFire, "set"
    print #main.Hitled, "set"

    IF ValFireSel <> 1 THEN
        print #main.burstcombo, "disable"
    END IF

'#####################################################################################
'#####################################################################################
'#####################################################################################
[main.inputLoop]   'wait here for input event


    wait

'#####################################################################################
[firemodeDoubleClick]
    print #main.firemodecombo, "selectionindex? IndexFireSel"
    ValFireSel = IndexFireSel - 1
    IF ValFireSel = 1 THEN
        print #main.burstcombo, "enable"
    ELSE
        print #main.burstcombo, "disable"
    END IF

    wait

'#####################################################################################
[burstDoubleClick]
    print #main.burstcombo, "selectionindex? IndexBurst"
    ValBurst = IndexBurst + 2

    wait

'#####################################################################################
'When the send button is pressed, values are read into the variables from the text
'boxes, then all variables are transmitted via the COM port at 9600 baud.  Delays are
'utilized to provide pacing between the tranmitted bytes.  This prevents receive
'problems on the PIC16F684.

[sendbuttonClick]
    print #main.volume, "!contents? ValVolume"
    IF ValVolume > 4 THEN
        NOTICE "Volume must be between 0 and 4"
        GOTO [main.inputLoop]
    END IF

    print #main.player, "!contents? ValPlayerID"
    IF ValPlayerID > 50 THEN
        NOTICE "Player ID must be between 0 and 50"
        GOTO [main.inputLoop]
    END IF

    print #main.rounds, "!contents? ValRounds"
    IF ValRounds < 1 OR ValRounds > 255 THEN
        NOTICE "Rounds must be between 1 and 255"
        GOTO [main.inputLoop]
    END IF

    print #main.clips, "!contents? ValClips"
    IF ValClips < 1 OR ValValClips > 200 THEN
        NOTICE "Clips must be between 1 and 200"
        GOTO [main.inputLoop]
    END IF

    print #main.reload, "!contents? ValRelDelay"
    IF ValRelDelay < 1 OR ValRelDelay > 30 THEN
        NOTICE "Reload Delay must be between 1 and 30"
        GOTO [main.inputLoop]
    END IF

    print #main.health, "!contents? ValHealth"
    IF ValHealth < 1 OR ValHealth > 64 THEN
        NOTICE "Health must be between 1 and 64"
        GOTO [main.inputLoop]
    END IF

    print #main.bright, "!contents? ValBright"
    IF ValBright < 0 OR ValBright > 255 THEN
        NOTICE "Brightness must be between 0 and 255"
        GOTO [main.inputLoop]
    END IF


    print #commHandle, chr$(197)            'Start Byte 0xC5
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)              'Zone
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValTeamID)       'TeamID
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValPlayerID)    'PlayerID
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValVolume)      'Volume Max
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 4
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 5
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 6
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 7
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 8
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 9
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 10
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 11
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValDamage)     'Damage Pointer 12
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValRounds)   'Rounds 13
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValClips)     'Clips 14
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValFireSel)     'FireSel 15
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValBurst)     'Burst 16
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValCyclic)     'Cyclic Pointer 17
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValRelDelay)     'RelDelay 18
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 19
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(10)    'Range 20
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValHitledEn)    'Hit LED Enable    'CLass Bits (8=HitLED ON, 0=HitLED OFF) 21
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValHealth)    'Health Pointer 22
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 23
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 24
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 25
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValFriendFire)     'Friendly Fire (Flags Variable, Bit 2) 26
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(0)     'Dummy Data 27
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValHitDelay)     'HitDelay 28
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(ValBright)     'LED Brightness 29
    call waitMilliseconds PaceDelay

    print #commHandle, chr$(172)     'End Byte
    call waitMilliseconds PaceDelay

    wait

'#####################################################################################
[portselectDoubleClick]       'Close the old one (if any) and open the new one
    gosub [CloseSerialPort]
    gosub [OpenSerialPort]

    wait

'#####################################################################################
[teamIDDoubleClick]
    print #main.teamIDcombo, "selectionindex? IndexTeamID"
    ValTeamID = IndexTeamID - 1

    wait

'#####################################################################################
[FriendFireSet]
        ValFriendFire  = 4
    wait

[FriendFireReset]
        ValFriendFire  = 0
    wait
'#####################################################################################
[HitledSet]
        ValHitledEn  = 8
    wait

[HitledReset]
        ValHitledEn  = 0
    wait
'#####################################################################################
[damageDoubleClick]
    print #main.damagecombo, "selectionindex? IndexDamage"
    ValDamage = IndexDamage - 1

    wait

'#####################################################################################
[cyclicDoubleClick]
    print #main.cycliccombo, "selectionindex? IndexCyclic"
    ValCyclic = IndexCyclic - 1

    wait

'#####################################################################################
[hitdelayDoubleClick]
    print #main.hitdelaycombo, "selectionindex? IndexHitDelay"
    ValHitDelay = IndexHitDelay - 1
    'print #main.hdtest, ValHitDelay
    wait

'#####################################################################################
[quit.main] 'End the program

    close #commHandle

    close #main
    end





'#####################################################################################
'#####################################################################################

'#############################
'# COM PORT UTILITY ROUTINES #
'#############################
'
' GetCOMPorts
'
' This subroutine will populate the drop-down list called #main.CombComPort
' with all of the currently available COM ports. You can change #main.COMPort
' to whatever you use in your program.
'
[GetCOMPorts]
    ' First clear out our 'new' COM Ports array
    for Port = 0 to MAXCOMPORTS
        NewCOMPort$(Port) = ""
    next Port

    ' We start with index 1, since that's where the combo box starts
    CurrentComboIndex = 1
    ' We're going to loop from 1 to 254 and see if any of those COM ports are available
    for TestPort = 1 to HIGHESTCOMPORT
        ' If we hit upon our currently opened COM port, we add that to our list too to keep things simple
        if IsCOMPortAvailable(TestPort) = 1 OR "COM" + str$(TestPort) = OpenedCOMPort$ then
            ' Add this port into 'new' COM port array
            NewCOMPort$(CurrentComboIndex) = "COM" + str$(TestPort)
            CurrentComboIndex = CurrentComboIndex + 1
        end if
    next TestPort

    ' Now compare to see if there are any differences between the two lists
    ' of COM ports (old vs new)
    bAreTheSame = TRUE
    for Port = 1 to MAXCOMPORTS
        ' Now, if the COM port strings at this index are not the same, then something's changed
        ' and we need to remember that fact and exit the loop
        if NewCOMPort$(Port) <> COMPort$(Port) then
            bAreTheSame = FALSE
            exit for
        end if
    next Port

    ' If they are not the same, then copy over new to old and update combobox
    if (bAreTheSame = FALSE) then
        for Port = 1 to MAXCOMPORTS
            COMPort$(Port) = NewCOMPort$(Port)
        next Port

        ' Now reload the combbox to get the new array values
        print #main.COMPort, "reload"
    end if

    ' If the currently selected COM port is empty, and we have some COM port in our list, then
    ' set our currently selected COM port to the first one.
    print #main.COMPort, "contents? CurComPort$"
    if CurComPort$ = "" and CurrentComboIndex > 1 then
        print #main.COMPort, "selectindex 1"
        gosub [CloseSerialPort]
        gosub [OpenSerialPort]
    end if

    return

' Helper function for [GetCOMPorts]
' Returns 0 if the Port is not available on this system (either already opened or doesn't
' exist) and 1 if the port is available.
function IsCOMPortAvailable(Port)
    struct lpSecurityAttributes, _
        nLength as long, _
        lpSecurityDescriptor as long, _
        bInheritHandle as long

    lpFileName$ = "\\.\COM" + str$(Port)

    dwDesiredAccess = 0
    dwShareMode = 1 + 2
    dwCreationDisposition = 3
    dwFlagsAndAttributes = 128
    hTemplateFile = 0

    calldll #kernel32, "CreateFileA", _
        lpFileName$ as ptr, _
        dwDesiredAccess as long, _
        dwShareMode as long, _
        lpSecurityAttributes as struct, _
        dwCreationDisposition as long, _
        dwFlagsAndAttributes as long, _
        hTemplateFile as long, _
        retcode as long
    hObject = retcode

    ' Test the return code to see if the port is available and free
    if retcode = -1 then
        IsCOMPortAvailable = 0
    else
        ' If we were able to open it, then close it and return TRUE
        calldll #kernel32, "CloseHandle", _
            hObject as long, _
            retcode as long

        IsCOMPortAvailable = 1
    end if

    ' The other reason we could have an available com

end function


' Handle opening the serial port
[OpenSerialPort]

    oncomerror [ComErrorHandler]

    print #main.COMPort, "contents? ComPort$"

    if ComPort$ <> "" and OpenedCOMPort$ <> ComPort$ then

        print #main.COMPort, "contents? ComPort$"
        OpenedCOMPort$ = ComPort$

        open ComPort$;":9600,8,N,1,RS,DS0,CS0" for random as #commHandle

    end if

    return

' Handle closing the serial port
[CloseSerialPort]
    if OpenedCOMPort$ <> "" then
        close #commHandle
        OpenedCOMPort$ = ""
    end if

    return

' We get here if there is a com error at some point.
[ComErrorHandler]
    ' Display the com error box
    notice "Com Error"; chr$(13); "Some type of com port error has occured. Your computer may not have a valid com port at COM"; str$(ComPortNumber); " or there might be another error. "; chr$(13); "Error: "; ComError$; " Port: "; str$(ComPortNumber); " ErrorNo: "; str$(ComErrorNumber)

    gosub [CloseSerialPort]

    wait

' Delay Subroutine
sub waitMilliseconds ms
    timer ms, [stopWaiting]
    wait
    [stopWaiting]
    timer 0
end sub
