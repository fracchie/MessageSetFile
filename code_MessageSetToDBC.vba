Sub MessageSet_to_DBC()

    '======= The debug stuff
    Debug.Print ("To detect which line (hence signal) is blocking the code, type the following")
    Debug.Print ("?SignalNameRangeD.Cells(i,1).value")
    Debug.Print ("")
    Debug.Print ("To erase the whole Debug log (Immediate Window), clink in it, select all - ctrl + a - end press enter, or type the following")
    Debug.Print ("")
    '============

    Worksheets("MessageSet").Activate

    Dim DebugMode As Integer: DebugMode = 1

    '================ external file declaration & creation ==================
    'Be sure you set a reference to the VB script run-time library. Follow https://stackoverflow.com/questions/3233203/how-do-i-use-filesystemobject-in-vba
    Dim filePath As String
    Dim fileName As String
    Dim objShell As Object, objFolder As Object, objFolderItem As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(&H0&, "Choose file's path", &H1&)
    Set objFolderItem = objFolder.Items.Item
    filePath = objFolderItem.Path
    Dim TempByteSent As String
    Dim MyFSO As New FileSystemObject
    If MyFSO.FolderExists(filePath) Then
        'MsgBox "The Folder already exists"
    Else
        MyFSO.CreateFolder (filePath) '<- Here the
    End If

    'Either as a TextStream or as an Object
    'Dim FileOut As TextStream
    'TODO get filename from dashboard
    fileName = "CANdbcTest.dbc"
    'Set FileOut = MyFSO.CreateTextFile(filePath + "\" + fileName, True, True)
    'write
    'FileOut.WriteLine ("prova")
    Dim stream As Object
    Set stream = MyFSO.CreateTextFile(filePath + "\" + fileName, True, False)
    'stream.WriteLine "helloWorld"

    'Configure message set table
    Call Expand_All 'if some group is closed, macro will not be able to get their info
    Dim HeadersRangeD As Range: Set HeadersRangeD = Range("A1", Range("A1").End(xlToRight).Address)
    HeadersRangeD.Select
    'would like to format the whole thing as a tab, and maybe formatting the headers as text
    'Find the needed columns in the header list. By default is NOT CASE SENSITIVE
    Dim SignalNameRangeD As Range: Set SignalNameRangeD = Range(HeadersRangeD.Find("Signal Name", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Name", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignalDescriptionRangeD As Range: Set SignalDescriptionRangeD = Range(HeadersRangeD.Find("Description", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Description", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FrameNameRangeD As Range: Set FrameNameRangeD = Range(HeadersRangeD.Find("Frame Name", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Name", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FrameIDRangeD As Range: Set FrameIDRangeD = Range(HeadersRangeD.Find("Frame ID (Hexa)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame ID (Hexa)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FrameSenderRangeD As Range: Set FrameSenderRangeD = Range(HeadersRangeD.Find("Sender", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Sender", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FrameSizeRangeD As Range: Set FrameSizeRangeD = Range(HeadersRangeD.Find("Frame Size (Bytes)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Size (Bytes)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FramePeriodRangeD As Range: Set FramePeriodRangeD = Range(HeadersRangeD.Find("Period (ms)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Period (ms)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim BytePositionRangeD As Range: Set BytePositionRangeD = Range(HeadersRangeD.Find("Byte Position", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Byte Position", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim BitPositionRangeD As Range: Set BitPositionRangeD = Range(HeadersRangeD.Find("Bit Position", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Bit Position", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ValueTypeRangeD As Range: Set ValueTypeRangeD = Range(HeadersRangeD.Find("Value Type", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value Type", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ValueEndianRangeD As Range: Set ValueEndianRangeD = Range(HeadersRangeD.Find("Endian", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Endian", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignalSizeRangeD As Range: Set SignalSizeRangeD = Range(HeadersRangeD.Find("Signal Size (Bit)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Size (Bit)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignalUnitRangeD As Range: Set SignalUnitRangeD = Range(HeadersRangeD.Find("Unit", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unit", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignalResolutionRangeD As Range: Set SignalResolutionRangeD = Range(HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignalOffsetRangeD As Range: Set SignalOffsetRangeD = Range(HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignalMinRangeD As Range: Set SignalMinRangeD = Range(HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignalMaxRangeD As Range: Set SignalMaxRangeD = Range(HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignalValueTableRangeD As Range: Set SignalValueTableRangeD = Range(HeadersRangeD.Find("Value Table", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value Table", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SenderNameRangeD As Range: Set SenderNameRangeD = Range(HeadersRangeD.Find("Sender", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Sender", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim StartBitRangeD As Range: Set StartBitRangeD = Range(HeadersRangeD.Find("Start Bit", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Start Bit", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))

    '======== .dbc headers =========
    stream.WriteLine ("VERSION " + Chr(34) + Chr(34))
    stream.WriteBlankLines (2)

    'NS, BS ======
    stream.WriteLine ("NS_ :")
    stream.WriteLine Chr(9) + "NS_DESC_"
    stream.WriteLine Chr(9) + "CM_"
    stream.WriteLine Chr(9) + "BA_DEF_"
    stream.WriteLine Chr(9) + "BA_"
    stream.WriteLine Chr(9) + "VAL_"
    stream.WriteLine Chr(9) + "CAT_DEF_"
    stream.WriteLine Chr(9) + "CAT_"
    stream.WriteLine Chr(9) + "Filter"
    stream.WriteLine Chr(9) + "BA_DEF_DEF_"
    stream.WriteLine Chr(9) + "EV_DATA_"
    stream.WriteLine Chr(9) + "ENVVAR_DATA_"
    stream.WriteLine Chr(9) + "SGTYPE_"
    stream.WriteLine Chr(9) + "SGTYPE_VAL_"
    stream.WriteLine Chr(9) + "BA_DEF_SGTYPE_"
    stream.WriteLine Chr(9) + "BA_SGTYPE_"
    stream.WriteLine Chr(9) + "SIG_TYPE_REF_"
    stream.WriteLine Chr(9) + "SIG_GROUP_"
    stream.WriteLine Chr(9) + "SIG_VALTYPE_"
    stream.WriteLine Chr(9) + "SIGTYPE_VALTYPE_"
    stream.WriteLine Chr(9) + "BO_TX_BU_"
    stream.WriteLine Chr(9) + "BA_DEF_REL_"
    stream.WriteLine Chr(9) + "BA_REL_"
    stream.WriteLine Chr(9) + "BA_DEF_DEF_REL_"
    stream.WriteLine Chr(9) + "BU_SG_REL_"
    stream.WriteLine Chr(9) + "BU_EV_REL_"
    stream.WriteLine Chr(9) + "BU_BO_REL_"
    stream.WriteLine Chr(9) + "SG_MUL_VAL_"
    stream.WriteBlankLines (1)
    stream.WriteLine ("BS_ :")
    stream.WriteBlankLines (1)

    'BU_: list of ECUs
    'normalise names so that CANdb accepts them i.e. no number, no special characters, only "_"
    'stream.WriteLine ("BU_ : ECU_A ECU_B")
    Dim ECU_list As New Collection
    For i = 2 To SenderNameRangeD.Count
        If SenderNameRangeD.Cells(i, 1).value <> SenderNameRangeD.Cells(i - 1, 1).value Then
            If Not CollectionContainsString(ECU_list, SenderNameRangeD.Cells(i, 1).value) Then
                ECU_list.Add (SenderNameRangeD.Cells(i, 1).value)
            End If
        End If
    Next i

    stream.Write ("BU_:")
    For i = 1 To ECU_list.Count
        stream.Write (" " + ECU_list(i))
    Next i
    stream.WriteLine

    'To store the VAL to be written at the end. In the BO/SG loop, write the line for the signal, and store the VAL_ to be written later
    Dim Val_List As New Collection: Set Val_List = Nothing
    'Store frame period and info for the BA section at the end
    Dim BA_List As New Collection: Set BA_List = Nothing
    'Store comment for later
    Dim Comment_List As New Collection: Set Comment_List = Nothing

    'start with signals
    For i = 2 To SignalNameRangeD.Count
        'compare frame ID to divide frames
        'Debug.Print (SignalNameRangeD.Cells(i, 1).value)

        If FrameIDRangeD.Cells(i, 1) <> FrameIDRangeD.Cells(i - 1, 1) Then 'new frame
            'write new frame container
            If (DebugMode) Then
                Debug.Print (FrameIDRangeD.Cells(i, 1).value)
            End If

            '======= BO_: ==========
            stream.WriteBlankLines (1)
            stream.WriteLine ("BO_ " + Str(CLng("&H" & FrameIDRangeD.Cells(i, 1).value)) + " " + FrameNameRangeD.Cells(i, 1).value + ": " + Replace(Str(FrameSizeRangeD.Cells(i, 1).value), Space(1), Space(0)) + " " + FrameSenderRangeD.Cells(i, 1).value)

            'store frame info for BA section later
            If (FramePeriodRangeD.Cells(i, 1).value <> "-") Then
                BA_List.Add ("BA_ " + Chr(34) + "CycleTime" + Chr(34) + " BO_ " + Str(CLng("&H" & FrameIDRangeD.Cells(i, 1).value)) + Str(FramePeriodRangeD.Cells(i, 1).value) + ";")
            Else
                'TODO when the signal is event based only
            End If
        End If

        'each line is a new signal
        Dim text As String

        If (DebugMode) Then
            Debug.Print (SignalNameRangeD.Cells(i, 1).value)
        End If

        'Comment
        If (SignalDescriptionRangeD.Cells(i, 1).value <> "") Then
            Comment_List.Add ("CM_ SG_ " + Str(CLng("&H" & FrameIDRangeD.Cells(i, 1).value)) + " " + SignalNameRangeD.Cells(i, 1).value + " " + Chr(34) + SignalDescriptionRangeD.Cells(i, 1).value + Chr(34) + ";")
        End If

        'name and position
        text = " SG_ " + SignalNameRangeD.Cells(i, 1).value + " : "

        'position
        Dim startPositionStr As String
        'Bit positions are counted from byte 0 upwards by their significance, regardless of the endianness. The first message byte has bits 0…7 with bit 7 being the most significant bit of the byte. The second byte has bits 8…15 with bit 15 being the MSB, and so on: 7 6 5 4 3 2 1 0 15 14 13 12 11 10 9 8 ...
        'For big endian values, signal start bit positions are given for the most significant bit. For little endian values, the start position is that of the least significant bit.
        If ValueEndianRangeD.Cells(i, 1).value = "Little Endian" Then
            If Not IsEmpty(StartBitRangeD.Cells(i, 1).value) Then
                startPositionStr = StartBitRangeD.Cells(i, 1).value
            Else
                'startbit cell is not filled, maybe the startByte and StartBit?
                MsgBox ("Error: Startbit cell not filled")
                Exit Sub
            End If
        ElseIf ValueEndianRangeD.Cells(i, 1).value = "Big Endian" Then 'Big endian
            If Not IsEmpty(StartBitRangeD.Cells(i, 1).value) Then
                startPositionStr = StartBitRangeD.Cells(i, 1).value

                'Manage the trick that candb import does (what written in dbc file, when it has the big endian label @0, changes taking the value of th msb on candb
                Dim temp As Integer: temp = SignalSizeRangeD.Cells(i, 1).value
                Dim bit As Integer: bit = StartBitRangeD.Cells(i, 1).value Mod 8
                Dim ByteN As Integer: ByteN = StartBitRangeD.Cells(i, 1).value \ 8
                Do While (temp > (8 - bit)) 'it goes up in the previous byte
                    ByteN = ByteN - 1
                    temp = temp - (8 - bit)
                    bit = 0
                Loop
                startPositionStr = ByteN * 8 + bit + temp - 1

            Else
                'startbit cell is not filled, maybe the startByte and StartBit?
                MsgBox ("Error: Startbit cell not filled")
                ' OLD, not working anymore startPositionStr = Replace(BytePositionRangeD.Cells(i, 1).value * 8 + BitPositionRangeD.Cells(i, 1).value, Space(1), Space(0))

            End If

        Else
            MsgBox ("Error: endian unknown?")
            Exit Sub
        End If
        text = text + startPositionStr + "|" + Replace(Str(SignalSizeRangeD.Cells(i, 1).value), Space(1), Space(0))

        'Endian
        If (ValueEndianRangeD.Cells(i, 1).value = "Little Endian") Then
            text = text + "@1"
        Else 'Big Endian"
            text = text + "@0"
        End If

        'Signed Unsigned
        Select Case (ValueTypeRangeD.Cells(i, 1).value)
            Case "Signed"
                text = text + "-"
            Case Else 'Unsigned, List, Hexa..
                text = text + "+"
        End Select

        'Resolution and offset
        Select Case (ValueTypeRangeD.Cells(i, 1).value)
            Case "Signed", "Unsigned"
                text = text + " (" + Replace(Str(SignalResolutionRangeD.Cells(i, 1).value), Space(1), Space(0)) + "," + Replace(Str(SignalOffsetRangeD.Cells(i, 1).value), Space(1), Space(0)) + ")"

            Case Else 'List, Hexa"
                text = text + " (1,0)"
        End Select

        'Min max values
        Select Case (ValueTypeRangeD.Cells(i, 1).value)
            Case "Signed", "Unsigned"
                text = text + " [" + Replace(Str(SignalMinRangeD.Cells(i, 1).value), Space(1), Space(0)) + "|" + Replace(Str(SignalMaxRangeD.Cells(i, 1).value), Space(1), Space(0)) + "] " + Chr(34) + SignalUnitRangeD.Cells(i, 1).value + Chr(34)
            Case "List"
                text = text + "[0|0] " + Chr(34) + Chr(34)
                'store val_list array to be written later
                Dim temp_list() As String
                temp_list = Split(SignalValueTableRangeD.Cells(i, 1).value, vbLf)
                Dim val_list_n As String
                val_list_n = "VAL_ " + Replace(Str(CLng("&H" & FrameIDRangeD.Cells(i, 1).value)), Space(1), Space(0)) + " " + SignalNameRangeD.Cells(i, 1).value
                For l = 0 To UBound(temp_list)
                    val_list_n = val_list_n + " " + Left(temp_list(l), InStr(temp_list(l), ":") - 1) + " " + Chr(34) + Right(temp_list(l), Len(temp_list(l)) - InStr(temp_list(l), ":") - 1) + Chr(34)
                Next l
                val_list_n = val_list_n + ";"

                Val_List.Add val_list_n

            Case "Hexa"
                text = text + "[0|0] " + Chr(34) + Chr(34) + " "

        End Select

        'receiver ECUs for each signal
        For j = SenderNameRangeD.column To HeadersRangeD.Count
            If (Cells(i, j).value = "R") Then
                text = text + Cells(1, j).value + ","
            End If
        Next j

        text = text + " Vector__XXX"

        'final write text
        stream.WriteLine text


    Next i

    stream.WriteBlankLines (2)
    'CM_:
    For i = 1 To Comment_List.Count
        stream.WriteLine Comment_List(i)
    Next i

    'BA_DEF
    stream.WriteLine ("BA_DEF_ BO_ " + Chr(34) + "CycleTime" + Chr(34) + " INT 0 10000;")
    stream.WriteLine ("BA_DEF_ BO_ " + Chr(34) + "FrameType" + Chr(34) + " STRING;")
    stream.WriteLine ("BA_DEF_DEF_ " + Chr(34) + "CycleTime" + Chr(34) + " 100;")
    stream.WriteLine ("BA_DEF_DEF_ " + Chr(34) + "FrameType" + Chr(34) + " " + Chr(34) + "-" + Chr(34) + ";")

    'BA_:
    'TODO? for the moment, did not find the need
    For i = 1 To BA_List.Count
        stream.WriteLine BA_List(i)
    Next i

    'VAL_
    For i = 1 To Val_List.Count
        stream.WriteLine Val_List(i)
    Next i

    MsgBox (".dbc file created")

End Sub
