Sub CAN_Message_Composer()


    Dim bit As Integer
    Dim ByteN As Integer
    Dim endian As String
    Dim size As Integer
    Dim DLC As Integer
    Dim temp As Integer
    Dim StartBit As Integer

    Worksheets("Tools").Activate
    Dim HeadersRangeD As Range: Set HeadersRangeD = Range("MessageComposer", Range("MessageComposer").End(xlToRight).Address)
    HeadersRangeD.Select
    Dim SizeRange As Range: Set SizeRange = Range(HeadersRangeD.Find("Size", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Size", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    Dim StartBitRange As Range: Set StartBitRange = Range(HeadersRangeD.Find("Start bit", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Start bit", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))
    endian = Range("MessageComposerEndianValue").value
    Dim DLCRange As Range: Set DLCRange = Range(HeadersRangeD.Find("DLC", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("DLC", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).End(xlDown))

    bit = 8 ' from 7 to 0
    ByteN = 0 'from 0 to 7
    size = 0
    DLC = 0 'final length of the message to contain all the signals

    If (endian = "Little Endian (Intel)") Then
        ' Little endian: Start bit here 0 = start bit in candb 0, the msb.
        ' x x x x x x x x
        ' msb          lsb
        bit = 0 ' from 7 (left) to 0 (right)

        For i = 2 To SizeRange.Count

            If (bit = 8) Then
                bit = 0
                ByteN = ByteN + 1
            End If

            StartBit = ByteN * 8 + bit

            size = SizeRange.Cells(i, 1).value

            If (size > (8 - bit)) Then 'need extra byte
                temp = size
                Do While (temp > (8 - bit)) ' add bytes if needed
                    temp = temp - (8 - bit)
                    ByteN = ByteN + 1
                    bit = 0
                Loop
                'StartBit = ByteN * 8 + bit
                bit = bit + temp + 1

            Else 'just place it in the existing byte
                'StartBit = ByteN * 8 + bit
                bit = bit + size

            End If

            StartBitRange.Cells(i, 1).value = StartBit

        Next i

        'DLC
        DLC = ByteN + 1
        DLCRange.Cells(2, 1).value = DLC

    ElseIf (endian = "Big Endian (Motorola)") Then

        'TODO this is the right calculation, as you would see in candb. but when writing startbit x in the dbc text file, with the big endian tag, it translates it into another value in candb. Still don't know how that works

        ' Big endian: Start bit here 7 => start bit in candb is 0, the lsb
        '              msb
        ' . . . . . . . x
        ' x x x x x x x .
        '            lsb

        bit = 7 ' from 7 (left) to 0 (right)
        For i = 2 To SizeRange.Count
            If (bit = -1) Then
                ByteN = ByteN + 1
                bit = 7
            End If

            size = SizeRange.Cells(i, 1).value
            If (size > (bit + 1)) Then 'need extra bytes
                temp = size
                Do While (temp > (bit + 1)) ' add bytes if needed
                    temp = temp - (bit + 1)
                    ByteN = ByteN + 1
                    bit = 7
                Loop

                StartBit = ByteN * 8 + bit + 1 - temp
                bit = bit - temp

            Else 'just place it in the existing byte
                StartBit = (ByteN * 8) + bit + 1 - size
                bit = bit - size
            End If

            StartBitRange.Cells(i, 1).value = StartBit

        Next i

        'DLC
        DLC = ByteN + 1
        DLCRange.Cells(2, 1).value = DLC

         '   temp = 8 - size
          '  Do While (temp < 0)
           '     ByteN = ByteN + 1
           '     temp = Abs(temp)
           '     temp = 8 - temp
           ' Loop
           ' bit = temp


    Else
        MsgBox ("endian unknown?")

    End If

End Sub
