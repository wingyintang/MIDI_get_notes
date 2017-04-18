Public Class Form1
    Dim hr As HexReader
    Dim outtext As String = ""
    Dim count1 As Integer = 0
    Dim count2 As Integer = 0
    Dim count3 As Integer = 0
    Sub write(t1 As String, Optional write As Boolean = False)
        'If write Then
        '    ' TextBox1.Text &= hr.stored.PadRight(40, ".")
        '    Debug.Write(hr.stored.PadRight(40, "."))
        'End If
        '' TextBox1.Text &= t1
        'Debug.Write(t1)
    End Sub
    Sub writeln(t1 As String, Optional write As Boolean = False)
        'If write Then
        '    ' TextBox1.Text &= hr.stored.PadRight(40, ".")
        '    Debug.Write(hr.stored.PadRight(40, "."))
        'End If
        ''TextBox1.Text &= t1 & vbCrLf
        'Debug.Write(t1 & vbCrLf)
    End Sub
    Sub writeln2(t1 As String, t2 As String)
        'TextBox1.Text &= hr.stored.PadRight(40, ".") & t1.PadRight(40, ".") & t2 & vbCrLf
        'Debug.Print(hr.stored.PadRight(40, ".") & t1.PadRight(40, ".") & t2 & vbCrLf)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            outtext = ""
            hr = New HexReader(OpenFileDialog1.FileName)
            TextBox1.Clear()
            TextBox1.Text = "Waiting."
            Me.Refresh()
            Dim mymidi As New MIDI
            'Try
            count1 = 0
            count2 = 0
            count3 = 0
            Dim deltasum As Integer = 0
            Dim b As String = hr.ReadTByte(4)
            writeln("<Header Chunk>")
            If b = "MThd" Then
                writeln2("Text: MThd", "<chunk type> = MThd: Valid MIDI")
            Else
                writeln2("Text: " & b, "MThd: Invalid MIDI")
                Exit Sub
            End If
            b = hr.ReadSByte(4)
            If b = "00000006" Then
                writeln2("Hex: 00000006", "<length> = 6 Head Trunk Length")
                mymidi.header_length = 6
            Else
                writeln2("Hex: " & b, "Invalid Head Trunk Length")
                Exit Sub
            End If
            mymidi.format = hr.ReadSIByte(2)
            Select Case mymidi.format
                Case "0"
                    writeln2("Integer: " & mymidi.format, "<format>=0: single multi-channel track")
                Case "1"
                    writeln2("Integer: " & mymidi.format, "<format>=1: one or more simultaneous tracks (or MIDI outputs) of a sequence")
                Case "2"
                    writeln2("Integer: " & mymidi.format, "<format>=2: one or more sequentially independent single-track patterns")
                Case Else
                    writeln2("Integer: " & mymidi.format, "Invalid Head Trunk Length")
                    Exit Sub
            End Select
            mymidi.n = hr.ReadSIByte(2)
            writeln2("n: " & mymidi.n, mymidi.n & " track chunks follow the header chunk")
            mymidi.division = hr.ReadSIByte(2)
            writeln2("division: " & mymidi.division, mymidi.division & " ticks per beat")
            writeln("")
            For i As Integer = 0 To mymidi.n - 1
                Dim notest As Boolean = False
                Dim channel_using As Integer = 0
                deltasum = 0
                writeln("<Track Chunk " & (i + 1).ToString & " Start>")
                b = hr.ReadTByte(4)
                If b = "MTrk" Then
                    writeln2("Text: MTrk", "<chunk type> = MTrk: Valid Track")
                Else
                    writeln2("Text: " & b, "MTrk: Invalid Track")
                    Exit Sub
                End If
                mymidi.track_chunk(i) = New MIDITrack
                mymidi.track_chunk(i).length = hr.ReadSIByte(4)
                writeln2("Integer: " & mymidi.track_chunk(i).length, mymidi.track_chunk(i).length & " bytes in the track chunk following")
                hr.skip = mymidi.track_chunk(i).length
                While hr.skip > 0
                    mymidi.track_chunk(i).addEvent()
                    writeln("<Event " & mymidi.track_chunk(i).eventcount.ToString & " Start>")
                    Dim bits As Long = 0
                    bits = hr.Read7Bit
                    mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).v_time = bits
                    writeln2("v_time: " & bits, "delta time: " & bits)
                    deltasum += bits
                    Me.Text = deltasum
                    Dim bytes As Byte = hr.ReadByte(True)
                    mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).eventcode = bytes
                    'writeln2("event: " & bytes, "Event: " & bytes)
                    Select Case bytes \ 16
                        Case 8 '8 Note Off
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).channel = bytes Mod 16
                            channel_using = bytes Mod 16
                            writeln2("Type: 8", "Note Off, Channel " & (bytes Mod 16).ToString)
                            Dim pitch As Byte = hr.ReadByte(True)
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).note_num = pitch
                            writeln2("Note: " & pitch, "which is " & mymidi.ByteToNote(pitch))
                            mymidi.addNotesEnd(deltasum, pitch, bytes Mod 16, i)
                            bytes = hr.ReadByte(True)
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).velocity = bytes
                            writeln("Velocity: " & bytes, True)
                            count2 += 1
                        Case 9 '9 Note On
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).channel = bytes Mod 16
                            channel_using = bytes Mod 16
                            writeln2("Type: 9", "Note On, Channel " & (bytes Mod 16).ToString)
                            Dim pitch As Byte = hr.ReadByte(True)
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).note_num = pitch
                            writeln2("Note: " & pitch, "which is " & mymidi.ByteToNote(pitch))
                            Dim velocity As Byte = hr.ReadByte(True)
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).velocity = velocity
                            writeln("Velocity: " & velocity, True)
                            If velocity > 0 Then
                                mymidi.addNotesStart(deltasum, pitch, bytes Mod 16, i, velocity)
                                count1 += 1
                            Else
                                mymidi.addNotesEnd(deltasum, pitch, bytes Mod 16, i)
                                count2 += 1
                            End If
                            notest = True
                        Case 10 'A Polyphonic Key Pressure 
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).channel = bytes Mod 16
                            writeln2("Type: 10", "Polyphonic Key Pressure, Channel " & (bytes Mod 16).ToString)
                            b = hr.ReadSByte(2)
                            writeln("Data byte", True)
                        Case 11 'B Control Change
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).channel = bytes Mod 16
                            writeln2("Type: 11", "Control Change, Channel " & (bytes Mod 16).ToString)
                            b = hr.ReadSByte(2)
                            writeln("Data byte", True)
                            notest = False
                        Case 12 'C Program Change
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).channel = bytes Mod 16
                            writeln2("Type: 12", "Program Change, Channel " & (bytes Mod 16).ToString)
                            b = hr.ReadSByte(1)
                            writeln("Data byte", True)
                        Case 13 'D Channel Pressure (After-touch)
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).channel = bytes Mod 16
                            writeln2("Type: 13", "Channel Pressure, Channel " & (bytes Mod 16).ToString)
                            b = hr.ReadSByte(1)
                            writeln("Data byte", True)
                        Case 14 'E Pitch Wheel Change
                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).channel = bytes Mod 16
                            writeln2("Type: 14", "Pitch Wheel Change, Channel " & (bytes Mod 16).ToString)
                            b = hr.ReadSByte(2)
                            writeln("Data byte", True)
                            notest = False
                        Case 15 'FX
                            writeln2("event: " & bytes, "Event: " & bytes)
                            Select Case bytes Mod 16
                                Case 0 'F0: sysex_event until another F7 is found
                                    writeln("meta_event_type: " & bytes, True)
                                    bits = hr.Read7Bit
                                    mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).v_length = bits
                                    writeln("event_length: " & bits, True)
                                    b = hr.ReadSByte(bits)
                                    mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).event_data_bytes = b
                                    writeln("sysex_event: " & b, True)
                                Case 7 'F7: sysex_event until another F7 is found
                                    writeln("meta_event_type: " & bytes, True)
                                    bits = hr.Read7Bit
                                    mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).v_length = bits
                                    writeln("event_length: " & bits, True)
                                    b = hr.ReadSByte(bits)
                                    mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).event_data_bytes = b
                                    writeln("sysex_event: " & b, True)
                                Case 15 'FF: meta_event
                                    bytes = hr.ReadByte(True)
                                    writeln("meta_event_type: " & bytes, True)
                                    Dim _7bits As Long = hr.Read7Bit
                                    mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).v_length = _7bits
                                    writeln("event_length: " & _7bits, True)
                                    Select Case bytes
                                        Case 81 '51 Tempo setting
                                            Dim tem As Integer = hr.ReadSIByte(_7bits)
                                            mymidi.addTempo(deltasum, tem)
                                        Case Else
                                            b = hr.ReadSByte(_7bits)
                                            mymidi.track_chunk(i).track_event(mymidi.track_chunk(i).eventcount - 1).event_data_bytes = b
                                            writeln("event_data_bytes: " & b, True)
                                            'Case 0 '00 Sequence number
                                            'Case 1 '01 Text event
                                            'Case 2 '02 Copyright notice
                                            'Case 3 '03 Sequence or track name
                                            'Case 4 '04 Instrument name
                                            'Case 5 '05 Lyric text
                                            'Case 6 '06 Marker text
                                            'Case 7 '07 Cue point
                                            'Case 32 '20 MIDI channel prefix assignment
                                            'Case 47 '2F End of track
                                            'Case 84 '54 SMPTE offset
                                            'Case 88 '58 Time signature
                                            'Case 89 '59 Key signature
                                            'Case 127 '7F Sequencer specific event
                                    End Select
                                Case Else
                                    writeln("Exit: meta_event_type: " & "F-" & (bytes Mod 16).ToString, True)
                                    Exit While
                            End Select
                        Case Else '0 to 7 for Notes 'MIDI Controller Messages
                            If notest Then
                                writeln("Type (Note On/Off): " & bytes, True)
                                Dim velocity As Byte = hr.ReadSIByte(1)
                                If velocity > 0 Then
                                    mymidi.addNotesStart(deltasum, bytes, channel_using, i, velocity)
                                    count1 += 1
                                Else
                                    mymidi.addNotesEnd(deltasum, bytes, channel_using, i)
                                    count2 += 1
                                End If
                                writeln("Note on/off", True)
                            Else
                                'If bytes >= 123 And bytes <= 127 Then
                                '    mymidi.EndallNotes(deltasum)
                                '    count3 += 1
                                'End If
                                bytes = hr.ReadByte()
                                writeln("MIDI Controller Message", True)
                            End If
                    End Select
                    writeln("<Event " & mymidi.track_chunk(i).eventcount.ToString & " End>")
                End While
                'End Track
                hr.jump()
                writeln("Skipped", True)
                writeln("<Track Chunk " & (i + 1).ToString & " End>")
                writeln("")
            Next
            writeln("<Eof> End of file")
            'Catch ex As Exception

            'End Try
            hr.Close()
            Debug.Print("Finished. Note on=" & count1 & ", note off=" & count2 & ", all note off=" & count3)
            'For i = 0 To mymidi.notelist.Length - 1
            '    Dim a As Notes = mymidi.notelist(i)
            '    Debug.Print("Note " & (i + 1).ToString & ": Channel=" & a.channel & ", Track=" & a.track & ", Start=" & a.starttime & "(" & a.starttimesec & " sec), End=" & a.endtime & "(" & a.endtimesec & " sec), Pitch=" & mymidi.ByteToNote(a.pitch) & ", Velocity=" & a.velocity)
            'Next
            mymidi.sortnotes()
            Debug.Print("Start")
            For i = 0 To mymidi.notelist.Length - 1
                Debug.Write(Math.Round(mymidi.notelist(i).starttimesec, 3) & ",")
            Next
            Debug.WriteLine("")
            Debug.Print("End")
            For i = 0 To mymidi.notelist.Length - 1
                Debug.Write(Math.Round(mymidi.notelist(i).endtimesec, 3) & ",")
            Next
            Debug.WriteLine("")
            Debug.Print("Pitch")
            For i = 0 To mymidi.notelist.Length - 1
                Debug.Write(mymidi.notelist(i).pitch & ",")
            Next
            Debug.WriteLine("")
            Debug.Print("Track")
            For i = 0 To mymidi.notelist.Length - 1
                Debug.Write(mymidi.notelist(i).channel & ",")
            Next
            Debug.WriteLine("")
            Me.Refresh()
            TextBox1.Text = outtext
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub
End Class
