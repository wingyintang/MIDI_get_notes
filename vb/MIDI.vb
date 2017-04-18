Public Class MIDI
    Public header_length As Integer = 6
    'length of the header chunk (always 6 bytes long--the size of the next three fields which are considered the header chunk).
    Public format As Integer = 0
    '0 = single track file format 
    '1 = multiple track file format 
    '2 = multiple song file format (i.e., a series of type 0 files)
    Public n_value As Integer = 0
    Property n As Integer
        Get
            Return n_value
        End Get
        Set(value As Integer)
            n_value = value
            Array.Resize(track_chunk, n)
        End Set
    End Property
    'number of track chunks that follow the header chunk
    Public div As Integer = 0
    Property division As Integer
        Get
            Return div
        End Get
        Set(value As Integer)
            div = value
            tempolist(0).onedelta = 120 / (1000000 * value)
        End Set
    End Property
    'unit of time for delta timing. If the value is positive, then it represents the units per beat. For example, +96 would mean 96 ticks per beat. If the value is negative, delta times are in SMPTE compatible units.
    Public track_chunk(-1) As MIDITrack
    Function ByteToNote(b As Byte) As String
        Return Chr(64 + Math.Floor(((1.15 * (b Mod 12) + 0.26) / 2) + 2) Mod 7 + 1) & If(Math.Floor(1.15 * (b Mod 12) + 0.26) Mod 2 = 1, "#", "") & ((b \ 12) - 1).ToString
    End Function
    Public notelist(-1) As Notes
    Public tempolist() As Tempo = {New Tempo With {.starttimesec = 0, .tempo = 120, .time = 0}}
    Sub sortnotes()
        MergeSort.sort(notelist, 0, notelist.Length - 1, "starttimesec")
    End Sub
    Sub addTempo(delta As Integer, tempo As Integer)
        Array.Resize(tempolist, tempolist.Length + 1)
        tempolist(tempolist.Length - 1) = New Tempo
        tempolist(tempolist.Length - 1).onedelta = tempo / (1000000 * division)
        tempolist(tempolist.Length - 1).time = delta
        tempolist(tempolist.Length - 1).tempo = tempo
        Dim time As Double = tempolist(tempolist.Length - 2).starttimesec + tempolist(tempolist.Length - 2).onedelta * (delta - tempolist(tempolist.Length - 2).time)
        tempolist(tempolist.Length - 1).starttimesec = time
    End Sub
    Function delta2time(delta As Integer) As Double
        Dim finding As Integer = -1
        Dim leave As Boolean = False
        While Not leave AndAlso finding < tempolist.Length - 1
            If delta >= tempolist(finding + 1).time Then
                finding += 1
            Else
                leave = True
            End If
        End While
        Dim result As Double = tempolist(finding).starttimesec + tempolist(finding).onedelta * (delta - tempolist(finding).time)
        Return result
    End Function
    Sub addNotesStart(delta As Integer, pitch As Integer, channel As Integer, track As Integer, velocity As Integer)
        Array.Resize(notelist, notelist.Length + 1)
        notelist(notelist.Length - 1) = New Notes(delta, pitch, channel, track, velocity)
        notelist(notelist.Length - 1).starttimesec = delta2time(delta)
    End Sub
    Sub addNotesEnd(delta As Integer, pitch As Integer, channel As Integer, track As Integer)
        Dim found = False
        Dim i As Integer = 0
        While Not found AndAlso i <= notelist.Length - 1
            If notelist(i).endtime = -1 And notelist(i).pitch = pitch AndAlso notelist(i).channel = channel AndAlso notelist(i).track = track Then
                notelist(i).endtime = delta
                notelist(i).endtimesec = delta2time(delta)
                found = True
            End If
            i += 1
        End While
    End Sub
    Sub EndallNotes(delta As Integer)
        For i = 0 To notelist.Length - 1
            If notelist(i).endtime = -1 Then
                notelist(i).endtime = delta
                notelist(i).endtimesec = delta2time(delta)
            End If
        Next
    End Sub
End Class

Public Class MIDITrack
    Public length As Integer = 0
    Public track_event(-1) As MIDIEvent
    Public eventcount As Integer = 0
    Sub addEvent()
        Array.Resize(track_event, track_event.Length + 1)
        track_event(track_event.Length - 1) = New MIDIEvent
        eventcount += 1
    End Sub
    Sub removeEvent()
        Array.Resize(track_event, track_event.Length - 1)
        eventcount -= 1
    End Sub
End Class
Public Class Tempo
    Public time As Integer = -1
    Public tempo As Integer = 0
    Public onedelta As Double = 0
    Public starttimesec As Double = 0
End Class
Public Class Notes
    Public starttime As Integer = -1
    Public starttimesec As Double = -1
    Public endtime As Integer = -1
    Public endtimesec As Double = -1
    Public pitch As Integer = -1 '0 to 127
    Public duration As Double = -1
    Public channel As Integer = -1 '0 to 15
    Public track As Integer = -1 '0 to 127
    Public velocity As Byte = 0
    Sub New(delta As Integer, p As Integer, c As Integer, t As Integer, v As Integer)
        starttime = delta
        pitch = p
        channel = c
        track = t
        velocity = v
    End Sub

End Class

Public Class MIDIEvent
    Public length As Integer = 0
    Public v_time As Integer = 0
    Public v_length As Integer = 0
    Public eventcode As Byte = 0
    Public meta_type As Byte = 0
    Public event_data_bytes As String = ""
    Public channel As Byte = 0
    Public note_num As Byte = 0
    Public velocity As Byte = 0
End Class

Module MergeSort
    Public Sub sort(Of T)(ByRef a() As T, ByVal l As Integer, ByVal r As Integer, ByVal n As String)
        If l < r Then
            Dim m As Integer = (l + r) \ 2
            sort(Of T)(a, l, m, n)
            sort(Of T)(a, m + 1, r, n)
            merge(Of T)(a, l, r, n)
        End If
    End Sub
    Private Sub merge(Of T)(ByRef a() As T, ByVal l As Integer, ByVal r As Integer, ByVal n As String)
        Dim m As Integer = (l + r) \ 2
        Dim temp(a.Length - 1) As T
        Dim x, y, z As Integer
        x = l
        y = m + 1
        z = l
        While (x <= m) And (y <= r)
            If comparison(a(x), a(y), n) Then
                temp(z) = a(x)
                x += 1
            Else
                temp(z) = a(y)
                y += 1
            End If
            z += 1
        End While
        If x > m Then
            For i As Integer = y To r
                temp(z) = a(i)
                z += 1
            Next
        Else
            For i As Integer = x To m
                temp(z) = a(i)
                z += 1
            Next
        End If
        For i As Integer = l To r
            a(i) = temp(i)
        Next
    End Sub

    Private Function comparison(ByVal c1 As Object, ByVal c2 As Object, ByVal n As String) As Boolean
        Select Case n.ToUpper
            Case "NAME"
                Return c1.name < c2.name
            Case "STRING"
                Return c1.ToString < c2.ToString
            Case "VALUE"
                Return Val(c1) < Val(c2)
            Case "STARTTIMESEC"
                Return c1.starttimesec < c2.starttimesec
            Case ""
                Return c1 < c2
            Case Else
                Return c1 < c2
        End Select
    End Function
End Module
