Sub VRG_Audit()
Dim MARSHA As String
Dim Host
Dim screenshot As Integer
Dim mDate As Variant
Dim Restr As Variant
Dim RateP As Integer
Dim Row As Integer
Dim Row2 As Integer
Dim Day As Integer
Dim DName As String


Set Host = CreateObject("BZWhll.WhllObj")
Host.Connect ""

Host.sendkey ("<CLEAR>"): Host.waitready 10, 1

MARSHA = wsHome.Range("B2")

wsVRGdata.Range("C1") = Now()

For RateP = 1 To wsHome.Range("C17").End(xlUp).row - 6
Day = 1
DName = ""

    Host.sendkey "VRG" & MARSHA & "/" & wsHome.Range("C7:C17").Cells(RateP)
    Host.sendkey ("<ENTER>"): Host.waitready 10, 1

    For screenshot = 1 To 18
    
        If screenshot = 1 Then Row2 = 5 Else Row2 = 1
        For Row = Row2 To 20
            Host.readscreen mDate, 5, Row, 8
            Host.readscreen Restr, 9, Row, 39
            
            
            If RateP > 1 Then
                If InStr(DName, mDate) = 0 Then
                    wsVRGdata.Range("C3:L365").Cells(Day, RateP).Value = Restr
                    Day = Day + 1
                    DName = DName & mDate
                End If
            Else
                If InStr(DName, mDate) = 0 Then
                    wsVRGdata.Range("B3:B365").Cells(Day).Value = mDate
                    wsVRGdata.Range("C3:L365").Cells(Day, RateP).Value = Restr
                    Day = Day + 1
                    DName = DName & mDate
                End If
            End If
            
        Next Row

    Host.WriteScreen "md", 22, 2
   
    Host.sendkey ("<Enter>"): Host.waitready 10, 1
        
    Next screenshot

Next RateP


MsgBox "Audit has finished"

Exit Sub

NotConnected:
        MsgBox ("Could not establish Marsha Connection")