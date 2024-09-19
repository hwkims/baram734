Attribute VB_Name = "modDat"
' Baram Dat Edit Module,
' Last Modified 2007. 5. 17. PM 7:10
' By Moon Eun Jung.

Type DatType
    Name As String
    OffSet As Long
    Size As Long
    Data() As Byte
End Type

Private Declare Sub Copy Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Dim DatFiles() As DatType
Public Function GetFileName(ByVal Name As String) As String
    On Error GoTo Err
        Dim Splits() As String
        Splits = Split(Name, "\")
        GetFileName = Splits(UBound(Splits))
    Exit Function
Err:
    GetFileName = ""
End Function
Public Sub AddDat(ByVal Name As String)
    On Error GoTo Err
        ReDim Buffer(FileLen(Name) - 1) As Byte
        Close #1
        Open Name For Binary Access Read As #1
            Get #1, , Buffer
        Close #1
        Dim Temp As DatType
        Temp.Name = GetFileName(Name)
        Temp.Size = UBound(Buffer) + 1
        Temp.Data = Buffer
        If UBound(DatFiles) = 0 Then
            If DatFiles(0).OffSet = 0 And DatFiles(0).Size = 0 Then
                DatFiles(0) = Temp
            Else
                GoTo Add
            End If
        Else
Add:
            ReDim Preserve DatFiles(UBound(DatFiles) + 1)
            DatFiles(UBound(DatFiles)) = Temp
        End If
        RefreshDat
    Exit Sub
Err:
    MsgBox "에러가 발생하였습니다.", vbApplicationModal + vbCritical, "Error!"
End Sub
Public Sub RemoveDat(ByVal DatId As Long)
    On Error Resume Next
    If UBound(DatFiles) = 0 Then
        ClearDat
    Else
        ReDim NewDatFiles(UBound(DatFiles) - 1) As DatType
        Dim i As Long, j As Long
        For i = 0 To UBound(DatFiles)
            If i <> DatId Then
                NewDatFiles(j) = DatFiles(i)
                j = j + 1
            End If
        Next i
        DatFiles = NewDatFiles
    End If
    RefreshDat
End Sub
Public Sub RefreshDat()
    On Error Resume Next
    frmMain.lstDat.ListItems.Clear
    Dim i As Long
    If UBound(DatFiles) = 0 Then
        If DatFiles(0).OffSet = 0 And DatFiles(0).Size = 0 Then
            Exit Sub
        End If
    End If
    For i = 0 To UBound(DatFiles)
        frmMain.lstDat.ListItems.Add(, , DatFiles(i).Name).SubItems(1) = DatFiles(i).Size & "바이트"
    Next i
End Sub
Public Sub ClearDat()
    On Error Resume Next
    frmMain.lstDat.ListItems.Clear
    ReDim DatFiles(0)
    DatFiles(0).Name = ""
    DatFiles(0).OffSet = 0
    DatFiles(0).Size = 0
End Sub
Public Sub OpenDat(ByVal Name As String)
    On Error GoTo Err
    ClearDat
    Dim TotalCount As Long, i As Long, totSize As Long
    ReDim Buffer(FileLen(Name) - 1) As Byte
    Close #1
    Open Name For Binary Access Read As #1
        Get #1, , Buffer
        Seek #1, 1
        Get #1, , TotalCount
    ReDim DatFiles(TotalCount - 2)
    For i = 0 To (TotalCount - 1)
        Dim Temp As DatType, TempName(0 To 12) As Byte
            Get #1, , Temp.OffSet
            Get #1, , TempName '13
            If i > 0 Then
                DatFiles(i - 1).Size = Temp.OffSet - DatFiles(i - 1).OffSet
                ReDim DatFiles(i - 1).Data(DatFiles(i - 1).Size - 1)
                Call Copy(DatFiles(i - 1).Data(0), Buffer(DatFiles(i - 1).OffSet), DatFiles(i - 1).Size)
            End If
        Temp.Name = StrConv(TempName, vbUnicode)
        Temp.Name = Left$(Temp.Name, InStr(1, Temp.Name, Chr$(0)) - 1)
        If i = (TotalCount - 1) Then Exit For
        DatFiles(i) = Temp
    Next i
    Close #1
    RefreshDat
    Exit Sub
Err:
    ClearDat
    MsgBox "잘못된 Dat 파일입니다.", vbApplicationModal + vbCritical, "Error!"
End Sub
Public Sub SaveDat(ByVal Name As String)
    On Error GoTo Err
    If Dir(Name) <> "" Then Kill Name
    Close #1
    Open Name For Binary Access Write As #1
    Put #1, , (UBound(DatFiles) + 2)
    Dim i As Long, totSize As Long, HeadSize As Long
    '*Making Header
    HeadSize = 17 * (UBound(DatFiles) + 2) + 4
    totSize = 0
    For i = 0 To UBound(DatFiles)
        Put #1, , (HeadSize + totSize)
            Dim Temp() As Byte
                Temp = StrConv(DatFiles(i).Name, vbFromUnicode)
                ReDim Preserve Temp(12) As Byte 'Keep Size 0 to 12 (13 Bytes)
        Put #1, , Temp
        totSize = totSize + DatFiles(i).Size
    Next i
    '*Making EOF!
        'DataEOF Structure!
        '[0] TotSize
        '[1] (0)*6 (255)*4 (0)*3 'Name
    Put #1, , (HeadSize + totSize)
        Dim EoF(0 To 12) As Byte
            EoF(0) = 0
            EoF(1) = 0
            EoF(2) = 0
            EoF(3) = 0
            EoF(4) = 0
            EoF(5) = 0
            EoF(6) = 255
            EoF(7) = 255
            EoF(8) = 255
            EoF(9) = 255
            EoF(10) = 0
            EoF(11) = 0
            EoF(12) = 0
    Put #1, , EoF
    '*Put Data,
    For i = 0 To UBound(DatFiles)
        Put #1, , DatFiles(i).Data
    Next i
    ReDim Temp(0 To 7)
    Put #1, , Temp
    Close #1
    Exit Sub
Err:
    MsgBox "에러가 발생하였습니다.", vbApplicationModal + vbCritical, "Error!"
End Sub
Public Sub ExtractDat(ByVal DatId As Long, ByVal Name As String)
    On Error GoTo Err
    If (UBound(DatFiles) + 1) < DatId Then Exit Sub
    If Dir(Name) <> "" Then Kill Name
        Close #1
        Open Name For Binary Access Write As #1
            Put #1, , DatFiles(DatId).Data
        Close #1
    Exit Sub
Err:
    MsgBox "에러가 발생하였습니다.", vbApplicationModal + vbCritical, "Error!"
End Sub
