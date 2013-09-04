' Usage1: csv_addbom_v3.vbs file.csv
' Usage2: drag a input csv file and drop on csv_addbom_v3.vbs
' Notes:
' this isn't suitable for large files unless you have a lot of memory - ADODB.Stream reads the entire file into
' memory, then builds the output buffer in memory as well. #stupid

 
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: csv_addbom_v3.vbs file.csv"
    WScript.Quit
End If
 
Dim fIn, fOut, sFilename, sBOM
sFilename = WScript.Arguments(0)

Dim fs, file_path
Set fs = CreateObject("Scripting.FileSystemObject")
file_path = fs.GetParentFolderName(sFilename) & "\"

 
Set fIn = CreateObject("adodb.stream")
fIn.Type = 1 'adTypeBinary
fIn.Mode = 3 'adModeRead
fIn.Open
fIn.LoadFromFile sFilename

 
sBOM = fIn.Read(5)
' UTF8 BOM is 0xEF,0xBB,0xBF (decimal 239, 187, 191)
If AscB(MidB(sBOM, 1, 1)) = 239 _
And AscB(MidB(sBOM, 2, 1)) = 187 _
And AscB(MidB(sBOM, 3, 1)) = 191 Then

MsgBox "UTF-8-BOM check: The input file has BOM already."

Else
    fIn.Position = 0

    Set fOut = CreateObject("adodb.stream")
    fOut.Type = 1 'adTypeBinary
    fOut.Mode = 3 'adModeReadWrite
    fOut.Open

    DIM sT, sB
    sT = chrB(239) & chrB(187) & chrB(191)
    sB = MultiByteToBinary(sT)
    fOut.Write sB    'Add UTF8_BOM
    
    fIn.CopyTo fOut

    DIM filename
    filename = DatePart("yyyy",Date) & "-" & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)_
    & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".csv"
    
    
    If file_path ="\" Then
        fOut.SaveToFile filename, 2
        MsgBox "File was saved to: " & vbCrLf & vbCrLf & filename
    Else
        fOut.SaveToFile file_path & filename, 2
        MsgBox "File was saved to: " & vbCrLf & vbCrLf & file_path & filename
    End If
    
    fOut.Flush
    fOut.Close

    
End If


Function MultiByteToBinary(MultiByte)
    'c 2000 Antonin Foller, http://www.motobit.com
    ' MultiByteToBinary converts multibyte string To real binary data (VT_UI1 | VT_ARRAY)
    ' Using recordset
    Dim RS, LMultiByte, Binary
    Const adLongVarBinary = 205  'ADO data type: OLEObject
    Set RS = CreateObject("ADODB.Recordset")
    LMultiByte = LenB(MultiByte)
    If LMultiByte>0 Then
        RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
        RS.Open
        RS.AddNew
        RS("mBinary").AppendChunk MultiByte & ChrB(0)  'ASCII 0 => Null
        RS.Update
        Binary = RS("mBinary").GetChunk(LMultiByte)
    End If
    MultiByteToBinary = Binary
End Function