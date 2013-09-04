' Usage: csv_addbom_v2.vbs file.csv
' Notes:
' this isn't suitable for large files unless you have a lot of memory - ADODB.Stream reads the entire file into
' memory, then builds the output buffer in memory as well. #stupid

 
If WScript.Arguments.Count <> 1 Then
WScript.Echo "Usage: csv_addbom_v2.vbs file.csv"
WScript.Quit
End If
 
Dim fIn, fOut, sFilename, sBOM
sFilename = WScript.Arguments(0)
 
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
     
    fOut.SaveToFile "out1.csv", 2
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