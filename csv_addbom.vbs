' Usage: csv_addbom.vbs file.csv
' Notes:
' 必須要有initial.csv檔案放在同一個目錄
' this isn't suitable for large files unless you have a lot of memory - ADODB.Stream reads the entire file into
' memory, then builds the output buffer in memory as well. #stupid

 
If WScript.Arguments.Count <> 1 Then
WScript.Echo "Usage: ukik_addbom.vbs file.csv"
WScript.Quit
End If
 
Dim fIn, fOut, sFilename, fBOM, sBOM
sFilename = WScript.Arguments(0)
 
Set fIn = CreateObject("adodb.stream")
fIn.Type = 1 'adTypeBinary
fIn.Mode = 3 'adModeRead
fIn.Open
fIn.LoadFromFile sFilename

Set fBOM = CreateObject("adodb.stream")
fBOM.Type = 1 'adTypeBinary
fBOM.Mode = 3 'adModeRead
fBOM.Open
fBOM.LoadFromFile "initial.csv"

 
sBOM = fBOM.Read(5)
' UTF8 BOM is 0xEF,0xBB,0xBF (decimal 239, 187, 191)
If AscB(MidB(sBOM, 1, 1)) = 239 _
And AscB(MidB(sBOM, 2, 1)) = 187 _
And AscB(MidB(sBOM, 3, 1)) = 191 Then

rem fIn.Position = 0

Set fOut = CreateObject("adodb.stream")
fOut.Type = 1 'adTypeBinary
fOut.Mode = 3 'adModeReadWrite
fOut.Open

fOut.Write sBOM 'Add UTF8_BOM
 
fIn.CopyTo fOut
 
fOut.SaveToFile "out.csv", 2
fOut.Flush
fOut.Close

End If
