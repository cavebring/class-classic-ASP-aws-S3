<%
'=============================================================================================================
' Amazon Class Module
' ---------------------
'
' Created By  :  John Cavebring 
' Created date:  2014-03-29
'=============================================================================================================


'-- Amazon Web Services > My Account > Access Credentials > Access Keys --'
Const strAccessKeyID = "YOUR ID"
Const strSecretAccessKey = "YOUR SECRET KEY"
Const strLocalTempDir = "c:\temp\"

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'ררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררר
Class amazon


Private s3_strBinaryData 'As String
Private s3_strLocalFile 'As String
Private s3_strRemoteFile 'As String
Private s3_strBucket 'As String
Private s3_strOutFileName 'As String

	
'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'ררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררר
Private Sub Class_Initialize()
   
   
End Sub

Private Sub Class_Terminate()

End Sub
'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Return Property
'ררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררר

Public Property Get s3_LocalFile() 'As String
   s3_LocalFile = s3_strLocalFile
End Property

Public Property Get s3_RemoteFile() 'As String
   s3_RemoteFile = s3_strRemoteFile
End Property

Public Property Get s3_Bucket() 'As String
   s3_Bucket = s3_strBucket
End Property

Public Property Get s3_OutFileName() 'As String
   s3_OutFileName = s3_strOutFileName
End Property


Public Property Get s3_BinaryData() 'As String
   s3_BinaryData = s3_strBinaryData
End Property




'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Set Preperty
'ררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררר

Public Property Let s3_LocalFile(ByVal NewValue) 'As String)
   s3_strLocalFile = NewValue
End Property

Public Property Let s3_RemoteFile(ByVal NewValue) 'As String)
   s3_strRemoteFile = NewValue
End Property

Public Property Let s3_Bucket(ByVal NewValue) 'As String)
   s3_strBucket = NewValue
End Property

Public Property Let s3_OutFileName(ByVal NewValue) 'As String)
   s3_strOutFileName = NewValue
End Property

Public Property Let s3_BinaryData(ByVal NewValue) 'As String)
   s3_strBinaryData = NewValue
End Property




'___________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'ררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררר


Function NowInGMT()
  'This is probably not the best implementation, but it works ;-) --'
  Dim sh: Set sh = Server.CreateObject("WScript.Shell")
  Dim iOffset: iOffset = sh.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
  Dim dtNowGMT: dtNowGMT = DateAdd("n", iOffset, Now())
  Dim strDay: strDay = "NA"
  Select Case Weekday(dtNowGMT)
    Case 1 strDay = "Sun"
    Case 2 strDay = "Mon"
    Case 3 strDay = "Tue"
    Case 4 strDay = "Wed"
    Case 5 strDay = "Thu"
    Case 6 strDay = "Fri"
    Case 7 strDay = "Sat"
    Case Else strDay = "Error"
  End Select
  Dim strMonth: strMonth = "NA"
  Select Case Month(dtNowGMT)
    Case 1 strMonth = "Jan"
    Case 2 strMonth = "Feb"
    Case 3 strMonth = "Mar"
    Case 4 strMonth = "Apr"
    Case 5 strMonth = "May"
    Case 6 strMonth = "Jun"
    Case 7 strMonth = "Jul"
    Case 8 strMonth = "Aug"
    Case 9 strMonth = "Sep"
    Case 10 strMonth = "Oct"
    Case 11 strMonth = "Nov"
    Case 12 strMonth = "Dec"
    Case Else strMonth = "Error"
  End Select
  Dim strHour: strHour = CStr(Hour(dtNowGMT))
  If Len(strHour) = 1 Then strHour = "0" & strHour End If
  Dim strMinute: strMinute = CStr(Minute(dtNowGMT))
  If Len(strMinute) = 1 Then strMinute = "0" & strMinute End If
  Dim strSecond: strSecond = CStr(Second(dtNowGMT))
  If Len(strSecond) = 1 Then strSecond = "0" & strSecond End If
  Dim strNowInGMT: strNowInGMT = _
    strDay & _
    ", " & _
    Day(dtNowGMT) & _
    " " & _
    strMonth & _
    " " & _
    Year(dtNowGMT) & _
    " " & _
    strHour & _
    ":" & _
    strMinute & _
    ":" & _
    strSecond & _
    " +0000"
  NowInGMT = strNowInGMT
End Function



'-- GetBytesFromString --------------------------------------------------------'
Function GetBytesFromString(strValue)
  Dim stm: Set stm = Server.CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 2
  stm.Charset = "ascii"
  stm.WriteText strValue
  stm.Position = 0
  stm.Type = 1
  GetBytesFromString = stm.Read
  Set stm = Nothing
End Function



'-- HMACSHA1 ------------------------------------------------------------------'
Function HMACSHA1(strKey, strValue)
  Dim sha1: Set sha1 = Server.CreateObject("System.Security.Cryptography.HMACSHA1")
  sha1.key = GetBytesFromString(strKey)
  HMACSHA1 = sha1.ComputeHash_2(GetBytesFromString(strValue))
  Set sha1 = Nothing
End Function



'-- ConvertBytesToBase64 ------------------------------------------------------'
Function ConvertBytesToBase64(byteValue)
  Dim dom: Set dom = Server.CreateObject("MSXML2.DomDocument")
  Dim elm: Set elm = dom.CreateElement("b64")
  elm.dataType = "bin.base64"
  elm.nodeTypedValue = byteValue
  ConvertBytesToBase64 = elm.Text
  Set elm = Nothing
  Set dom = Nothing
End Function



'-- GetBytesFromFile ----------------------------------------------------------'
Function GetBytesFromFile(strFileName)
  Dim stm: Set stm = Server.CreateObject("ADODB.Stream")
  stm.Type = 1 'adTypeBinary --'
  stm.Open
  stm.LoadFromFile strFileName
  stm.Position = 0
  GetBytesFromFile = stm.Read
  stm.Close
  Set stm = Nothing
End Function

'-- GetBytesFromFile ----------------------------------------------------------'
Function GetBytesFromStream(strBinary)
  Dim stm: Set stm = Server.CreateObject("ADODB.Stream")
  stm.Type = 1 'adTypeBinary --'
  stm.Open
  stm.Write strBinary
  stm.Position = 0
  GetBytesFromStream = stm.Read
  stm.Close
  Set stm = Nothing
End Function






Public Function s3_Delete() ' as string (responsecode)

'-- Authentication: --'
Dim strNowInGMT: strNowInGMT = NowInGMT()
Dim strStringToSign: strStringToSign = _
  "DELETE" & vbLf & _
  "" & vbLf & _
  "text/xml" & vbLf & _
  strNowInGMT & vbLf & _
  "/" & s3_strBucket + "/" & s3_strRemoteFile
Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature



'-- Download: --'


Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
xhttp.open "DELETE", "http://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
xhttp.setRequestHeader "Content-Type", "text/xml"
xhttp.setRequestHeader "Date", strNowInGMT 'Yes, this line is mandatory ;-) --'
xhttp.setRequestHeader "Authorization", strAuthorization
xhttp.send

If xhttp.status = "204" Then '204 = delete ok'
  s3_Delete = "1"
Else
  s3_Delete = "0:" & xhttp.responseText
End If

Set xhttp = Nothing
'-- NowInGMT ------------------------------------------------------------------'


End Function






Public Function s3_UploadBinary() ' as string (responsecode)

'-- Authentication: --'
Dim strNowInGMT: strNowInGMT = NowInGMT()
Dim strStringToSign: strStringToSign = _
  "PUT" & vbLf & _
  "" & vbLf & _
  "text/xml" & vbLf & _
  strNowInGMT & vbLf & _
  "/" & s3_strBucket + "/" & s3_strRemoteFile
Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature
'-- Upload: --'
Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
xhttp.open "PUT", "http://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
xhttp.setRequestHeader "Content-Type", "text/xml"
xhttp.setRequestHeader "Date", strNowInGMT 'Yes, this line is mandatory ;-) --'
xhttp.setRequestHeader "Authorization", strAuthorization
xhttp.send GetBytesFromStream(s3_strBinaryData)

If xhttp.status = "200" Then
  s3_Upload = "1"
Else
  s3_Upload = "0:" & xhttp.responseText
End If

Set xhttp = Nothing
'-- NowInGMT ------------------------------------------------------------------'



End Function






Public Function s3_Upload() ' as string (responsecode)

'-- Authentication: --'
Dim strNowInGMT: strNowInGMT = NowInGMT()
Dim strStringToSign: strStringToSign = _
  "PUT" & vbLf & _
  "" & vbLf & _
  "text/xml" & vbLf & _
  strNowInGMT & vbLf & _
  "/" & s3_strBucket + "/" & s3_strRemoteFile
Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature
'-- Upload: --'
Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
xhttp.open "PUT", "http://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
xhttp.setRequestHeader "Content-Type", "text/xml"
xhttp.setRequestHeader "Date", strNowInGMT 'Yes, this line is mandatory ;-) --'
xhttp.setRequestHeader "Authorization", strAuthorization
xhttp.send GetBytesFromFile(s3_strLocalFile)

If xhttp.status = "200" Then
  s3_Upload = "1"
Else
  s3_Upload = "0:" & xhttp.responseText
End If

Set xhttp = Nothing
'-- NowInGMT ------------------------------------------------------------------'



End Function




Public Function s3_Download() ' as string (responsecode)

'-- Authentication: --'
Dim strNowInGMT: strNowInGMT = NowInGMT()
Dim strStringToSign: strStringToSign = _
  "GET" & vbLf & _
  "" & vbLf & _
  "text/xml" & vbLf & _
  strNowInGMT & vbLf & _
  "/" & s3_strBucket + "/" & s3_strRemoteFile
Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature



'-- Download: --'


Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
xhttp.open "GET", "http://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
xhttp.setRequestHeader "Content-Type", "text/xml"
xhttp.setRequestHeader "Date", strNowInGMT 'Yes, this line is mandatory ;-) --'
xhttp.setRequestHeader "Authorization", strAuthorization
xhttp.send

If xhttp.status = "200" Then

	Set oStream = Server.CreateObject("ADODB.Stream")
	oStream.Open
	oStream.Type = 1
	oStream.Write xhttp.responseBody
	oStream.SaveToFile s3_strLocalFile, 2
	oStream.Close

  s3_Download = "1"
Else
  s3_Download = "0:" & xhttp.responseText
End If

Set xhttp = Nothing
'-- NowInGMT ------------------------------------------------------------------'



End Function



Public Function s3_StreamToBrowser() ' as string (responsecode)

'-- Authentication: --'
Dim strNowInGMT: strNowInGMT = NowInGMT()
Dim strStringToSign: strStringToSign = _
  "GET" & vbLf & _
  "" & vbLf & _
  "text/xml" & vbLf & _
  strNowInGMT & vbLf & _
  "/" & s3_strBucket + "/" & s3_strRemoteFile
Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature



'-- Download: --'


Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
xhttp.open "GET", "http://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
xhttp.setRequestHeader "Content-Type", "text/xml"
xhttp.setRequestHeader "Date", strNowInGMT 'Yes, this line is mandatory ;-) --'
xhttp.setRequestHeader "Authorization", strAuthorization
xhttp.send

If xhttp.status = "200" Then

	Set oStream = Server.CreateObject("ADODB.Stream")
	oStream.Open
	oStream.Type = 1
	oStream.Write xhttp.responseBody


	TempFile = strLocalTempDir & timer


	oStream.SaveToFile TempFile & fname, 2





	select case lcase(right(s3_strOutFileName,3))
	case "pdf"
	Response.ContentType = "application/pdf"
	case "htm","tml"
	Response.ContentType = "text/HTML"
	case "gif"
	Response.ContentType = "image/GIF"
	case "jpg","peg"
	Response.ContentType = "image/JPEG"
	case "txt"
	Response.ContentType = "text/plain"
	case "zip"
	Response.ContentType = "application/zip"
	case Else
	Response.ContentType = "application/octet-stream"
	end select




	Response.Charset = "UTF-8"

	Response.AddHeader "Content-Disposition", "attachment; filename="& s3_strOutFileName


	oStream.LoadFromFile(TempFile)




	do while not oStream.EOS
	response.binaryWrite oStream.read(3670016) 
	response.flush
	loop


	oStream.Close



        Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.Fileexists(TempFile) Then objFSO.DeleteFile TempFile
        Set objFSO = Nothing



End If

Set xhttp = Nothing
'-- NowInGMT ------------------------------------------------------------------'



End Function



Public Function s3_DeleteLocalFile() ' as string (responsecode)


        Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.Fileexists(s3_strLocalFile) Then objFSO.DeleteFile s3_strLocalFile
        Set objFSO = Nothing




End Function







'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'ררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררר
End Class
'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'ררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררררר

%>