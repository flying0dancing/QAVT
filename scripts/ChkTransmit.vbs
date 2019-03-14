''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  develop by Kun Shen, send email to Kun.Shen@lombardrisk.com if any issue 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
Dim strexl,strexcel,shtname,DatabaseName,Tablename,ServerName,User,Password,constr,rowcount,strQuery,logfl,fsobjt,whandle,trnexcelname

If WScript.Arguments.length=8 Then
	strexl=WScript.Arguments(0)
	shtname=WScript.Arguments(1)
	ServerName=WScript.Arguments(2)
	DatabaseName=WScript.Arguments(3)
	Tablename=WScript.Arguments(4)
	trnexcelname=WScript.Arguments(5)
	strexcel=WScript.Arguments(6)
	logfl=WScript.Arguments(7)
	constr="Provider=SQLOLEDB; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName&"'; Integrated Security=SSPI;"
ElseIf WScript.Arguments.length=10 Then
	strexl=WScript.Arguments(0)
	shtname=WScript.Arguments(1)
	ServerName=WScript.Arguments(2)
	DatabaseName=WScript.Arguments(3)
	Tablename=WScript.Arguments(4)
	trnexcelname=WScript.Arguments(5)
	strexcel=WScript.Arguments(6)
	logfl=WScript.Arguments(7)
	User=WScript.Arguments(8)
	Password=WScript.Arguments(9)
	constr="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName&"'; User ID='"&User&"'; Password='"&Password&"'"
Else
	WScript.Echo "Arguments are not correct. Terminated scripts..."
	WScript.Echo "Argument List:DFM template excel,sheet name,DB server,Database,table,transmission results,output excel(fullpath),output log(fullpath),user,password"
	WScript.Quit
End If

'strQuery="SELECT [CELLITEM],[CELLVALUE] FROM "&Tablename
strQuery="SELECT distinct a.STBITEM ,b.S_FormValue,a.S_FormAlphaValue FROM "&Tablename&" as a INNER JOIN (select STBITEM,SUM(S_FormValue)as S_FormValue FROM "&Tablename&" where STBSTATUS='A' GROUP BY STBITEM) as b on a.STBITEM=b.STBITEM"
WScript.Echo strexcel
WScript.Echo trnexcelname
Set fsobjt=CreateObject("Scripting.FileSystemObject")
Set whandle=fsobjt.CreateTextFile(logfl,True,TristateUseDefault)
whandle.WriteLine("This information for QAVT Do Transmission Check")
WScript.Echo "Start at "&Date&" "& Time
whandle.WriteLine("Start at "&Date&" "& Time)
whandle.WriteBlankLines(1)
whandle.WriteLine("Connect to:"&vbTab&constr)
whandle.WriteLine("SQL QUERY:"&vbTab&strQuery)
whandle.WriteLine("DFM Template @"&vbTab&strexl)
whandle.WriteLine("Transmission Results @"&vbTab&trnexcelname)
whandle.WriteLine("Check Transmission Output @"&vbTab&strexcel)
whandle.WriteLine("More Details Log:")
whandle.Close
Set whandle=Nothing
If fsobjt.FileExists(strexl) And fsobjt.FileExists(trnexcelname) Then
	fsobjt.CopyFile strexl,strexcel
	
'	call WriteDataToExl(strexcel,shtname,constr,strQuery,logfl)
	If WriteDataToExl(strexcel,shtname,constr,strQuery,logfl)=0 Then
	Call CompTransmission(strexcel,shtname,trnexcelname,logfl)
'WScript.Echo "Call CompTransmission(strexcel,shtname,trnexcelname,logfl)"
	End If
	
Else
	WScript.Echo strexl&" doesn't exist or transmission results "&trnexcelname&"doesn't exist."
End If

WScript.Echo "End at "&Date&" "& Time
Set fsobjt=Nothing
On Error Goto 0
WScript.Quit


Sub CompTransmission(excelname,sheetname,strexlname,logfile)
On Error Resume Next
Dim oexcel,openexcel,osheet,objexcel,transexcel,transheet
Dim varsheet,shtrows,shtcols,i,j
Dim fso,wh
Set fso=CreateObject("Scripting.FileSystemObject")
Set wh=fso.OpenTextFile(logfile,8,TristateUseDefault)
Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
oexcel.DisplayAlerts=False
Set varsheet=CreateObject("Excel.Sheet")
Set openexcel=oexcel.Workbooks.Open(excelname)
If Err.Number<>0 Then
	WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	wh.WriteLine("Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source)
	Err.Clear
	oexcel.Workbooks.Close
	oexcel.Quit
	Set oexcel=Nothing
	wh.Close
	Set wh=Nothing
	Set fso=Nothing
	Exit Sub
End If
openexcel.Activate
oexcel.Visible=true
For Each varsheet In openexcel.Worksheets
	If StrComp(varsheet.Name,sheetname,vbTextCompare)=0 Then
		Set osheet=openexcel.Worksheets(sheetname)
		Exit For
	End If
Next
Set varsheet=Nothing
shtrows=osheet.UsedRange.Rows.Count
shtcols=osheet.UsedRange.Columns.Count

Set objexcel=CreateObject("Excel.Application")
Set transexcel=objexcel.Workbooks.Open(strexlname)
If Err.Number<>0 Then
	WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	wh.WriteLine("Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source)
	Err.Clear
	objexcel.Workbooks.Close
	objexcel.Quit
	Set objexcel=Nothing
	wh.Close
	Set wh=Nothing
	Set fso=Nothing
	Exit Sub
End If
objexcel.Visible=True
Set varsheet=CreateObject("Excel.Sheet")
For Each varsheet In transexcel.Worksheets
	If StrComp(varsheet.Name,sheetname,vbTextCompare)=0 Then
		Set transheet=transexcel.Worksheets(sheetname)
		Exit For
	End If
Next
Set varsheet=Nothing
For i=2 To shtrows
	For j=1 To shtcols
		If StrComp(osheet.Cells(i,j).Value,transheet.Cells(i,j).Value,vbTextCompare)<>0 Or osheet.Cells(i,j).Font.Color <> transheet.Cells(i,j).Font.Color or osheet.Cells(i,j).Font.Size <> transheet.Cells(i,j).Font.Size Or osheet.Cells(i,j).Interior.ColorIndex <> transheet.Cells(i,j).Interior.ColorIndex Then
'			WScript.Echo "Warning: "&"cell("&i&","&j&") of "&sheetname&" is "&transheet.Cells(i,j).Value&",expected is "&osheet.Cells(i,j).Value
'			wh.WriteLine("Warning: "&"cell("&i&","&j&") of "&sheetname&" is "&transheet.Cells(i,j).Value&",expected is "&osheet.Cells(i,j).Value)
			wh.WriteLine("Warning: "&"cell("&i&","&j&") of "&sheetname&" is unexpected.")
		End If

	Next
Next

openexcel.Saved=True
oexcel.Workbooks.Close
oexcel.Quit
Set osheet=Nothing
Set oexcel=Nothing
Set openexcel=Nothing
transexcel.Saved=True
objexcel.Workbooks.Close
objexcel.Quit
Set transheet=Nothing
Set transexcel=Nothing
Set objexcel=Nothing
wh.Close
Set wh=Nothing
Set fso=Nothing
On Error Goto 0
End Sub

''''''''''''''''''''''''''''''''''
'wirte data from database to excel template
Function WriteDataToExl(excelname,sheetname,constring,sqlquery,logfile)
On Error Resume Next
Dim objConn,objRS,i,j,shtrows,shtcols,flag,tempstr,alphavalue,numericvalue
Dim oexcel,openexcel,osheet,varsheet
Dim fso,wh
Set objConn=CreateObject("ADODB.Connection")
Set fso=CreateObject("Scripting.FileSystemObject")
Set wh=fso.OpenTextFile(logfile,8,TristateUseDefault)

objConn.Open constring
If Err.Number<>0 Then
	WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	Err.Clear
	objConn.Close
	Set objConn=Nothing
	wh.Close
	Set wh=Nothing
	Set fso=Nothing
	WriteDataToExl=-1
	Exit Function
'	Exit Sub
End If
Set objRS=CreateObject("ADODB.Recordset")
objRS.Open sqlquery,objConn,1,1
If Err.Number<>0 Then
	WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	Err.Clear
	objRS.Close
	Set objRS=Nothing
	wh.Close
	Set wh=Nothing
	Set fso=Nothing
	WriteDataToExl=-1
	Exit Function
'	Exit Sub
End If
    
If Not (objRS.EOF And objRS.BOF) Then

	Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
	oexcel.DisplayAlerts=False
	Set varsheet=CreateObject("Excel.Sheet")
	If fso.FileExists(excelname) Then
	
		Set openexcel=oexcel.Workbooks.Open(excelname)
		If Err.Number<>0 Then
			WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
			wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
			Err.Clear
			oexcel.Workbooks.Close
			oexcel.Quit
			wh.Close
			Set wh=Nothing
			Set fso=Nothing
			Set oexcel=Nothing
			WriteDataToExl=-1
			Exit Function
'			Exit Sub
		End If
		openexcel.Activate
		oexcel.Visible=True
		For Each varsheet In openexcel.Worksheets
			If StrComp(varsheet.Name,sheetname,vbTextCompare)=0 Then
				Set osheet=openexcel.Worksheets(sheetname)
				Exit For
			End If
		Next
		Set varsheet=Nothing
		shtrows=osheet.UsedRange.Rows.Count
		shtcols=osheet.UsedRange.Columns.Count
		wh.WriteLine "Search Range: (1,1)-("&shtrows&","&shtcols&") in "&sheetname
	End If

objRS.MoveFirst
Do While Not objRS.EOF
	 tempstr= Trim(objRS.Fields(0).Value)
	 alphavalue=Trim(objRS.Fields(2).Value)
	 numericvalue=Trim(objRS.Fields(1).Value)
	 wh.WriteLine "search......"&tempstr '&",alpha:"&alphavalue&",value:"&numericvalue
	 For  j=1 To shtcols
		 For  i=1 To shtrows
			Dim tmpcell
			tmpcell=osheet.Cells(i,j).Value&""
			flag=0
			If Trim(tmpcell)<>"" Then

				If RegExpTest(tmpcell,tempstr) Then
					If IsNull(alphavalue) or alphavalue="" Then
						osheet.Cells(i,j).Value=numericvalue
						wh.WriteLine "Replaced :"&tmpcell&" with "&osheet.Cells(i,j).Value&"."
						flag=1
						Exit For

					Else
						osheet.Cells(i,j).Value=RegExpReplace(tmpcell,tempstr,alphavalue)
						wh.WriteLine "Replaced :"&tmpcell&" with "&osheet.Cells(i,j).Value&"."
						openexcel.Save
						flag=1
						Exit For
					End If
					
				End If
			End If				 
		 Next
		 If flag=1 Then
			 Exit For
		 End If
	 Next
	 If flag=0 Then
'		 WScript.Echo "Warning: cannot find "&tempstr&" in sheet of "&sheetname
		 wh.WriteLine "Warning: cannot find "&tempstr&" in "&sheetname
	 End If
	objRS.MoveNext
Loop
openexcel.Save
Else
	objConn.Close
	Set objConn=Nothing
	wh.Close
	Set wh=Nothing
	Set fso=Nothing
	objRS.Close
	Set objRS=Nothing
	WriteDataToExl=-1
	Exit Function
'	Exit Sub
End If

If Not oexcel.ActiveWorkbook.Saved Then
	openexcel.Save
End If

oexcel.Workbooks.Close
oexcel.Quit

Set osheet=Nothing
Set oexcel=Nothing
Set openexcel=Nothing

objRS.Close
objConn.Close
set objConn=nothing
Set objRS=Nothing

wh.Close
Set wh=Nothing
Set fso=Nothing
On Error Goto 0
WriteDataToExl=0

End Function
'End Sub


Function RegExpReplace(str1,patrn,replStr)
  Dim regEx             ' 建立变量。
  Set regEx = New RegExp               ' 建立正则表达式。
  regEx.Pattern = patrn               ' 设置模式。
  regEx.IgnoreCase = True         ' 设置不区分大小写。
  regEx.Global = True
  RegExpReplace= regEx.Replace(str1,replStr)         ' 作替换。

End Function

Function RegExpTest(strng,patrn)
  Dim regEx, retVal            ' 建立变量。
  Set regEx = New RegExp         ' 建立正则表达式。
  regEx.Pattern = patrn         ' 设置模式。
  regEx.IgnoreCase =True         ' 设置不区分大小写。
  regEx.Global = True
  retVal = regEx.Test(strng)         ' 执行搜索测试。
  RegExpTest = retVal

End Function


