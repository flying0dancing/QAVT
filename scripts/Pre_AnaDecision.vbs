Dim strexcel,DatabaseName,Forms_str,Forms_arr,Formname,logfile,log_tmp,ServerName,User,Password,constring,rownum,sheetname,strQuery,inifile

If WScript.Arguments.length=3 Then
 strexcel=Trim(WScript.Arguments(0))
 logfile=Trim(Wscript.Arguments(1))
 inifile=Trim(Wscript.Arguments(2))
Else
 WScript.Echo "miss argument, Order is full path of excel,log,ini file."
 WScript.Quit(-1)
End If
 ServerName=ReadINI(inifile,"_QAVT_CONFIG_DBSERVER_INSTANCE")
 User=ReadINI(inifile,"_QAVT_CONFIG_DBSERVER_USER")
 Password=ReadINI(inifile,"_QAVT_CONFIG_DBSERVER_PASSWORD")
 DatabaseName=ReadINI(inifile,"_QAVT_CONFIG_DATABASE")
 Forms_str=ReadINI(inifile,"_QAVT_CONFIG_FORM")
 Select Case UCase(ReadINI(inifile,"_QAVT_CONFIG_DBSERVER_TYPE"))
   Case "ORACLE"
   		constring="Provider=MSDAORA.1; Persist Security Info=True; Data Source='"&ServerName&"'; User ID='"&DatabaseName&"'; Password='"&Password&"'"
   Case "SQL"
   		constring="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName&"'; User ID='"&User&"'; Password='"&Password&"'"
   Case Else
   		WScript.Echo "_QAVT_CONFIG_DBSERVER_TYPE should be one of ORACLE,SQL."
   		WScript.Quit
 End Select
log_tmp=Replace(logfile,".log","")
Forms_arr=Split(Forms_str,",",-1,1)
Dim i,j
For i=0 To UBound(Forms_arr)
	Dim tmp
	tmp=Trim(UCase(Forms_arr(i)))
	Forms_arr(i)=tmp
	For j=i+1 To UBound(Forms_arr)
		If tmp=Trim(UCase(Forms_arr(j))) Then
			Forms_arr(j)=""
		End If
	Next
Next

For Each Formname In Forms_arr
	If Formname<>"" Then
	   sheetname=Formname
	   logfile= log_tmp &"_"& sheetname & ".log"
	   strQuery="select c.STBDMPFORM,c.STBDMPITEM,m.* from "&Formname&"_CHILD c , "&Formname&"_MASTER m  where c.STBDMPRECORDNO=m.STBDMPRECORDNO " &_
"order by STBDMPFORM,c.STBDMPITEM,STBDMPCONDITION,STBDMPAMT,STBDMPDRCR,STBDMPGLC1,STBDMPGLC1NOT,STBDMPGLC2,STBDMPGLC2NOT,STBDMPGLC3,STBDMPGLC3NOT,STBDMPGLC4,STBDMPGLC4NOT,STBDMPINST,STBDMPINSTNOT"

	   rownum=RWDataLog(constring,strQuery,strexcel,sheetname,logfile)
	   
	End If
Next


Wscript.Quit(1)

''''''''''''''''''''''''''''''''''
Function RWDataLog(constr,strQuery,excelfl,sheetname,logfl)
On Error Resume Next
'Step1
Dim objConn,objRS,objCmd,i,j,rowcount,colcount,wh
Dim fso,oexcel,openexcel,osheet,varsheet
Dim colname_str,colname_arr
colname_str="STBDMPFORM,STBDMPITEM,STBDMPRECORDNO,STBDMPGLC1NOT,STBDMPGLC1,STBDMPGLC2NOT,STBDMPGLC2,STBDMPGLC3NOT,STBDMPGLC3,STBDMPGLC4NOT,STBDMPGLC4,STBDMPINSTNOT,STBDMPINST,STBDMPDRCR,STBDMPAMT,STBDMPCONDITION"
rowcount=-1
Set objConn=CreateObject("ADODB.Connection")
Set fso=CreateObject("Scripting.FileSystemObject")
Set wh=fso.OpenTextFile(logfl,2,True,TristateUseDefault)
wh.WriteLine("********************* This is log information for LombardRisk QA Decision Table Analysis ****************")
wh.WriteLine("Connect To:"& vbCrLf &constr& vbCrLf)
wh.WriteLine("Execute SQL:"& vbCrLf &strQuery& vbCrLf)


objConn.Open constr

If objConn.State=0 Then
  WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
  wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
  Err.Clear
  objConn.Close
  wh.Close
  Set objConn=Nothing
  Set wh=Nothing
  RWDataLog=-1
  Exit Function
End If  
  'new add
Set objCmd=CreateObject("ADODB.COMMAND")
Set objRS=CreateObject("ADODB.Recordset")
objCmd.ActiveConnection=objConn
objCmd.CommandText=strQuery
objRS.CursorLocation=3
objRS.Open objCmd

 'step2
'  objRS.Open strQuery,objConn,1,1
  If Err.Number<>0 Then
    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
    wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
    Err.Clear
    objCmd.ActiveConnection.Close
    objRS.Close
    wh.Close
    Set objRS=Nothing
    Set objCmd=Nothing
    Set wh=Nothing
    RWDataLog=-1
    Exit Function
  End If
 
    'step3
    rowcount=objRS.RecordCount
    colcount=objRS.Fields.Count
    If colcount=20 Then
     colname_str="STBDMPFORM,STBDMPITEM,STBDMPRECORDNO,STBDMPGLC1NOT,STBDMPGLC1,STBDMPGLC2NOT,STBDMPGLC2,STBDMPGLC3NOT,STBDMPGLC3,STBDMPGLC4NOT,STBDMPGLC4,STBDMPGLC5NOT,STBDMPGLC5,STBDMPGLC6NOT,STBDMPGLC6,STBDMPINSTNOT,STBDMPINST,STBDMPDRCR,STBDMPAMT,STBDMPCONDITION"
    End If
    If colcount=24 Then
     colname_str="STBDMPFORM,STBDMPITEM,STBDMPRECORDNO,STBDMPGLC1NOT,STBDMPGLC1,STBDMPGLC2NOT,STBDMPGLC2,STBDMPGLC3NOT,STBDMPGLC3,STBDMPGLC4NOT,STBDMPGLC4,STBDMPGLC5NOT,STBDMPGLC5,STBDMPGLC6NOT,STBDMPGLC6,STBDMPGLC7NOT,STBDMPGLC7,STBDMPGLC8NOT,STBDMPGLC8,STBDMPINSTNOT,STBDMPINST,STBDMPDRCR,STBDMPAMT,STBDMPCONDITION"
    End If
    If Not (objRS.EOF And objRS.BOF) Then
	'step4
	Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
	oexcel.DisplayAlerts=False
	Set varsheet=CreateObject("Excel.Sheet")
	 If fso.FileExists(excelfl) Then
	
	  Set openexcel=oexcel.Workbooks.Open(excelfl)
	  If Err.Number<>0 Then
	    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	    wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	    Err.Clear
	    oexcel.Quit
	    wh.Close
	    objConn.Close
	    objCmd.ActiveConnection.Close
    	objRS.Close
	    Set oexcel=Nothing
	    Set wh=Nothing
	    Set objConn=Nothing
	    Set objCmd=Nothing
	    Set objRS=Nothing
	    RWDataLog=-1
	    Exit Function
      End If
	  openexcel.Activate
	  oexcel.Visible=False
	    For Each varsheet In openexcel.Worksheets
	      If StrComp(UCase(varsheet.Name),UCase(sheetname),vbTextCompare)=0 Then
	        openexcel.Worksheets(varsheet.Name).Delete
	        Exit For
	      End If
	    Next
	  openexcel.Sheets.Add
	  If Err.Number<>0 Then
	    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	    wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	    Err.Clear
	    oexcel.Quit
	    wh.Close
	    objConn.Close
	    objCmd.ActiveConnection.Close
    	objRS.Close
	    Set oexcel=Nothing
	    Set wh=Nothing
	    Set objConn=Nothing
	    Set objCmd=Nothing
	    Set objRS=Nothing
	    RWDataLog=-1
	    Exit Function
      End If
	  openexcel.ActiveSheet.name=sheetname
	  Set osheet=openexcel.Worksheets(sheetname)
	 Else
	 
	  Set openexcel=oexcel.Workbooks.Add()
	  openexcel.SaveAs(excelfl)
	  openexcel.Activate
	  oexcel.Visible=False
	  openexcel.Sheets.Add
	  openexcel.ActiveSheet.name=sheetname
	  Set osheet=openexcel.Worksheets(sheetname)
	 End If
	 
	'Step5
	  colname_arr=Split(colname_str,",",-1,1) 'new add
	  For j=0 To UBound(colname_arr)-LBound(colname_arr) Step 1 'new add
	  	osheet.Cells(1,j+1).Value=colname_arr(j) 'new add
	  Next 'new add
	  i=1 'new add
	  'i=0 'old 
	  objRS.MoveFirst
	  
	  Do While Not objRS.EOF
	    For j=0 To colcount-1
	     osheet.Cells(i+1,j+1).Value= objRS.Fields(j).Value
	    Next
	     i=i+1
	   objRS.MoveNext
	  Loop
	
	End If



  If Not oexcel.ActiveWorkbook.Saved Then
     openexcel.Save
     wh.WriteLine("Results in: " & vbCrLf& strexcel& vbCrLf)
     wh.Close
  End If
  
  oexcel.Workbooks.Close
  oexcel.Quit
  
  Set wh=Nothing
  Set varsheet=Nothing
  Set fso=Nothing
  Set osheet=Nothing
  Set oexcel=Nothing
  Set openexcel=Nothing
'If objRS.State=adStateOpen Then
objRS.Close
'End If
'If objConn.State=adStateOpen Then
objConn.Close
'End If
Set objCmd=Nothing
set objConn=nothing
Set objRS=Nothing
On Error GoTo 0
'step6
RWDataLog=rowcount
End Function


Function ReadINI(FilePath,PrimaryKey)
Dim fso,INIfile,strline
ReadINI=-1
Set fso=CreateObject("Scripting.FileSystemObject")
Set INIfile=fso.OpenTextFile(FilePath,1)
do until INIfile.AtEndOfStream
  strline=Trim(INIfile.ReadLine)
  If strline<>"" then
   If StrComp(UCase(Trim(Left(strline,instr(strline,"=")-1))),UCase(PrimaryKey))=0 then
       ReadINI = Trim(Right(strline,len(strline)-instr(strline,"="))) '读取等号后的部分
       Exit do
   End If
  End If
loop
INIfile.Close
Set INIfile=Nothing
Set fso=Nothing
End Function