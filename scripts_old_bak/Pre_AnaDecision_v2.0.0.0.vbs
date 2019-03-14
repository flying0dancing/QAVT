Dim strexcel,DatabaseName,table,logfile,ServerName,User,Password,constring,rownum,sheetname,strQuery
If WScript.Arguments.length=5 Then
 strexcel=WScript.Arguments(0)
 ServerName=WScript.Arguments(1)
 DatabaseName=WScript.Arguments(2)
 table=WScript.Arguments(3)
 logfile=Wscript.Arguments(4)
 constring="Provider=SQLOLEDB; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName&"'; Integrated Security=SSPI;"
ElseIf WScript.Arguments.length=7 Then
 strexcel=WScript.Arguments(0)
 ServerName=WScript.Arguments(1)
 DatabaseName=WScript.Arguments(2)
 table=WScript.Arguments(3)
 logfile=Wscript.Arguments(4)
 User=WScript.Arguments(5)
 Password=WScript.Arguments(6)
 constring="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName&"'; User ID='"&User&"'; Password='"&Password&"'"
End If
sheetname=table
strQuery="select c.STBDMPFORM,c.STBDMPITEM,m.* from dbo."&table&"_CHILD as c , dbo."&table&"_MASTER as m where c.STBDMPRECORDNO=m.STBDMPRECORDNO " &_
"order by STBDMPFORM,c.STBDMPITEM,STBDMPCONDITION,STBDMPAMT,STBDMPDRCR,STBDMPGLC1,STBDMPGLC1NOT,STBDMPGLC2,STBDMPGLC2NOT,STBDMPGLC3,STBDMPGLC3NOT,STBDMPGLC4,STBDMPGLC4NOT,STBDMPINST,STBDMPINSTNOT"

rownum=RWDataLog(constring,strQuery,strexcel,sheetname,logfile)

Wscript.Quit(rownum)

''''''''''''''''''''''''''''''''''
Function RWDataLog(constr,strQuery,excelfl,sheetname,logfl)
On Error Resume Next
'Step1
Dim objConn,objRS,i,j,rowcount,colcount,wh
Dim fso,oexcel,openexcel,osheet,varsheet
rowcount=-1
Set objConn=CreateObject("ADODB.Connection")
Set objRS=CreateObject("ADODB.Recordset")
Set fso=CreateObject("Scripting.FileSystemObject")
Set wh=fso.OpenTextFile(logfl,2,True,TristateUseDefault)
wh.WriteLine("********************* This is log information for LombardRisk QA Decision Table Analysis ****************")
wh.WriteLine("Execute SQL:"& vbCrLf &strQuery& vbCrLf)


objConn.Open constr
If Err.Number<>0 Then
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
  'step2
  objRS.Open strQuery,objConn,1,1
  If Err.Number<>0 Then
    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
    wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
    Err.Clear
    objRS.Close
    wh.Close
    Set objRS=Nothing
    Set wh=Nothing
    RWDataLog=-1
    Exit Function
  End If
    'step3
    rowcount=objRS.RecordCount
    colcount=objRS.Fields.Count
    
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
	    Set oexcel=Nothing
	    Set wh=Nothing
	    RWDataLog=-1
	    Exit Function
      End If
	  openexcel.Activate
	  oexcel.Visible=False
	    For Each varsheet In openexcel.Worksheets
	      If StrComp(varsheet.Name,sheetname,vbTextCompare)=0 Then
	        openexcel.Worksheets(sheetname).Delete
	        Exit For
	      End If
	    Next
	  openexcel.Sheets.Add
	  If Err.Number<>0 Then
	    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	    wh.WriteLine "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	    Err.Clear
	    oexcel.Quit
	    Set oexcel=Nothing
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
	  i=0
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
set objConn=nothing
Set objRS=Nothing
On Error GoTo 0
'step6
RWDataLog=rowcount
End Function

