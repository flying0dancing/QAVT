''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  develop by Kun Shen, send email to Kun.Shen@lombardrisk.com if any issue 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strexcel,DatabaseName1,DatabaseName2,Forms_str,Forms_arr,Formname,constr1,constr2,rowcount1,rowcount2,sheet_new,sheet_old,strQuery,inifile
'Dim ServerName1,User1,Password1,ServerName2,User2,Password2

If WScript.Arguments.length=2 Then
WScript.Echo "Compare Decision script is running..."
WScript.Echo "Start at "&Date&" "& Time
 strexcel=Trim(WScript.Arguments(0))
 inifile=Trim(Wscript.Arguments(1))
Else
 WScript.Echo "miss argument, argument order: excel file,ini."
 WScript.Quit
End If
 Forms_str=ReadINI(inifile,"common","_QAVT_CONFIG_FORM")
' ServerName1=ReadINI(inifile,"DB","_QAVT_CONFIG_DBSERVER_INSTANCE")
' User1=ReadINI(inifile,"DB","_QAVT_CONFIG_DBSERVER_USER")
' Password1=ReadINI(inifile,"DB","_QAVT_CONFIG_DBSERVER_PASSWORD")
' ServerName2=ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DBSERVER_INSTANCE")
' User2=ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DBSERVER_USER")
' Password2=ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DBSERVER_PASSWORD")
 DatabaseName1=ReadINI(inifile,"DB","_QAVT_CONFIG_DATABASE")
 DatabaseName2=ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DATABASE")
constr1=SetConnectStr(ReadINI(inifile,"DB","_QAVT_CONFIG_DBSERVER_TYPE"),ReadINI(inifile,"DB","_QAVT_CONFIG_DBSERVER_INSTANCE"),ReadINI(inifile,"DB","_QAVT_CONFIG_DBSERVER_USER"),ReadINI(inifile,"DB","_QAVT_CONFIG_DBSERVER_PASSWORD"),DatabaseName1)
constr2=SetConnectStr(ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DBSERVER_TYPE"),ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DBSERVER_INSTANCE"),ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DBSERVER_USER"),ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DBSERVER_PASSWORD"),DatabaseName2)
'constr1=SetConnectStr(ReadINI(inifile,"DB","_QAVT_CONFIG_DBSERVER_TYPE"),ServerName1,User1,Password1,DatabaseName1)
'constr2=SetConnectStr(ReadINI(inifile,"CMPDB","_QAVT_CONFIG_DBSERVER_TYPE"),ServerName2,User2,Password2,DatabaseName2)
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
	 sheet_old="old "& Formname
	 sheet_new="new "& Formname
	 strQuery="select c.STBDMPFORM,c.STBDMPITEM,m.* from "&Formname&"_CHILD c , "&Formname&"_MASTER m where c.STBDMPRECORDNO=m.STBDMPRECORDNO " &_
"order by c.STBDMPRECORDNO"

	 rowcount1=RWData(constr1,strQuery,strexcel,sheet_new)
'	 WScript.Echo "rowcount1: "&rowcount1
	 'WScript.Sleep(100)
	 rowcount2=RWData(constr2,strQuery,strexcel,sheet_old)
'	 WScript.Echo "rowcount2: "&rowcount2
'	 rowcount1=rowcount1+1
'	 rowcount2=rowcount2+1
	 If rowcount1>0 And rowcount2>0 Then
		 Call ComparedSheets(DatabaseName1,DatabaseName2)
		 WScript.Echo Formname& " compared Results at " & strexcel
	 Else
' 	     WScript.Echo "Warning: retrieve records from "&Formname&" Decision Table in ["&DatabaseName1&"] or ["&DatabaseName2&"] are empty."
		 WScript.Echo "Warning: no need compare decision tables in Form ["&Formname&"],"&vbCrLf&" one of these DataBases doesn't have records."
	 End If 
	End If
Next

WScript.Echo "End at "&Date&" "& Time
Wscript.Quit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ComparedSheets(newDB,oldDB)
On Error Resume Next
Dim oexcel,openexcel,osheet1,osheet2,osheet3,varsheet,shtname
shtname=Formname&" Compared Log"

Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
Set varsheet=CreateObject("Excel.Sheet")
oexcel.DisplayAlerts=False

Set openexcel=oexcel.Workbooks.Open(strexcel)
openexcel.Activate
Set osheet1=openexcel.Worksheets(sheet_old)'opposite from main func
Set osheet2=openexcel.Worksheets(sheet_new)'opposite from man func
   For Each varsheet In openexcel.Worksheets
      If StrComp(UCase(varsheet.Name),UCase(shtname),vbTextCompare)=0 Then
       openexcel.Worksheets(varsheet.Name).Delete
       Exit For
      End If
   Next
openexcel.Sheets.Add
openexcel.ActiveSheet.name=shtname
Set osheet3=openexcel.Worksheets(shtname)
oexcel.Visible=false

Dim r1,r2,r3,str1,str2,cmpStrs,color1,color2,color3,sheet1rows,sheet2rows,sheet3rows
color1=28
color2=22
color3=50
osheet3.Range("A1:D1").Font.Size=16
osheet3.Range("A1:D1").Font.Color=RGB(100,100,255)
osheet3.Range("A1:D1").Font.Bold=True
osheet3.Cells(1,1).Value=Formname&" Decision Compared Results"
osheet3.Range("A1:D1").Merge
osheet3.Range("A1:D1").HorizontalAlignment=2
osheet3.Cells(2,1).Value="old database: "&oldDB&"; new database: "&newDB
osheet3.Range("A2:D2").Merge
osheet3.Cells(3,1).Value=osheet1.Cells(1,3).Value '="STBDMPRECORDNO"
osheet3.Cells(3,2).Value="Deleted in "&newDB
osheet3.Cells(3,3).Value="New Add in "&newDB
osheet3.Cells(3,4).Value="Modified in "&newDB
osheet3.Columns.AutoFit()

sheet1rows=osheet1.UsedRange.Rows.Count
sheet2rows=osheet2.UsedRange.Rows.Count
sheet3rows=osheet3.UsedRange.Rows.Count
sheet1cols=osheet1.UsedRange.Columns.Count
r1=2
r2=2
r3=sheet3rows+1
While r1<=sheet1rows And r2<=sheet2rows
  str1=CInt(Trim(osheet1.Cells(r1,3).Value))
  str2=CInt(Trim(osheet2.Cells(r2,3).Value))

  If str1<str2 Then
'    WScript.Echo str1& "<" & str2&"::::"&r1 & "vs" &r2
    osheet1.Rows(r1).Interior.ColorIndex=color1
'    osheet3.Cells(r3,1).Value=str1 & " is deleted in """ & osheet2.name&"""."
    osheet3.Cells(r3,1).Value=str1
    osheet3.Cells(r3,2).Value="Y"
    r3=r3+1
    r1=r1+1

  ElseIf str1>str2 Then
'    WScript.Echo str1& ">" & str2&"::::"&r1 & "vs" &r2
    osheet2.Rows(r2).Interior.ColorIndex=color2
'    osheet3.Cells(r3,1).Value=str2 & " is new added in """ & osheet2.name&"""."
    osheet3.Cells(r3,1).Value=str2
    osheet3.Cells(r3,3).Value="Y"
    r3=r3+1
    r2=r2+1

  Else
'    WScript.Echo str1& "=" & str2&"::::"&r1 & "vs" &r2
    Dim col
    For col=1 To sheet1cols Step 1
      cmpStrs=StrComp(osheet1.Cells(r1,col).Value,osheet2.Cells(r2,col).Value, vbTextCompare)
      
      If cmpStrs<>0 Then
        osheet1.Rows(r1).Interior.ColorIndex=color3
        osheet2.Rows(r2).Interior.ColorIndex=color3
        osheet3.Cells(r3,1).Value=str1 & " changed. "&osheet1.Cells(1,col).Value&": "& osheet1.Cells(r1,col).Value &" ==>> "& osheet2.Cells(r2,col).Value
        osheet3.Cells(r3,4).Value="Y"
        r3=r3+1
      End If
      
    Next
    r1=r1+1
    r2=r2+1
  End If
  
Wend

While r1<=sheet1rows
    osheet1.Rows(r1).Interior.ColorIndex=color1
'    osheet3.Cells(r3,1).Value=osheet1.Cells(r1,3) & " is deleted in """ & osheet2.name&"""."
    osheet3.Cells(r3,1).Value=Trim(osheet1.Cells(r1,3).Value)
    osheet3.Cells(r3,2).Value="Y"    
    r3=r3+1
    r1=r1+1
Wend

While r2<=sheet2rows
    osheet2.Rows(r2).Interior.ColorIndex=color2
'    osheet3.Cells(r3,1).Value=osheet1.Cells(r2,3)& " is new added in """ & osheet2.name&"""."
    osheet3.Cells(r3,1).Value=Trim(osheet2.Cells(r2,3).Value)
    osheet3.Cells(r3,3).Value="Y"
    r3=r3+1
    r2=r2+1
Wend

openexcel.Save
'oexcel.Visible=true
oexcel.Workbooks.Close
oexcel.Quit

Set varsheet=Nothing
Set osheet1=Nothing
Set osheet2=Nothing
Set osheet3=Nothing
Set openexcel=Nothing
Set oexcel=Nothing
ComparedSheets=1
On Error GoTo 0
End Function

''''''''''''''''''''''''''''''''''
'================================================
''''''''''''''''''''''''''''''''''
Function RWData(constr,strQuery,excelfl,sheetname)
On Error Resume Next
'Step1
Dim objConn,objRS,objCmd,i,j,rowcount,colcount
Dim fso,oexcel,openexcel,osheet,varsheet
Dim colname_str,colname_arr
colname_str="STBDMPFORM,STBDMPITEM,STBDMPRECORDNO,STBDMPGLC1NOT,STBDMPGLC1,STBDMPGLC2NOT,STBDMPGLC2,STBDMPGLC3NOT,STBDMPGLC3,STBDMPGLC4NOT,STBDMPGLC4,STBDMPINSTNOT,STBDMPINST,STBDMPDRCR,STBDMPAMT,STBDMPCONDITION"
rowcount=-2
Set objConn=CreateObject("ADODB.Connection")
Set fso=CreateObject("Scripting.FileSystemObject")


objConn.Open constr
If objConn.State=0 Then
  WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
  Err.Clear
  objConn.Close
  Set objConn=Nothing
  RWData=-1
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
    Err.Clear
    objCmd.ActiveConnection.Close
    objRS.Close
    Set objRS=Nothing
    Set objCmd=Nothing
    RWData=-1
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
	    Err.Clear
	    oexcel.Quit
	    objConn.Close
	    objCmd.ActiveConnection.Close
    	objRS.Close
	    Set oexcel=Nothing
	    Set objConn=Nothing
	    Set objCmd=Nothing
	    Set objRS=Nothing
	    RWData=-1
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
	  openexcel.ActiveSheet.name=sheetname
	  Set osheet=openexcel.Worksheets(sheetname)
	 Else
	 
	  Set openexcel=oexcel.Workbooks.Add()
	  If Err.Number<>0 Then
	    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	    Err.Clear
	    oexcel.Quit
	    objConn.Close
	    objCmd.ActiveConnection.Close
    	objRS.Close
	    Set oexcel=Nothing
	    Set objConn=Nothing
	    Set objCmd=Nothing
	    Set objRS=Nothing
	    RWData=-1
	    Exit Function
      End If
	  openexcel.SaveAs(excelfl)
	  openexcel.Activate
	  oexcel.Visible=False
	  openexcel.Sheets.Add
	  openexcel.ActiveSheet.name=sheetname
	  Set osheet=openexcel.Worksheets(sheetname)
	 End If
	 
	'Step5
	 
	  colname_arr=Split(colname_str,",",-1,1)
	  For j=0 To UBound(colname_arr)-LBound(colname_arr) Step 1
	  	osheet.Cells(1,j+1).Value=colname_arr(j)
	  Next
	  i=1
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
  End If
  
  oexcel.Workbooks.Close
  oexcel.Quit
  
  Set varsheet=Nothing
  Set fso=Nothing
  Set osheet=Nothing
  Set oexcel=Nothing
  Set openexcel=Nothing
'If objRS.State=adStateOpen Then
objRS.Close

objConn.Close
Set objCmd=Nothing
set objConn=nothing
Set objRS=Nothing
On Error GoTo 0
'step6
RWData=rowcount
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ReadINI(FilePath,Bar,PrimaryKey)
Dim fso,strline,INIfile
ReadINI=-1
Set fso=CreateObject("Scripting.FileSystemObject")
Set INIfile=fso.opentextfile(FilePath,1)

Do Until INIfile.AtEndOfStream
	strline=Trim(INIfile.ReadLine)
	If strline<>"" And UCase(strline)="["&UCase(Bar)&"]" Then
		do until INIfile.AtEndOfStream
  			strline=Trim(INIfile.ReadLine)
  			If instr(strline,"[")>0 Then
  				Exit Do
  			End If
  			If strline<>"" Then
   				If StrComp(UCase(Trim(Left(strline,instr(strline,"=")-1))),UCase(PrimaryKey))=0 Then
       				ReadINI=Trim(Right(strline,len(strline)-instr(strline,"="))) '读取等号后的部分
       				Exit Do
   				End If
  			End If
		Loop
		Exit Do
	End If
Loop

INIfile.Close
Set INIfile=Nothing
Set fso=nothing
End Function

''''''''''''''''''''''''
Function SetConnectStr(ServerType,ServerName,User,Password,DatabaseName)
Dim constring
 Select Case UCase(Trim(ServerType))
   Case "ORACLE"
   		constring="Provider=MSDAORA.1; Persist Security Info=True; Data Source='"&ServerName&"'; User ID='"&DatabaseName&"'; Password='"&Password&"'"
   Case "SQL"
   		constring="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName&"'; User ID='"&User&"'; Password='"&Password&"'"
   Case Else
   		constring=-1
   		
 End Select
 SetConnectStr=constring
End Function