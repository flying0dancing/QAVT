''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  develop by Kun Shen, send email to Kun.Shen@lombardrisk.com if any issue 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strexcel,DatabaseName1,DatabaseName2,Forms_str,Forms_arr,Formname,ServerName,User,Password,constr1,constr2,rowcount1,rowcount2,sheet_new,sheet_old,strQuery
WScript.Echo "Compare Decision script is running..."
WScript.Echo "Start at "&Date&" "& Time
If WScript.Arguments.length=5 Then
 strexcel=Trim(WScript.Arguments(0))
 ServerName=Trim(WScript.Arguments(1))
 DatabaseName1=Trim(WScript.Arguments(2))
 DatabaseName2=Trim(WScript.Arguments(3))
 Forms_str=Trim(WScript.Arguments(4))
 constr1="Provider=SQLOLEDB; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName1&"'; Integrated Security=SSPI;"
 constr2="Provider=SQLOLEDB; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName2&"'; Integrated Security=SSPI;"
ElseIf WScript.Arguments.length=7 Then
 strexcel=Trim(WScript.Arguments(0))
 ServerName=Trim(WScript.Arguments(1))
 DatabaseName1=Trim(WScript.Arguments(2))
 DatabaseName2=Trim(WScript.Arguments(3))
 Forms_str=Trim(WScript.Arguments(4))
 User=Trim(WScript.Arguments(5))
 Password=Trim(WScript.Arguments(6))
 constr1="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName1&"'; User ID='"&User&"'; Password='"&Password&"'"
 constr2="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName2&"'; User ID='"&User&"'; Password='"&Password&"'"
End If
Forms_arr=Split(Forms_str,",",-1,1)
For Each Formname In Forms_arr
'	sheet_old=DatabaseName1 &" "& Formname
'	sheet_new=DatabaseName2 &" "& Formname
	sheet_old="old "& Formname
	sheet_new="new "& Formname
	strQuery="select c.STBDMPFORM,c.STBDMPITEM,m.* from dbo."&Formname&"_CHILD c , dbo."&Formname&"_MASTER m where c.STBDMPRECORDNO=m.STBDMPRECORDNO " &_
"order by c.STBDMPRECORDNO"

	rowcount1=RWData(constr1,strQuery,strexcel,sheet_old)
	'WScript.Echo "rowcount1: "&rowcount1
	WScript.Sleep(100)
	rowcount2=RWData(constr2,strQuery,strexcel,sheet_new)
	'WScript.Echo "rowcount2: "&rowcount2
	rowcount1=rowcount1+1
	rowcount2=rowcount2+1
	If rowcount1>1 And rowcount2>1 Then
		Call ComparedSheets
		WScript.Echo "Compared Results at " & strexcel
	Else
 	   WScript.Echo "Warning: retrieve records from "&Formname&" Decision Table in ["&DatabaseName1&"] or ["&DatabaseName2&"] are empty."
	End If 
Next

WScript.Echo "End at "&Date&" "& Time
Wscript.Quit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ComparedSheets
On Error Resume Next
Dim oexcel,openexcel,osheet1,osheet2,osheet3,varsheet,shtname
shtname=Formname&" Compared Log"

Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
Set varsheet=CreateObject("Excel.Sheet")
oexcel.DisplayAlerts=False

Set openexcel=oexcel.Workbooks.Open(strexcel)
openexcel.Activate
Set osheet1=openexcel.Worksheets(sheet_old)
Set osheet2=openexcel.Worksheets(sheet_new)
   For Each varsheet In openexcel.Worksheets
      If StrComp(varsheet.Name,shtname,vbTextCompare)=0 Then
       openexcel.Worksheets(shtname).Delete
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
osheet3.Range("A1:D1").Font.Size=21
osheet3.Range("A1:D1").Font.Color=RGB(100,100,255)
osheet3.Range("A1:D1").Font.Bold=True
osheet3.Cells(1,1).Value=Formname&" Decision Compared Results"
osheet3.Range("A1:D1").Merge
osheet3.Range("A1:D1").HorizontalAlignment=2
osheet3.Cells(2,1).Value="STBDMPRECORDNO"
osheet3.Cells(2,2).Value="Deleted in "&DatabaseName2
osheet3.Cells(2,3).Value="New Add in "&DatabaseName2
osheet3.Cells(2,4).Value="Modified in "&DatabaseName2

sheet1rows=osheet1.UsedRange.Rows.Count
sheet2rows=osheet2.UsedRange.Rows.Count
sheet3rows=osheet3.UsedRange.Rows.Count
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
    For col=4 To 16 Step 1
      cmpStrs=StrComp(osheet1.Cells(r1,col).Value,osheet2.Cells(r2,col).Value, vbTextCompare)
      
      If cmpStrs<>0 Then
        osheet1.Rows(r1).Interior.ColorIndex=color3
        osheet2.Rows(r2).Interior.ColorIndex=color3
        osheet3.Cells(r3,1).Value=str1 & " is modified. "&osheet1.name&".cell(" & r1 &"," & col & ")=" & osheet1.Cells(r1,col).Value &" ==>> "&osheet2.name&".cell(" & r2 &"," & col & ")=" & osheet2.Cells(r2,col).Value
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
On Error GoTo 0
End Sub

''''''''''''''''''''''''''''''''''
'================================================
''''''''''''''''''''''''''''''''''
Function RWData(constr,strQuery,excelfl,sheetname)
On Error Resume Next
'Step1
Dim objConn,objRS,i,j,rowcount,colcount
Dim fso,oexcel,openexcel,osheet,varsheet
Dim colname_str,colname_arr
colname_str="STBDMPFORM,STBDMPITEM,STBDMPRECORDNO,STBDMPGLC1NOT,STBDMPGLC1,STBDMPGLC2NOT,STBDMPGLC2,STBDMPGLC3NOT,STBDMPGLC3,STBDMPGLC4NOT,STBDMPGLC4,STBDMPINSTNOT,STBDMPINST,STBDMPDRCR,STBDMPAMT,STBDMPCONDITION"
rowcount=-1
Set objConn=CreateObject("ADODB.Connection")
Set objRS=CreateObject("ADODB.Recordset")
Set fso=CreateObject("Scripting.FileSystemObject")

'strQuery="select c.STBDMPFORM,c.STBDMPITEM,m.* from dbo."&tbl&"_CHILD c , dbo."&tbl&"_MASTER m where c.STBDMPRECORDNO=m.STBDMPRECORDNO " &_
'"order by c.STBDMPRECORDNO"

objConn.Open constr
If Err.Number<>0 Then
  WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
  Err.Clear
  objConn.Close
  Set objConn=Nothing
  RWData=-1
  Exit Function
End If
  'step2
  objRS.Open strQuery,objConn,1,1
  If Err.Number<>0 Then
    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
    Err.Clear
    objRS.Close
    Set objRS=Nothing
    RWData=-1
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
	    Err.Clear
	    oexcel.Quit
	    Set oexcel=Nothing
	    RWData=-1
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
	  openexcel.ActiveSheet.name=sheetname
	  Set osheet=openexcel.Worksheets(sheetname)
	 Else
	 
	  Set openexcel=oexcel.Workbooks.Add()
	  If Err.Number<>0 Then
	    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
	    Err.Clear
	    oexcel.Quit
	    Set oexcel=Nothing
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
	  For j=0 To UBound(colname_arr)-LBound(colname_arr)
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
'End If
'If objConn.State=adStateOpen Then
objConn.Close
'End If
set objConn=nothing
Set objRS=Nothing
On Error GoTo 0
'step6
RWData=rowcount
End Function

