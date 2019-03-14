''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  develop by Kun Shen, send email to Kun.Shen@lombardrisk.com if any issue 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strexcel,DatabaseName1,DatabaseName2,Formname,ServerName,User,Password,constr1,constr2,rowcount1,rowcount2,sheet_new,sheet_old,strQuery,shtname

WScript.Echo "Start at "&Date&" "& Time
If WScript.Arguments.length=5 Then
 strexcel=WScript.Arguments(0)
 ServerName=WScript.Arguments(1)
 'old one
 DatabaseName1=WScript.Arguments(2)
 'new one
 DatabaseName2=WScript.Arguments(3)
 Formname=WScript.Arguments(4)
 constr1="Provider=SQLOLEDB; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName1&"'; Integrated Security=SSPI;"
 constr2="Provider=SQLOLEDB; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName2&"'; Integrated Security=SSPI;"
ElseIf WScript.Arguments.length=7 Then
 strexcel=WScript.Arguments(0)
 ServerName=WScript.Arguments(1)
 DatabaseName1=WScript.Arguments(2)
 DatabaseName2=WScript.Arguments(3)
 Formname=WScript.Arguments(4)
 User=WScript.Arguments(5)
 Password=WScript.Arguments(6)
 constr1="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName1&"'; User ID='"&User&"'; Password='"&Password&"'"
 constr2="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source='"&ServerName&"'; Initial Catalog='"&DatabaseName2&"'; User ID='"&User&"'; Password='"&Password&"'"
End If
strQuery="SELECT [ReturnId],[ExpId],[ExpOrder],[DestFld],[Expression] FROM "&Formname
sheet_old=DatabaseName1 &" "& Formname
sheet_new=DatabaseName2 &" "& Formname
shtname=Formname&" Compared Log"
On Error Resume Next
rowcount1=RWData(constr1,strQuery,strexcel,sheet_old)
WScript.Sleep(100)
rowcount2=RWData(constr2,strQuery,strexcel,sheet_new)
'WScript.Echo rowcount1 & rowcount2

If rowcount1>0 Then
  If rowcount2>0 Then
	  Call ComparedSheets
	  WScript.Echo "Compared Results at " & strexcel
  Else
	  WScript.Echo "Warning: retrieve records from "&Formname&" in ["&DatabaseName2&"] are empty."
  End If
Else
	  WScript.Echo "Warning: retrieve records from "&Formname&" in ["&DatabaseName1&"] are empty."
End If

If rowcount2>0 Then
  WScript.Sleep(150)
  Call DuplicatedItem
Else
  WScript.Echo "Warning: retrieve records from "&Formname&" in ["&DatabaseName2&"] are empty."
End If
On Error Goto 0

WScript.Echo "End at "&Date&" "& Time
Wscript.Quit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DuplicatedItem
On Error Resume Next

Dim oexcel,openexcel,osheet2,osheet3,varsheet,sheet2rows,sheet3rows,r2,tmp,varrow,varreturnId,returnId2,varnum,color3

Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
Set varsheet=CreateObject("Excel.WorkSheet")
oexcel.DisplayAlerts=False
Set openexcel=oexcel.Workbooks.Open(strexcel)
openexcel.Activate
Set osheet2=openexcel.Worksheets(sheet_new)
varnum=0
color3=48
For Each varsheet In openexcel.Worksheets
	  If StrComp(varsheet.Name,shtname,vbTextCompare)=0 Then
		   varnum=varnum+1
		   Exit For
	  End If
Next

If varnum=1 Then
	Set osheet3=openexcel.Worksheets(shtname)
Else
	openexcel.Sheets.Add
	openexcel.ActiveSheet.name=shtname
	Set osheet3=openexcel.Worksheets(shtname)
	osheet3.Cells(1,1).Value=Formname&" Compared Results"
	osheet3.Cells(2,1).Value="Expression"
	osheet3.Cells(2,5).Value="Duplicate in "&DatabaseName2
End If

oexcel.Visible=False
sheet2rows=osheet2.UsedRange.Rows.Count
sheet3rows=osheet3.UsedRange.Rows.Count
For r2=2 To sheet2rows-1 Step 1

	tmp=osheet2.Cells(r2,5).Value
	returnId2=CInt(Trim(osheet2.Cells(r2,1).Value))
	For varrow=r2+1 To sheet2rows Step 1
		varreturnId=CInt(Trim(osheet2.Cells(varrow,1).Value))
		If varreturnId=returnId2 And osheet2.Cells(varrow,6).Value<>3 Then
			If StrComp(tmp,osheet2.Cells(varrow,5).Value)=0 And StrComp(osheet2.Cells(varrow,4).Value,osheet2.Cells(r2,4).Value, vbTextCompare)=0 Then
				sheet3rows=sheet3rows+1
				osheet3.Cells(sheet3rows,1).Value=tmp
				osheet3.Cells(sheet3rows,5).Value="Y"
				osheet2.Rows(varrow).Interior.ColorIndex=color3
				osheet2.Cells(varrow,6).Value=3
			Else
				Exit For
			End If
		Else
			Exit For
		End If
	Next
Next
openexcel.Save
oexcel.Workbooks.Close
oexcel.Quit

Set varsheet=Nothing
Set osheet2=Nothing
Set osheet3=Nothing
Set openexcel=Nothing
Set oexcel=Nothing
On Error GoTo 0

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ComparedSheets
On Error Resume Next
Dim oexcel,openexcel,osheet1,osheet2,osheet3,varsheet

Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
Set varsheet=CreateObject("Excel.WorkSheet")
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
oexcel.Visible=False

Dim r1,r2,r3,str1,str2,cmpStrs,color1,color2,col,returnId1,returnId2,sheet1rows,sheet2rows,sheet3rows
color1=28
color2=22
col=6

osheet3.Range("A1:E1").Font.Size=21
osheet3.Range("A1:E1").Font.Color=RGB(100,100,255)
osheet3.Range("A1:E1").Font.Bold=True
osheet3.Cells(1,1).Value=Formname&" Compared Results"
osheet3.Range("A1:E1").Merge
osheet3.Range("A1:E1").HorizontalAlignment=2
osheet3.Columns(1).ColumnsWidth=100
osheet3.Cells(2,1).Value="Expression in "&DatabaseName2 &" or "&DatabaseName1
osheet3.Cells(2,2).Value="Deleted"
osheet3.Cells(2,3).Value="New Add"
osheet3.Cells(2,4).Value="Modified"
osheet3.Cells(2,5).Value="Duplicate"

sheet1rows=osheet1.UsedRange.Rows.Count
sheet2rows=osheet2.UsedRange.Rows.Count
sheet3rows=osheet3.UsedRange.Rows.Count


For r1=2 To sheet1rows Step 1
	osheet1.Cells(r1,col).Value=0 '"old"
Next

For r2=2 To sheet2rows Step 1
	osheet2.Cells(r2,col).Value=0 '"new"
Next

For r1=2 To sheet1rows Step 1

	str1=Trim(osheet1.Cells(r1,5).Value) 'old
	returnId1=CInt(Trim(osheet1.Cells(r1,1).Value))
	For r2=2 To sheet2rows Step 1
	
		str2=Trim(osheet2.Cells(r2,5).Value)
		returnId2=CInt(Trim(Cosheet2.Cells(r2,1).Value))
		'    cmpStrs=StrComp(str1,str2, vbTextCompare)
		
		If osheet2.Cells(r2,col).Value=0 And returnId1=returnId2 And StrComp(str1,str2, vbTextCompare)=0 Then
			If StrComp(osheet1.Cells(r1,4).Value,osheet2.Cells(r2,4).Value, vbTextCompare)=0 Then
				osheet1.Cells(r1,col).Value=1 '"same"
				osheet2.Cells(r2,col).Value=1 '"same"
			Else
				osheet1.Cells(r1,col).Value=2 '"modified"
				osheet2.Cells(r2,col).Value=2 '"modified"
			End If
			
			Exit For
		End If
	Next
Next

r1=2
r2=2
r3=sheet3rows+1
While r1<=sheet1rows And r2<=sheet2rows
'  str1=Trim(osheet1.Cells(r1,5).Value)'old
'  str2=Trim(osheet2.Cells(r2,5).Value)'new
 
  If osheet1.Cells(r1,col).Value=1 Then
    r1=r1+1
  ElseIf osheet1.Cells(r1,col).Value=0 Then
    osheet1.Rows(r1).Interior.ColorIndex=color1
    If StrComp(osheet1.Cells(r1,4).Value,"",vbTextCompare)=0 Or StrComp(osheet1.Cells(r1,4).Value,"NULL",vbTextCompare)=0 Then
	    osheet3.Cells(r3,1).Value=Trim(osheet1.Cells(r1,5).Value)
    Else
	    osheet3.Cells(r3,1).Value=Trim(osheet1.Cells(r1,4).Value)&"="&Trim(osheet1.Cells(r1,5).Value)
    End If
    osheet3.Cells(r3,2).Value="Y"
    r3=r3+1
    r1=r1+1
  ElseIf osheet1.Cells(r1,col).Value=2 Then
    osheet1.Rows(r1).Interior.ColorIndex=color1
    osheet3.Cells(r3,1).Value="DestFld"
    osheet3.Cells(r3,4).Value="Y"
    r3=r3+1
    r1=r1+1
  End If
  
  If osheet2.Cells(r2,col).Value=1 Then
    r2=r2+1
  ElseIf osheet2.Cells(r2,col).Value=0 Then
    osheet2.Rows(r2).Interior.ColorIndex=color2
    If StrComp(osheet2.Cells(r2,4).Value,"",vbTextCompare)=0 Or StrComp(osheet2.Cells(r2,4).Value,"NULL",vbTextCompare)=0 Then
	    osheet3.Cells(r3,1).Value=Trim(osheet2.Cells(r2,5).Value)
    Else
	    osheet3.Cells(r3,1).Value=Trim(osheet2.Cells(r2,4).Value)&"="&Trim(osheet2.Cells(r2,5).Value)
    End If
    osheet3.Cells(r3,3).Value="Y"
    r3=r3+1
    r2=r2+1
  ElseIf osheet2.Cells(r2,col).Value=2 Then
    osheet2.Rows(r2).Interior.ColorIndex=color2
    osheet3.Cells(r3,1).Value="DestFld"
    osheet3.Cells(r3,4).Value="Y"
    r3=r3+1
    r2=r2+1
  End If
  
Wend

While r1<=sheet1rows
  If osheet1.Cells(r1,col).Value=0 Then
    osheet1.Rows(r1).Interior.ColorIndex=color1
    If StrComp(osheet1.Cells(r1,4).Value,"",vbTextCompare)=0 Or StrComp(osheet1.Cells(r1,4).Value,"NULL",vbTextCompare)=0 Then
	    osheet3.Cells(r3,1).Value=Trim(osheet1.Cells(r1,5).Value)
    Else
	    osheet3.Cells(r3,1).Value=Trim(osheet1.Cells(r1,4).Value)&"="&Trim(osheet1.Cells(r1,5).Value)
    End If
    osheet3.Cells(r3,2).Value="Y"
    r3=r3+1
    r1=r1+1
  ElseIf osheet1.Cells(r1,col).Value=1 Then
    r1=r1+1
  ElseIf osheet1.Cells(r1,col).Value=2 Then
    osheet1.Rows(r1).Interior.ColorIndex=color1
    osheet3.Cells(r3,1).Value="DestFld"
    osheet3.Cells(r3,4).Value="Y"
    r3=r3+1
    r1=r1+1
  End If
Wend

While r2<=sheet2rows
  If osheet2.Cells(r2,col).Value=0 Then
    osheet2.Rows(r2).Interior.ColorIndex=color2
    If StrComp(osheet2.Cells(r2,4).Value,"",vbTextCompare)=0 Or StrComp(osheet2.Cells(r2,4).Value,"NULL",vbTextCompare)=0 Then
	    osheet3.Cells(r3,1).Value=Trim(osheet2.Cells(r2,5).Value)
    Else
	    osheet3.Cells(r3,1).Value=Trim(osheet2.Cells(r2,4).Value)&"="&Trim(osheet2.Cells(r2,5).Value)
    End If
    osheet3.Cells(r3,3).Value="Y"
    r3=r3+1
    r2=r2+1
  ElseIf osheet2.Cells(r2,col).Value=2 Then
    osheet2.Rows(r2).Interior.ColorIndex=color2
    osheet3.Cells(r3,1).Value="DestFld"
    osheet3.Cells(r3,4).Value="Y"
    r3=r3+1
    r2=r2+1
  ElseIf osheet2.Cells(r2,col).Value=1 Then
    r2=r2+1
  End If
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
Function RWData(constr,strQuery,excelfl,sheetname)
On Error Resume Next
'Step1
Dim objConn,objRS,i,j,rowcount,colcount
Dim fso,oexcel,openexcel,osheet,varsheet
rowcount=-1
Set objConn=CreateObject("ADODB.Connection")
Set objRS=CreateObject("ADODB.Recordset")
Set fso=CreateObject("Scripting.FileSystemObject")

'strQuery="SELECT [ReturnId],[ExpId],[ExpOrder],[DestFld],[Expression] FROM "&tbl

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
	 osheet.Cells(1,1).Value="ReturnId"
	 osheet.Cells(1,2).Value="ExpId"
	 osheet.Cells(1,3).Value="ExpOrder"
	 osheet.Cells(1,4).Value="DestFld"
	 osheet.Cells(1,5).Value="Expression"
	 osheet.Cells(1,6).Value="Flag"
	'Step5
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

