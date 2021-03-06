''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  develop by Kun Shen, send email to Kun.Shen@lombardrisk.com if any issue 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strexcel,DatabaseName1,DatabaseName2,Formname,ServerName,User,Password,constr1,constr2,rowcount1,rowcount2,sheet_new,sheet_old,strQuery
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

rowcount1=RWData(constr1,strQuery,strexcel,sheet_old)
WScript.Sleep(100)
rowcount2=RWData(constr2,strQuery,strexcel,sheet_new)

If rowcount1>0 And rowcount2>0 Then
  Call QP(sheet_old,rowcount1)
  Call QP(sheet_new,rowcount2)
  Call ComparedSheets
  WScript.Echo "Compared Results at " & strexcel
Else
  WScript.Echo "Warning: retrieve records from "&Formname&" in ["&DatabaseName1&"] or ["&DatabaseName2&"] are empty."
End If  
Wscript.Quit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub QP(ByVal sheetname,ByVal rowcount)
On Error Resume Next
Dim oexcel,openexcel,osheet,i,j,k,tmp,flag
Set oexcel=CreateObject("Excel.Application")
Set openexcel=oexcel.Workbooks.Open(strexcel)
openexcel.Activate
Set osheet=openexcel.Worksheets(sheetname)
oexcel.Visible=False
  For i=1 To rowcount-1
    flag=0
    For j=1 To rowcount-i
      If StrComp(osheet.Cells(j,5).Value,osheet.Cells(j+1,5).Value, vbTextCompare)>0 Then
        For k=1 To 5
          tmp=osheet.Cells(j,k).Value
          osheet.Cells(j,k).Value=osheet.Cells(j+1,k).Value
          osheet.Cells(j+1,k).Value=tmp
        Next
        flag=1
      End If
    Next
    If flag=0 Then 
     Exit For
    End If
  Next
openexcel.Save
oexcel.Workbooks.Close
oexcel.Quit
Set osheet=Nothing
Set openexcel=Nothing
Set oexcel=Nothing
On Error GoTo 0
End Sub
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

Dim r1,r2,r3,str1,str2,cmpStrs,color1,color2,color3
r1=1
r2=1
r3=3
color1=28
color2=22
'color1=38
'color2=32
color3=50
osheet3.Range("A1:D1").Font.Size=21
osheet3.Range("A1:D1").Font.Color=RGB(100,100,255)
osheet3.Range("A1:D1").Font.Bold=True
osheet3.Cells(1,1).Value=Formname&" Compared Results"
osheet3.Range("A1:D1").Merge
osheet3.Range("A1:D1").HorizontalAlignment=2
osheet3.Cells(2,1).Value="Expression"
osheet3.Cells(2,2).Value="Deleted in "&DatabaseName2
osheet3.Cells(2,3).Value="New Add in "&DatabaseName2
osheet3.Cells(2,4).Value="Modified in "&DatabaseName2
While r1<=CInt(rowcount1) And r2<=CInt(rowcount2)
  str1=Trim(osheet1.Cells(r1,5).Value)'old
  str2=Trim(osheet2.Cells(r2,5).Value)'new
  cmpStrs=StrComp(str1,str2, vbTextCompare)
  
  If cmpStrs<0 Then

    osheet1.Rows(r1).Interior.ColorIndex=color1
'    osheet3.Cells(r3,1).Value=str1 & " is deleted in """ & osheet2.name&"""."
    osheet3.Cells(r3,1).Value=str1
    osheet3.Cells(r3,2).Value="Y"
    r3=r3+1
    r1=r1+1

  ElseIf cmpStrs>0 Then

    osheet2.Rows(r2).Interior.ColorIndex=color2
'    osheet3.Cells(r3,1).Value=str2 & " is new added in """ & osheet2.name&"""."
    osheet3.Cells(r3,1).Value=str2
    osheet3.Cells(r3,3).Value="Y"
    r3=r3+1
    r2=r2+1

  Else

    Dim comparedStrs,col
      col=4
      comparedStrs=StrComp(osheet1.Cells(r1,col).Value,osheet2.Cells(r2,col).Value, vbTextCompare)
      
      If comparedStrs<>0 Then
        osheet1.Rows(r1).Interior.ColorIndex=color3
        osheet2.Rows(r2).Interior.ColorIndex=color3
        osheet3.Cells(r3,1).Value=str1 & " is modified. "&osheet1.name&".cell(" & r1 &"," & col & ")=" & osheet1.Cells(r1,col).Value &" ==>> "&osheet2.name&".cell(" & r2 &"," & col & ")=" & osheet2.Cells(r2,col).Value
        osheet3.Cells(r3,4).Value="Y"
        r3=r3+1
      End If
      
    r1=r1+1
    r2=r2+1
  End If
  
Wend

While r1<=CInt(rowcount1)
    osheet1.Rows(r1).Interior.ColorIndex=color1
'    osheet3.Cells(r3,1).Value=osheet1.Cells(r1,3) & " is deleted in """ & osheet2.name&"""."
    osheet3.Cells(r3,1).Value=str1
    osheet3.Cells(r3,2).Value="Y"
    r3=r3+1
    r1=r1+1
Wend

While r2<=CInt(rowcount2)
    osheet2.Rows(r2).Interior.ColorIndex=color2
    'osheet3.Cells(r3,1).Value=osheet1.Cells(r2,3)& " is new added in """ & osheet2.name&"""."
    osheet3.Cells(r3,1).Value=str2
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
Else
  'step2
  objRS.Open strQuery,objConn,1,1
  If Err.Number<>0 Then
    WScript.Echo "Error: "&CStr(Err.Number)& Err.Description & vbCrLf & Err.Source
    Err.Clear
  Else
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

  End If
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

