''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  develop by Kun Shen, send email to Kun.Shen@lombardrisk.com if any issue 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strexcel,DatabaseName1,DatabaseName2,Formname,ServerName,User,Password,constr1,constr2,rowcount1,rowcount2,sheet_new,sheet_old
If WScript.Arguments.length=5 Then
 strexcel=WScript.Arguments(0)
 ServerName=WScript.Arguments(1)
 DatabaseName1=WScript.Arguments(2)
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
sheet_old=DatabaseName1 &" "& Formname
sheet_new=DatabaseName2 &" "& Formname
rowcount1=RWData(constr1,Formname,strexcel,sheet_old)
'WScript.Echo "rowcount1: "&rowcount1
WScript.Sleep(100)
rowcount2=RWData(constr2,Formname,strexcel,sheet_new)
'WScript.Echo "rowcount2: "&rowcount2
Call ComparedSheets
WScript.Echo "Compared Results at " & strexcel
Wscript.Quit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ComparedSheets
Dim oexcel,openexcel,osheet1,osheet2,osheet3
Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
'oexcel.DisplayAlerts=False
Set openexcel=oexcel.Workbooks.Open(strexcel)
openexcel.Activate
Set osheet1=openexcel.Worksheets(sheet_old)
Set osheet2=openexcel.Worksheets(sheet_new)
openexcel.Sheets.Add
openexcel.ActiveSheet.name=Formname&" Compared Results"
Set osheet3=openexcel.Worksheets(Formname&" Compared Results")
oexcel.Visible=false

Dim r1,r2,r3,str1,str2,cmpStrs,color1,color2,color3
r1=1
r2=1
r3=1
color1=30
color2=25
color3=50
While r1<=CInt(rowcount1) And r2<=CInt(rowcount2)
  str1=Trim(osheet1.Cells(r1,3).Value)
  str2=Trim(osheet2.Cells(r2,3).Value)

  If CInt(str1)<CInt(str2) Then
'    WScript.Echo str1& "<" & str2&"::::"&r1 & "vs" &r2
    osheet1.Rows(r1).Interior.ColorIndex=color1
    osheet3.Cells(r3,1).Value=str1 & " is deleted in """ & osheet2.name&"""."
    r3=r3+1
    r1=r1+1

  ElseIf CInt(str1)>CInt(str2) Then
'    WScript.Echo str1& ">" & str2&"::::"&r1 & "vs" &r2
    osheet2.Rows(r2).Interior.ColorIndex=color2
    osheet3.Cells(r3,1).Value=str2 & " is new added in """ & osheet2.name&"""."
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
        r3=r3+1
      End If
      
    Next
    r1=r1+1
    r2=r2+1
  End If
  
Wend

While r1<=CInt(rowcount1)
    osheet1.Rows(r1).Interior.ColorIndex=color1
    osheet3.Cells(r3,1).Value=osheet1.Cells(r1,3) & " is deleted in """ & osheet2.name&"""."
    r3=r3+1
    r1=r1+1
Wend

While r2<=CInt(rowcount2)
    osheet2.Rows(r2).Interior.ColorIndex=color2
    osheet3.Cells(r3,1).Value=osheet1.Cells(r2,3)& " is new added in """ & osheet2.name&"""."
    r3=r3+1
    r2=r2+1
Wend

openexcel.Save
oexcel.Visible=true
oexcel.Workbooks.Close
oexcel.Quit
Set osheet1=Nothing
Set osheet2=Nothing
Set osheet3=Nothing
Set openexcel=Nothing
Set oexcel=Nothing
End Sub

''''''''''''''''''''''''''''''''''
Function RWData(constr,tbl,excelfl,sheetname)
'Step1
Dim objConn
Set objConn=CreateObject("ADODB.Connection")
objConn.Open constr
'step2
Dim strQuery,objRS
Set objRS=CreateObject("ADODB.Recordset")
strQuery="select c.STBDMPFORM,c.STBDMPITEM,m.* from dbo."&tbl&"_CHILD as c , dbo."&tbl&"_MASTER as m where c.STBDMPRECORDNO=m.STBDMPRECORDNO " &_
"order by c.STBDMPRECORDNO"
objRS.Open strQuery,objConn,1,1
'step3
Dim fso
Set fso=CreateObject("Scripting.FileSystemObject")

'step4
Dim oexcel,openexcel,osheet
Set oexcel= CreateObject("Excel.Application")'创建EXCEL对象
oexcel.DisplayAlerts=False
If fso.FileExists(excelfl) Then
Set openexcel=oexcel.Workbooks.Open(excelfl)
openexcel.Activate
oexcel.Visible=False
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

'step5
Dim i,j,rowcount,colcount
rowcount=objRS.RecordCount
colcount=objRS.Fields.Count

i=0
objRS.MoveFirst
  Do While Not objRS.EOF
    For j=0 To colcount-1
     osheet.Cells(i+1,j+1).Value= objRS.Fields(j).Value
    Next
     i=i+1
  objRS.MoveNext
  Loop

If Not oexcel.ActiveWorkbook.Saved Then
openexcel.Save
End If

objConn.Close
oexcel.Workbooks.Close
oexcel.Quit

Set fso=Nothing
Set objws=Nothing
set objConn=nothing
Set objRS=Nothing
Set osheet=Nothing
Set oexcel=Nothing
Set openexcel=Nothing
'step6
RWData=rowcount
End Function

