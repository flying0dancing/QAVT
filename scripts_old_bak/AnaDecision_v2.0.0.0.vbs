'********************************************************************
'	if any bug please email kunshen@lombardrisk.com,Thank you.
'********************************************************************

Option Explicit
Dim file,rowcount,fso,oexcel,openexcel,osheet,logfile,rh,sheetname

If WScript.Arguments.length=3 Then
	file=WScript.Arguments(0)
	sheetname=WScript.Arguments(1)
'	rowcount=WScript.Arguments(2)
	logfile=WScript.Arguments(2)
Else
	MsgBox "Miss arguments! Order is full path of excel,Sheet name,full path of log" & vbLf & "AnaDecision.vbs is terminated...."
	WScript.Quit
End If

On Error Resume Next
Set fso=CreateObject("Scripting.FileSystemObject")
Set rh=fso.OpenTextFile(logfile,8,True)
rh.WriteLine("*********************	Next is error information *******************************************************")
	
Set oexcel=CreateObject("excel.Application")
oexcel.DisplayAlerts=False
Set openexcel=oexcel.Workbooks.Open(file,0,false)
openexcel.Activate
oexcel.Visible=false
Set osheet=openexcel.Worksheets(sheetname)

rowcount=osheet.UsedRange.Rows.Count
'WScript.Echo "row count is "&rowcount
If rowcount<=0 Then
	MsgBox "rows of "&sheetname&" less then zero" & vbLf & "AnaDecision.vbs is terminated...."
	WScript.Quit
End If

Dim row,column,subrow,outstr,arr1(9),arr2(9),i,color1,color2,color3,color4,errcode,entity,errcount
color1=7
color2=4
color3=8
color4=12
errcount=0

row=1
While row<CInt(rowcount)

 subrow=row+1
 If osheet.Cells(row,1).Value<>"" Then
 If StrComp(osheet.Cells(row,1).Value, osheet.Cells(subrow,1).Value, vbTextCompare)=0 And StrComp(osheet.Cells(row,2).Value, osheet.Cells(subrow,2).Value, vbTextCompare)=0 Then
 If StrComp(osheet.Cells(row,14).Value, osheet.Cells(subrow,14).Value, vbTextCompare)=0 and StrComp(osheet.Cells(row,15).Value, osheet.Cells(subrow,15).Value, vbTextCompare)=0 and StrComp(osheet.Cells(row,16).Value, osheet.Cells(subrow,16).Value, vbTextCompare)=0 Then

	outstr="EXCEL ROWNO:	" & row &"	vs	"&subrow&vbCrLf&"STBDMPRECORDNO: " & osheet.Cells(row,3).Value &"	vs	"& osheet.Cells(subrow,3).Value & vbCrLf & "ERROR: Occurred errcode"
    
	For i=0 To 9
     arr1(i)=osheet.Cells(row,i+4).Value
     arr2(i)=osheet.Cells(subrow,i+4).Value
    Next
	
    If Compares("0,1,2,3,4,5,6,7,8,9",arr1,arr2)=0 Then
    
        errcount=errcount+1
    	osheet.Rows(row).Interior.ColorIndex= 10
    	osheet.Rows(subrow).Interior.ColorIndex= 10
    	rh.WriteBlankLines(1)
    	rh.WriteLine(outstr & "[5], records are as same as each other.")
       
     ElseIf Compares("0,1,2,3,4,5,6,7",arr1,arr2)=0 Then
       
        errcode=HasError(arr1(8),arr1(9),arr2(8),arr2(9))
        
        If errcode<>0 Then 
          errcount=errcount+1
          rh.WriteBlankLines(1)
          entity="INST"
		  Call PaintColor
		End If
        
     ElseIf Compares("0,1,2,3,4,5,8,9",arr1,arr2)=0 Then

     	errcode=HasError(arr1(6),arr1(7),arr2(6),arr2(7))
     	
     	If errcode<>0 Then 
     	  errcount=errcount+1
     	  rh.WriteBlankLines(1)
     	  entity="GLC4"
		  Call PaintColor
     	End If 
	 	
     ElseIf Compares("0,1,2,3,6,7,8,9",arr1,arr2)=0 Then

     	errcode=HasError(arr1(4),arr1(5),arr2(4),arr2(5))
     	
     	If errcode<>0 Then 
     	  errcount=errcount+1
     	  rh.WriteBlankLines(1)
     	  entity="GLC3"
		  Call PaintColor
    	End If
     
     ElseIf Compares("0,1,4,5,6,7,8,9",arr1,arr2)=0 Then
     
        errcode=HasError(arr1(2),arr1(3),arr2(2),arr2(3))
        
        If errcode<>0 Then 
          errcount=errcount+1
     	  rh.WriteBlankLines(1)
     	  entity="GLC2"
		  Call PaintColor
     	End If

     ElseIf Compares("2,3,4,5,6,7,8,9",arr1,arr2)=0 Then
     
     	errcode=HasError(arr1(0),arr1(1),arr2(0),arr2(1))
     	
     	If errcode<>0 Then
     	  errcount=errcount+1
     	  rh.WriteBlankLines(1)
     	  entity="GLC1"
		  Call PaintColor
    	End If

    End If
    
   row=row+1
  Else
  
   row=row+1 
  End If
 Else 
   row=row+1 
 End If
 End If
Wend

if not oexcel.ActiveWorkbook.Saved then 
openexcel.Save
End If
'oexcel.Visible=True
oexcel.Workbooks.Close
oexcel.Quit

rh.WriteBlankLines(2)
rh.Write("Totally error count:"&errcount)
rh.Close

Set fso=Nothing
Set osheet=Nothing
Set oexcel=Nothing
Set openexcel=Nothing
Set rh=Nothing
On Error GoTo 0

WScript.Quit


'''''''''''''''''''''''''''''''''''''''''''''''
Sub PaintColor
      Select Case errcode
       Case 1 osheet.Rows(row).Interior.ColorIndex=color1:osheet.Rows(subrow).Interior.ColorIndex=color1:rh.WriteLine(outstr & "[1], STBDMP"& entity &"NOT are different and STBDMP"& entity &" are the same.")
       Case 2 osheet.Rows(row).Interior.ColorIndex=color2:osheet.Rows(subrow).Interior.ColorIndex=color2:rh.WriteLine(outstr & "[2], STBDMP"& entity &"NOT are 'N' and STBDMP"& entity &" have intersection.")
       Case 3 osheet.Rows(row).Interior.ColorIndex=color3:osheet.Rows(subrow).Interior.ColorIndex=color3:rh.WriteLine(outstr & "[3], STBDMP"& entity &"NOT at least have a 'Y' and STBDMP"& entity &" are different.")
       Case 4 osheet.Rows(row).Interior.ColorIndex=color4:osheet.Rows(subrow).Interior.ColorIndex=color4:rh.WriteLine(outstr & "[4], STBDMP"& entity &"NOT are 'N' and STBDMP"& entity &" at least have a -1.")
      
	  End Select
End Sub

Function Compares(str,arr1,arr2)
Dim sec,i,count
sec=Split(str,",",-1,1)
count=-1
For i=0 To UBound(sec)
 If StrComp(arr1(sec(i)),arr2(sec(i)),vbTextCompare)=0 Then
  count=count+1
 End If
Next
If count=UBound(sec) Then 
Compares=0
Else 
Compares=-1
End If
End Function


Function Intersection(str1,str2)
Dim arr1,arr2,c,i,j
Dim arr3()
arr1=Split(str1,",",-1,1)
arr2=Split(str2,",",-1,1)
c=0
For i=0 To UBound(arr1)
 For j=0 To UBound(arr2)
  If arr1(i)=arr2(j) Then 
ReDim Preserve arr3(c)
   arr3(c)=arr1(i)
   c=c+1
  End If   
 Next
Next
Intersection=Join(arr3,",")
End Function


Function HasError(glcnot1,glc1,glcnot2,glc2)
     HasError=0
     If StrComp(glc1,glc2,vbTextCompare)=0 And StrComp(glcnot1,glcnot2,vbTextCompare)<>0 Then
       HasError=1
     ElseIf Len(Intersection(glc1,glc2))>0 And StrComp(glcnot1,glcnot2,vbTextCompare)=0 And StrComp(glcnot1,"N",vbTextCompare)=0 Then
       HasError=2
     ElseIf Len(Intersection(glc1,glc2))<=0 And (StrComp(glcnot1,"Y",vbTextCompare)=0 Or StrComp(glcnot2,"Y",vbTextCompare)=0) Then
       HasError=3
     ElseIf StrComp(glcnot1,glcnot2,vbTextCompare)=0 And StrComp(glcnot1,"N",vbTextCompare)=0 And (StrComp(glc1,"-1",vbTextCompare)=0 Or StrComp(glc2,"-1",vbTextCompare)=0) Then
       HasError=4
     End If
End Function