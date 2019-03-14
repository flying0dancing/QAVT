'********************************************************************
'	if any bug please email kunshen@lombardrisk.com,Thank you.
'********************************************************************

Option Explicit
Dim file,rowcount,fso,ws,wsysenv,oexcel,openexcel,logfile,rh

If WScript.Arguments.length=2 Then
	file=WScript.Arguments(0)
	rowcount=WScript.Arguments(1)
Else
	MsgBox "Miss arguments! Order is full path of excel, count of rows you want to compared" & vbLf & "AnaDecision.vbs is terminated...."
	WScript.Quit 
End If

Set fso=CreateObject("Scripting.FileSystemObject")
Set ws=CreateObject("wscript.shell")
Set wsysenv= ws.Environment("Process")
logfile=wsysenv("TMP") & "\LombardRisk_DecisionTable_Error_" & Year(Now)&Month(Now)&Day(Now)&hour(now)&minute(now)&Second(now)&".log"
Set rh=fso.OpenTextFile(logfile,2,True)
rh.WriteLine("*********************This is error information for LombardRisk QA Decision Table Analysis******************")
	
Set oexcel=CreateObject("excel.Application")
Set openexcel=oexcel.Workbooks.Open(file,0,false)

oexcel.Visible=false
oexcel.Worksheets("Sheet1").Activate

Dim row,column,subrow,outstr,arr1(9),arr2(9),i,color1,color2,color3,color4,errcode,entity,errcount
color1=7
color2=4
color3=8
color4=12
errcount=0

row=1
While row<CInt(rowcount)

 subrow=row+1
 If oexcel.Cells(row,1).Value<>"" Then
 If StrComp(oexcel.Cells(row,1).Value, oexcel.Cells(subrow,1).Value, vbTextCompare)=0 And StrComp(oexcel.Cells(row,2).Value, oexcel.Cells(subrow,2).Value, vbTextCompare)=0 Then
 If StrComp(oexcel.Cells(row,14).Value, oexcel.Cells(subrow,14).Value, vbTextCompare)=0 and StrComp(oexcel.Cells(row,15).Value, oexcel.Cells(subrow,15).Value, vbTextCompare)=0 and StrComp(oexcel.Cells(row,16).Value, oexcel.Cells(subrow,16).Value, vbTextCompare)=0 Then

	outstr="EXCEL ROWNO:	" & row &"	vs	"&subrow&vbCrLf&"STBDMPRECORDNO: " & oexcel.Cells(row,3).Value &"	vs	"& oexcel.Cells(subrow,3).Value & vbCrLf & "ERROR: Occurred errcode"
    
	For i=0 To 9
     arr1(i)=oexcel.Cells(row,i+4).Value
     arr2(i)=oexcel.Cells(subrow,i+4).Value
    Next
	
    If Compares("0,1,2,3,4,5,6,7,8,9",arr1,arr2)=0 Then
    
        errcount=errcount+1
    	oexcel.ActiveSheet.Rows(row).Interior.ColorIndex= 10
    	oexcel.ActiveSheet.Rows(subrow).Interior.ColorIndex= 10
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
oexcel.Visible=True
Rem oexcel.Workbooks.Close
Rem oexcel.Quit

rh.WriteBlankLines(2)
rh.WriteLine("Totally error count:"&errcount)
rh.Close

Set fso=Nothing
Set ws=Nothing
Set wsysenv=Nothing
Set oexcel=Nothing
Set openexcel=Nothing
Set rh=Nothing

Sub PaintColor
      Select Case errcode
       Case 1 oexcel.ActiveSheet.Rows(row).Interior.ColorIndex=color1:oexcel.ActiveSheet.Rows(subrow).Interior.ColorIndex=color1:rh.WriteLine(outstr & "[1], STBDMP"& entity &"NOT are different and STBDMP"& entity &" are the same.")
       Case 2 oexcel.ActiveSheet.Rows(row).Interior.ColorIndex=color2:oexcel.ActiveSheet.Rows(subrow).Interior.ColorIndex=color2:rh.WriteLine(outstr & "[2], STBDMP"& entity &"NOT are 'N' and STBDMP"& entity &" have intersection.")
       Case 3 oexcel.ActiveSheet.Rows(row).Interior.ColorIndex=color3:oexcel.ActiveSheet.Rows(subrow).Interior.ColorIndex=color3:rh.WriteLine(outstr & "[3], STBDMP"& entity &"NOT at least have a 'Y' and STBDMP"& entity &" are different.")
       Case 4 oexcel.ActiveSheet.Rows(row).Interior.ColorIndex=color4:oexcel.ActiveSheet.Rows(subrow).Interior.ColorIndex=color4:rh.WriteLine(outstr & "[4], STBDMP"& entity &"NOT are 'N' and STBDMP"& entity &" at least have a -1.")
       'Case 1 oexcel.Cells(row,col).Interior.ColorIndex=color1:oexcel.Cells(subrow,col).Interior.ColorIndex=color1:oexcel.Cells(row,col+1).Interior.ColorIndex=color1:oexcel.Cells(subrow,col+1).Interior.ColorIndex=color1:rh.WriteLine(outstr & " [STBDMP"& entity &"NOT] are different and [STBDMP"& entity &"] are the same.")
       'Case 2 oexcel.Cells(row,col).Interior.ColorIndex=color2:oexcel.Cells(subrow,col).Interior.ColorIndex=color2:oexcel.Cells(row,col+1).Interior.ColorIndex=color2:oexcel.Cells(subrow,col+1).Interior.ColorIndex=color2:rh.WriteLine(outstr & " [STBDMP"& entity &"NOT] are 'N' and [STBDMP"& entity &"] have intersection.")
       'Case 3 oexcel.Cells(row,col).Interior.ColorIndex=color3:oexcel.Cells(subrow,col).Interior.ColorIndex=color3:oexcel.Cells(row,col+1).Interior.ColorIndex=color3:oexcel.Cells(subrow,col+1).Interior.ColorIndex=color3:rh.WriteLine(outstr & " [STBDMP"& entity &"NOT] at least have an 'Y' and [STBDMP"& entity &"] are different.")
       'Case 4 oexcel.Cells(row,col).Interior.ColorIndex=color4:oexcel.Cells(subrow,col).Interior.ColorIndex=color4:oexcel.Cells(row,col+1).Interior.ColorIndex=color4:oexcel.Cells(subrow,col+1).Interior.ColorIndex=color4:rh.WriteLine(outstr & " [STBDMP"& entity &"NOT] are 'N' and [STBDMP"& entity &"] have an -1.")
      
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