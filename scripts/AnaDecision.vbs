'********************************************************************
'	if any bug please email kunshen@lombardrisk.com,Thank you.
'********************************************************************

Option Explicit
Dim file,fso,oexcel,openexcel,osheet,logfile,log_tmp,rh,sheetname,Forms_str,Forms_arr,inifile

If WScript.Arguments.length=3 Then
	file=Trim(WScript.Arguments(0))
	logfile=Trim(WScript.Arguments(1))
	inifile=Trim(WScript.Arguments(2))
Else
	WScript.Echo "Miss arguments! Order is full path of excel,log,ini file" & vbLf & "AnaDecision.vbs is terminated...."
	WScript.Quit(-1)
End If
On Error Resume Next
Forms_str=ReadINI(inifile,"_QAVT_CONFIG_FORM")
Set fso=CreateObject("Scripting.FileSystemObject")
Set oexcel=CreateObject("excel.Application")
oexcel.DisplayAlerts=False
Set openexcel=oexcel.Workbooks.Open(file,0,false)
openexcel.Activate
oexcel.Visible=False

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

For Each sheetname In Forms_arr
	If sheetname<>"" Then 
		logfile= log_tmp &"_"& sheetname & ".log"
		Set rh=fso.OpenTextFile(logfile,8,True)
		rh.WriteLine("*********************	Next is error information *******************************************************")
		Set osheet=openexcel.Worksheets(sheetname)
		Call AnalyzeDecisionFunc(rh,osheet)
		rh.Close
	End If
Next

if not oexcel.ActiveWorkbook.Saved then 
openexcel.Save
End If
'oexcel.Visible=True
oexcel.Workbooks.Close
oexcel.Quit

Set fso=Nothing
Set osheet=Nothing
Set oexcel=Nothing
Set openexcel=Nothing
Set rh=Nothing
On Error GoTo 0
WScript.Quit

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'===================================================
'''''''''''''''''''''''''''''''''''''''''''''''''''''	
Sub AnalyzeDecisionFunc(ByRef rh,ByRef osheet)
Dim rowcount,colcount
rowcount=CInt(osheet.UsedRange.Rows.Count)
colcount=CInt(osheet.UsedRange.Columns.Count)

If rowcount<=1 Then
	rh.WriteLine("ERROR:rows of "&sheetname&" less then zero")
	Exit Sub
End If

Dim row,column,subrow,outstr,arr1,arr2,i,errcode,entity,errcount
errcount=0
row=2
While row<rowcount

 subrow=row+1
 If osheet.Cells(row,1).Value<>"" Then
 If StrComp(osheet.Cells(row,1).Value, osheet.Cells(subrow,1).Value, vbTextCompare)=0 And StrComp(osheet.Cells(row,2).Value, osheet.Cells(subrow,2).Value, vbTextCompare)=0 Then
 If StrComp(osheet.Cells(row,colcount-2).Value, osheet.Cells(subrow,colcount-2).Value, vbTextCompare)=0 and StrComp(osheet.Cells(row,colcount-1).Value, osheet.Cells(subrow,colcount-1).Value, vbTextCompare)=0 and StrComp(osheet.Cells(row,colcount).Value, osheet.Cells(subrow,colcount).Value, vbTextCompare)=0 Then

	outstr="EXCEL ROWNO:	" & row &"	vs	"&subrow&vbCrLf&"STBDMPRECORDNO: " & osheet.Cells(row,3).Value &"	vs	"& osheet.Cells(subrow,3).Value & vbCrLf & "ERROR: Occurred errcode"
    If colcount=16 Then
      ReDim arr1(9),arr2(9)
	  For i=0 To 9
       arr1(i)=osheet.Cells(row,i+4).Value
       arr2(i)=osheet.Cells(subrow,i+4).Value
      Next
    ElseIf colcount=20 Then
      ReDim arr1(13),arr2(13)
	  For i=0 To 13
       arr1(i)=osheet.Cells(row,i+4).Value
       arr2(i)=osheet.Cells(subrow,i+4).Value
      Next    
    ElseIf colcount=24 Then
      ReDim arr1(17),arr2(17)
	  For i=0 To 17
       arr1(i)=osheet.Cells(row,i+4).Value
       arr2(i)=osheet.Cells(subrow,i+4).Value
      Next    
    End If
	
	Dim ruleCode,countEqual,countNotEqual,countIntersect,errColumnFlag,ruleTotalNum
	countEqual=0:countNotEqual=0:countIntersect=0:ruleTotalNum=(UBound(arr1)+1)/2
	
	For i=0 To UBound(arr1)-1 Step 2
		ruleCode=GetRuleCode(arr1(i),arr1(i+1),arr2(i),arr2(i+1))
	
		If ruleCode=0 Then 
			countEqual=countEqual+1
			If countEqual=ruleTotalNum Then errcode=100 End If 
		End If
		If ruleCode=-1 Then countNotEqual=countNotEqual+1 End If
		If ruleCode=1 Then 
			countIntersect=countIntersect+1
			errColumnFlag=i
			If (colcount=16 And errColumnFlag=8) Or (colcount=20 And errColumnFlag=12) Or (colcount=24 And errColumnFlag=16) Then
          		entity="INST"
          	Else	
          		entity="GLC"&errColumnFlag/2+1
          	End If
			errcode=HasError(arr1(i),arr1(i+1),arr2(i),arr2(i+1))
			
		End If
	Next
	If (countIntersect>=1 And countNotEqual=0) Or (countIntersect=0 And countNotEqual=0 And countEqual=ruleTotalNum) Then 
		errcount=errcount+1
		rh.WriteBlankLines(1)
		Call PaintColor(errcode,osheet,row,subrow,rh,outstr,entity)
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

rh.WriteBlankLines(2)
rh.WriteLine("Totally error count:"&errcount)

End Sub


Sub PaintColor(ByRef errcode,ByRef osheet,ByRef row,ByRef subrow,ByRef rh,ByRef outstr,ByRef entity)
	Dim color1,color2,color3,color4,color100
	color1=7:color2=4:color3=8:color4=12:color100=18
	Select Case errcode
       Case 1 osheet.Rows(row).Interior.ColorIndex=color1:osheet.Rows(subrow).Interior.ColorIndex=color1:rh.WriteLine(outstr & "[1], STBDMP"& entity &"NOT are the same and STBDMP"& entity &" have intersection, but they are not same.")
       Case 2 osheet.Rows(row).Interior.ColorIndex=color2:osheet.Rows(subrow).Interior.ColorIndex=color2:rh.WriteLine(outstr & "[2], STBDMP"& entity &"NOT at least has a 'Y' and STBDMP"& entity &" are different.")
       Case 3 osheet.Rows(row).Interior.ColorIndex=color3:osheet.Rows(subrow).Interior.ColorIndex=color3:rh.WriteLine(outstr & "[3], Contains STBDMP"& entity &"NOT is 'N' and STBDMP"& entity &" is -1, so they have intersection.")
       Case 4 osheet.Rows(row).Interior.ColorIndex=color4:osheet.Rows(subrow).Interior.ColorIndex=color4:rh.WriteLine(outstr & "[4], STBDMP"& entity &"NOT is 'Y' and STBDMP"& entity &" have intersection, but they are not same and intersection is equal to Y's GLC.")
       Case 100 osheet.Rows(row).Interior.ColorIndex=color100:osheet.Rows(subrow).Interior.ColorIndex=color100:rh.WriteLine(outstr & "[100], records are as same as each other.")
       Case 0 osheet.Rows(row).Interior.ColorIndex=color1:osheet.Rows(subrow).Interior.ColorIndex=color1:rh.WriteLine(outstr & "[0], please text or email kun shen, this error is not covered!!!")
	End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''
'===============OldCompares add "Old" in function name===========================
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
Function ComparesOld(str,arr1,arr2)
Dim sec,i,j,count,secnum
sec=Split(str,",",-1,1)
secnum=UBound(sec)
'WScript.Echo("UBound(sec):"&secnum)
count=0
For i=0 To UBound(arr1)

 If StrComp(arr1(i),arr2(i),vbTextCompare)=0 Then
  count=count+1
	If secnum>-1 Then
		For j=0 To UBound(sec)
			If StrComp(i,sec(j),vbTextCompare)=0 Then 
				count=count-1
				secnum=secnum-1 
			End If
		Next
	End If
 End If

 
Next
If count=UBound(arr1)-UBound(sec) Then
'WScript.Echo("count=UBound(arr1)-UBound(sec):"&count) 
Compares=0
Else 
Compares=-1
End If
End Function

' if str1,str2 have intersection, return their intersection.
Function Intersection(str1,str2)
Dim arr1,arr2,c,i,j,k,str11,str21
Dim arr3()
str11=DelDuplicateItems(str1)
str21=DelDuplicateItems(str2)
arr1=Split(str11,",",-1,1)
arr2=Split(str21,",",-1,1)
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


Function DelDuplicateItems(str)
Dim arr,c,i,j,k,strReplaced
strReplaced=Replace(str,"'","",1,-1,vbTextCompare)
arr=Split(strReplaced,",",-1,1)
c=0
c=UBound(arr)
For i=0 To c-1
 For j=i+1 To c
  If Trim(arr(i))=Trim(arr(j)) Then
  	For k=j To c-1
  		arr(k)=arr(k+1)
  	Next
  	c=c-1
  End If
 Next
Next
ReDim Preserve arr(c)

DelDuplicateItems=Join(arr,",")
End Function



'====================================
Function HasError(glcnot1,glc1,glcnot2,glc2)
     HasError=0
     If (StrComp(glc1,"-1",vbTextCompare)=0 And StrComp(glcnot1,"N",vbTextCompare)=0 And StrComp(glc2,"-1",vbTextCompare)<>0) Or (StrComp(glc2,"-1",vbTextCompare)=0 And StrComp(glcnot2,"N",vbTextCompare)=0 And StrComp(glc1,"-1",vbTextCompare)<>0) Then
       HasError=3
     ElseIf Len(Intersection(glc1,glc2))<=0 And (StrComp(glcnot1,"Y",vbTextCompare)=0 Or StrComp(glcnot2,"Y",vbTextCompare)=0) Then
       HasError=2
     ElseIf StrComp(glcnot1,glcnot2,vbTextCompare)<>0 And StrComp(glcnot1,"Y",vbTextCompare)=0 And Len(Intersection(glc1,glc2))>0 And Len(DelDuplicateItems(glc1))=Len(Intersection(glc1,glc2)) And Len(DelDuplicateItems(glc2))<>Len(Intersection(glc1,glc2)) Then
       HasError=4
       ElseIf StrComp(glcnot1,glcnot2,vbTextCompare)<>0 And StrComp(glcnot2,"Y",vbTextCompare)=0 And Len(Intersection(glc1,glc2))>0 And Len(DelDuplicateItems(glc1))<>Len(Intersection(glc1,glc2)) And Len(DelDuplicateItems(glc2))=Len(Intersection(glc1,glc2)) Then
       HasError=4
     ElseIf StrComp(glcnot1,glcnot2,vbTextCompare)=0 And Len(Intersection(glc1,glc2))>0 And (Len(DelDuplicateItems(glc1))<>Len(Intersection(glc1,glc2)) Or Len(DelDuplicateItems(glc2))<>Len(Intersection(glc1,glc2))) Then
       HasError=1
     End If
End Function

Function GetRuleCode(glcnot1,glc1,glcnot2,glc2)
	GetRuleCode=1
	If StrComp(glcnot1,glcnot2,vbTextCompare)=0 And Len(DelDuplicateItems(glc1))=Len(Intersection(glc1,glc2)) And Len(DelDuplicateItems(glc2))=Len(Intersection(glc1,glc2)) Then
		GetRuleCode=0
	ElseIf StrComp(glcnot1,glcnot2,vbTextCompare)=0 And StrComp(glcnot1,"N",vbTextCompare)=0 And Len(Intersection(glc1,glc2))<=0 And StrComp(glc1,"-1",vbTextCompare)<>0 And StrComp(glc2,"-1",vbTextCompare)<>0  Then
		GetRuleCode=-1
	ElseIf  StrComp(glcnot1,glcnot2,vbTextCompare)<>0 And Len(DelDuplicateItems(glc1))=Len(Intersection(glc1,glc2)) And Len(DelDuplicateItems(glc2))=Len(Intersection(glc1,glc2)) Then
		GetRuleCode=-1
	ElseIf  StrComp(glcnot1,glcnot2,vbTextCompare)<>0 And StrComp(glcnot1,"N",vbTextCompare)=0 And Len(DelDuplicateItems(glc1))=Len(Intersection(glc1,glc2)) And Len(DelDuplicateItems(glc2))<>Len(Intersection(glc1,glc2)) Then
		GetRuleCode=-1
	ElseIf  StrComp(glcnot1,glcnot2,vbTextCompare)<>0 And StrComp(glcnot2,"N",vbTextCompare)=0 And Len(DelDuplicateItems(glc1))<>Len(Intersection(glc1,glc2)) And Len(DelDuplicateItems(glc2))=Len(Intersection(glc1,glc2)) Then
		GetRuleCode=-1		
	End If
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