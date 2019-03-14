Option Explicit
Dim arr1(9),arr2(9)
arr1(0)="N"
arr2(0)="N"
arr1(1)="14"
arr2(1)="14"

arr1(2)="N"
arr2(2)="N"
arr1(3)="'480','493'"
arr2(3)="'480','493'"

arr1(4)="N"
arr2(4)="N"
arr1(5)="'123'"
arr2(5)="'123'"

arr1(6)="N"
arr2(6)="N"
arr1(7)="-1"
arr2(7)="-1"

arr1(8)="Y"
arr2(8)="Y"
arr1(9)="'100','101','102','103','104','105','310','312','313','314','315','316','317','343','344','345','410','421','422','423','424','425','435','436','437','438','441','444'"
arr2(9)="'100','101','102','103','104','105','310','312','313','314','315','316','317','343','344','345','410','421','422','423','424','425','435','436','437','438','441','444'"
'WScript.Echo "compares(,arr1,arr2):"&Compares("",arr1,arr2)
'WScript.Echo "compares(0,1,arr1,arr2):"&Compares("0,1",arr1,arr2)
'WScript.Echo "compares(2,3,arr1,arr2):"&Compares("2,3",arr1,arr2)
'WScript.Echo "compares(4,5,arr1,arr2):"&Compares("4,5",arr1,arr2)
'WScript.Echo "compares(6,7,arr1,arr2):"&Compares("6,7",arr1,arr2)
'WScript.Echo "compares(8,9,arr1,arr2):"&Compares("8,9",arr1,arr2)
WScript.Echo "Intersection(str1,str2):"&Intersection(arr1(5),arr2(5))
WScript.Echo "len(Intersection(str1,str2)):"&Len(Intersection(arr1(5),arr2(5)))
WScript.Echo "Len(DelDuplicateItems(arr1(5))):"&Len(DelDuplicateItems(arr1(5)))
WScript.Echo "Len(DelDuplicateItems(arr2(5))):"&Len(DelDuplicateItems(arr2(5)))
'WScript.Echo entity
'WScript.Echo "Intersection(str1,str2):"&Len(Intersection(arr1(5),arr2(5)))&","&Len(DelDuplicateItems(arr1(5)))&","&Len(DelDuplicateItems(arr2(5)))
WScript.Quit

Function Compares(str,arr1,arr2)
Dim sec,i,j,count,secnum
sec=Split(str,",",-1,1)
secnum=UBound(sec)
WScript.Echo("UBound(sec):"&secnum)
count=0
If StrComp(str,"",vbTextCompare)=0 Then
 For i=0 To UBound(arr1)
	If StrComp(arr1(i),arr2(i),vbTextCompare)=0 Then
		count=count+1
	End If
 Next
 
Else
	For i=0 To CInt(sec(0))-1
		If StrComp(arr1(i),arr2(i),vbTextCompare)=0 Then
		wscript.Echo("countA"&count)
			count=count+1
		End If
	Next
		
End If

If secnum>-1 Then
	WScript.Echo("count:"&count&"sec(0):"&sec(0))
	If count=CInt(sec(0)) Then
		Compares=0
	Else 
	Compares=-1
	End If
Else
	WScript.Echo("count:"&count&"UBound(arr1):"&UBound(arr1))
	If count-1=UBound(arr1) Then
		Compares=0
	Else 
	Compares=-1
	End If
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
Dim arr,c,i,j,k
arr=Split(str,",",-1,1)
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





































Function ComparesA(str,arr1,arr2)
Dim sec,i,j,count,secnum
sec=Split(str,",",-1,1)
secnum=UBound(sec)
WScript.Echo("UBound(sec):"&secnum)
count=0

For i=0 To UBound(arr1)
 WScript.Echo("1st for: arr1("&i&"):"&arr1(i) &", arr2("&i&"):"&arr2(i))
 If StrComp(arr1(i),arr2(i),vbTextCompare)=0 Then
  count=count+1
	If secnum>-1 Then
		For j=0 To UBound(sec)
			WScript.Echo("count:"&count& vbLf&" 2nd for: i:"&i &", sec("&j&"):"&sec(j))
			If StrComp(i,sec(j),vbTextCompare)=0 Then 
				count=count-1
				secnum=secnum-1 
				WScript.Echo("count:"&count&" secnum:"&secnum)
			End If
		Next
	End If
 End If

Next
WScript.Echo("UBound(arr1):"&UBound(arr1)&" UBound(sec):"&UBound(sec))
If count=UBound(arr1)-UBound(sec) Then
'WScript.Echo("count=UBound(arr1)-UBound(sec):"&count) 
Compares=0
Else 
Compares=-1
End If
End Function