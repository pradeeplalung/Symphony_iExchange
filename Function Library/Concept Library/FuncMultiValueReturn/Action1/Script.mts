'**** function return muliple value**********
Function meth()
Dim a(3)
b = "Chandler [USRN: 9999999]"
c1 = split(b,"[USRN: ")
a(0) = c1(0)
a(1) = c1(1)
d = split(a(1),"]")
a(2) = d(0)
meth = a
End Function
'***** assigning return values in array from meth () to b ****
b = meth
'*** uBound will give upper value of the array, in this case '3' ****
For i = 0 to ubound(b)
'	print b(i)
'    msgbox b(0)
'	msgbox b(1)
    k = b(i)
Next

  msgbox b(0) &"::::"&b(1)&"::::"&b(2)
'  b2 = split( b(1),"]")
'  b3 = b2(0)












