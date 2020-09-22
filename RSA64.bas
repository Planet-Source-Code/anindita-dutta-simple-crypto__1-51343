Attribute VB_Name = "RSA64"
Public Key(1 To 3) As Double
Private Function Mult(ByVal x As Double, ByVal pg As Double, ByVal m As Double) As Double
On Error GoTo errorHandler
y = 1
    Do While pg > 0
        Do While (pg / 2) = Int((pg / 2))
            x = nMod((x * x), m)
            pg = pg / 2
        Loop
        y = nMod((x * y), m)
        pg = pg - 1
    Loop
    Mult = y
    Exit Function
errorHandler:
y = 0
End Function
Private Function nMod(x As Double, y As Double) As Double
On Local Error Resume Next
nMod = x - (Int(x / y) * y)
End Function
Public Function enc(ByVal tIp As String, eE As Double, eN As Double) As String
On Local Error Resume Next
Dim encSt As String
    If tIp = "" Then Exit Function
    For I = 1 To Len(tIp)
     encSt = encSt & Mult(CLng(Asc(Mid(tIp, I, 1))), eE, eN) & ";"
    Next I
enc = encSt
End Function
Public Function dec(ByVal tIp As String, dD As Double, dN As Double) As String
On Local Error Resume Next
Dim decSt As String
For z = 1 To Len(tIp)
    ptr = InStr(z, tIp, ";")
    tok = Val(Mid(tIp, z, ptr))
    decSt = decSt + Chr(Mult(tok, dD, dN))
    z = ptr
Next z
dec = decSt
End Function

'       Reference :-
'   http://www.rsasecurity.com/rsalabs/

