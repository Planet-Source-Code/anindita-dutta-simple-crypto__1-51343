VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Programmed By : Anindita Dutta

Private Sub Command1_Click()        ' Encrypt
Text1.Text = enc(Text1.Text, Key(1), Key(3))
End Sub

Private Sub Command2_Click()        ' Decrypt
Text1.Text = dec(Text1.Text, Key(2), Key(3))
End Sub

Private Sub Form_Load()
Key(1) = 35429567    ' public key
Key(2) = 21444671    ' private key
Key(3) = 31393357    ' n=p*q
'                      Factors p = 3613 ;  q = 8689
End Sub

'-------------------------------------------------------------------

' The RSA Algorithm works as follows :-
' take two large primes, p and q, compute
' their product n=pq; n is called the
' modulus. e, another number less than n
' where e and (p-1)*(q-1) have no common
' factors except 1. Choose another number
' d, such as (e*d - 1) is divisible by
' (p-1)*(q-1). The values e and d are the
' public and private exponents, respectively.

