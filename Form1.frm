VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memisahkan string Berdasarkan Suatu Separator"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   2040
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Anda dapat men-split string "Kucing,Tikus,Belalang,Rusa"
'menjadi beberapa substrings "Kucing", "Tikus", "Belalang" 'dan "Rusa".
'Coding di bawah ini menggunakan fungsi 'split' yang 'hanya tersedia mulai Visual Basic 6.0 ke atas.

Private Sub Command1_Click()
Dim strAnimals As String
Dim iCounter As Integer
Dim arrAnimals() As String
  strAnimals = "Kucing,Tikus,Belalang,Rusa"
  'Ganti "," di bawah dengan separator yang Anda
  'inginkan
  arrAnimals = Split(strAnimals, ",")
  For iCounter = LBound(arrAnimals) To UBound(arrAnimals)
      MsgBox arrAnimals(iCounter)
  Next
End Sub


