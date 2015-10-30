VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOcena 
      Caption         =   "Izracunaj"
      Height          =   495
      Left            =   6000
      TabIndex        =   31
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtOcena 
      Height          =   285
      Left            =   6000
      TabIndex        =   29
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtBodovi 
      Height          =   285
      Left            =   6000
      TabIndex        =   27
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtRazlikaSekunde 
      Height          =   285
      Left            =   4200
      TabIndex        =   25
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtRazlikaMinuti 
      Height          =   285
      Left            =   4200
      TabIndex        =   23
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtRazlikaSati 
      Height          =   285
      Left            =   4200
      TabIndex        =   21
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton btnIzracunaj 
      Caption         =   "Izracunaj"
      Height          =   495
      Left            =   4200
      TabIndex        =   20
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtRazlika 
      Height          =   285
      Left            =   4200
      TabIndex        =   18
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton btnDatum2 
      Caption         =   "Prikazi datum 2"
      Height          =   495
      Left            =   2160
      TabIndex        =   17
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton btnDatum1 
      Caption         =   "Prikazi datum 1"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtDatum2 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtGodina2 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtMesec2 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtDan2 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtDatum1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtGodina1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtMesec1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtDan1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Ocena"
      Height          =   255
      Left            =   6000
      TabIndex        =   30
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Bodovi"
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Razlika u sekundama"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Razlika u minutima"
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Razlika u satima"
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Razlika u danima"
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Datum"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Godina"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Mesec"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Dan"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Datum"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Godina"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Mesec"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Dan"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDatum1_Click()
    Dim Dan As Integer, Mesec As Integer, Godina As Integer
    Dim Datum As Date
    Dim Rezultat As String

    Dan = txtDan1.Text
    Mesec = txtMesec1.Text
    Godina = txtGodina1.Text
    If Dan < 1 Then
        Dan = 1
    End If
    If Mesec < 1 Then
        Mesec = 1
    End If
    If Godina < 1900 Then
        Godina = 1900
    End If
    
    Rezultat = DateSerial(Godina, Mesec, Dan)

    txtDatum1.Text = Rezultat
End Sub

Private Sub btnDatum2_Click()

    Dim Dan As Integer, Mesec As Integer, Godina As Integer
    Dim Datum As Date
    Dim Rezultat As String

    Dan = txtDan2.Text
    Mesec = txtMesec2.Text
    Godina = txtGodina2.Text
    If Dan < 1 Then
        Dan = 1
    End If
    If Mesec < 1 Then
        Mesec = 1
    End If
    If Godina < 1900 Then
        Godina = 1900
    End If
    
    Rezultat = DateSerial(Godina, Mesec, Dan)
    
    txtDatum2.Text = Rezultat
End Sub

Private Sub btnIzracunaj_Click()
    Dim dat1 As String, dat2 As String
    dat1 = txtDatum1.Text
    dat2 = txtDatum2.Text
    txtRazlika.Text = DateDiff("d", dat1, dat2)
    txtRazlikaSati.Text = DateDiff("h", dat1, dat2)
    txtRazlikaMinuti.Text = DateDiff("N", dat1, dat2)
    txtRazlikaSekunde.Text = DateDiff("s", dat1, dat2)
End Sub

Private Sub btnOcena_Click()
    Dim Bodovi As Integer, Ocena As Integer
    
    Bodovi = txtBodovi.Text
    
    If Bodovi > 90 Then
        Ocena = 10
    ElseIf Bodovi > 80 Then
        Ocena = 9
    ElseIf Bodovi > 70 Then
        Ocena = 8
    ElseIf Bodovi > 60 Then
        Ocena = 7
    ElseIf Bodovi >= 50 Then
        Ocena = 6
    Else
        Ocena = 5
    End If
    txtOcena.Text = Ocena
    
End Sub
