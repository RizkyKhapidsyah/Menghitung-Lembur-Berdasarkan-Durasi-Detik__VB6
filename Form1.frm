VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung Lembur Berdasarkan Durasi Detik"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2640
      Top             =   4680
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4440
      Top             =   4800
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdMulai 
      Caption         =   "Mulai"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtTampungDetik 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox txtDurasiKedua 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtDurasiPertama 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox txtTglEsok 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox txtTglSistem 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtBesarUang 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtTotalDetik 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtDurasiLembur 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtJamSistem 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtAwalLembur 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtTglMulai 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Tampung Detik"
      Height          =   195
      Left            =   360
      TabIndex        =   23
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Durasi Kedua"
      Height          =   195
      Left            =   360
      TabIndex        =   22
      Top             =   3840
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Durasi Pertama"
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Esok"
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal System"
      Height          =   195
      Left            =   360
      TabIndex        =   19
      Top             =   2760
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Besar Uang"
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Detik"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   2085
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Durasi Lembur"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   1680
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Jam System"
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Awal Lembur"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Mulai"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totaldetik As Long
Dim hh, mm, ss As Integer
Dim tampungdetik As Long

Private Sub cmdMulai_Click()
   Timer1.Enabled = True
   txtAwalLembur.Text = Time
   txtDurasiPertama.Text = Format(CDate("23:59:59") _
    - CDate(txtAwalLembur) + CDate("00:00:01"), _
     "hh:mm:ss")
    'Ditambah satu detik karena belum bulat ke 24:00:00
    'dan angka 24:00:00 tsb tidak valid utk Time
   txtTglMulai.Text = Format(Date, "dd/mm/yyyy")
   txtTglEsok.Text = Format(Date + 1, "dd/mm/yyyy")
   txtDurasiKedua.Text = 0
   txtTampungDetik.Text = 0
   cmdMulai.Enabled = False
   cmdStop.Enabled = True
End Sub

Private Sub cmdStop_Click()
   Timer1.Enabled = False
   Timer2.Enabled = False
   cmdStop.Enabled = False
   cmdMulai.Enabled = True
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   txtTglSistem.Text = Format(Date, "dd/mm/yyyy")
   txtJamSistem.Text = Time
   If txtJamSistem.Text = "00:00:00" Then
      txtDurasiLembur.Text = Format(CDate("23:59:59") - CDate(txtAwalLembur) + CDate("00:00:01"), "hh:mm:ss")
   Else
      txtDurasiLembur.Text = Format((CDate(txtJamSistem.Text) - CDate(txtAwalLembur.Text)), "hh:mm:ss")
   End If
   If CDate(txtAwalLembur) = CDate(txtJamSistem) Then
      txtTampungDetik.Text = 0
   End If
    If CDate(txtTglSistem) = CDate(txtTglEsok) Then
      txtTglEsok.Text = Date + 1
      txtTotalDetik.Text = Format(totaldetik + 1, _
                           "0,0")
      txtTampungDetik.Text = totaldetik + 1
      Timer1.Enabled = False
      Timer2.Enabled = True
   End If
  
   hh = Hour(txtDurasiLembur)
   mm = Minute(txtDurasiLembur)
   ss = Second(txtDurasiLembur)
   totaldetik = hh * 3600 + mm * 60 + ss
   txtTotalDetik.Text = Format(totaldetik, "0,0")
   txtBesarUang.Text = Format(txtTotalDetik * 100, _
                       "0,0")
   txtDurasiKedua.Text = Format(Val(txtTotalDetik) - _
                         Val(txtTampungDetik), "0,0")
End Sub

Private Sub Timer2_Timer()
   On Error Resume Next
   txtTglSistem.Text = Format(Date, "dd/mm/yyyy")
   txtJamSistem.Text = Time
   If txtJamSistem.Text = "00:00:00" Then
      txtDurasiLembur.Text = Format(CDate("23:59:59") - CDate(txtAwalLembur) + CDate("00:00:01"), "hh:mm:ss")
   Else
      txtDurasiLembur.Text = Format(CDate(txtJamSistem.Text) + CDate(txtDurasiPertama.Text) - CDate("00:00:00"), "hh:mm:ss")
   End If
   
   If CDate(txtAwalLembur) = CDate(txtJamSistem) Then
      txtTampungDetik.Text = 0
   End If
   
   If CDate(txtTglSistem) = CDate(txtTglEsok) Then
      txtTglEsok.Text = Date + 1
      txtTotalDetik.Text = Format(totaldetik, "0,0")
      txtTampungDetik.Text = totaldetik + 1
      Timer2.Enabled = False
      Timer1.Enabled = True
   End If
   
   hh = Hour(txtDurasiLembur)
   mm = Minute(txtDurasiLembur)
   ss = Second(txtDurasiLembur)
   totaldetik = hh * 3600 + mm * 60 + ss
   txtTotalDetik.Text = Format(totaldetik, "0,0")
   txtBesrUang.Text = Format(txtTotalDetik * 100, "0,0")
   txtDurasiKedua.Text = Format(Val(txtTotalDetik) - Val(txtTampungDetik), "0,0")
End Sub


