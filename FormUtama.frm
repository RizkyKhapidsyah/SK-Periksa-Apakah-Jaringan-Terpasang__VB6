VERSION 5.00
Begin VB.Form Form_Utama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memeriksa Apakah Jaringan Terpasang"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5010
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Periksa 
      Caption         =   "PERIKSA"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Form_Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Periksa_Click()
    If ApakahJaringanTerpasang = True Then
        MsgBox "Jaringan Terpasang!", vbOKOnly, "Info"
    Else
        MsgBox "Jaringan Belum Terpasang!", vbOKOnly, "Info"
    End If
End Sub
