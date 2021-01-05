VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} セーブ選択 
   Caption         =   "赤点回避シミュレータ[セーブ選択]"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "セーブ選択.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "セーブ選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub 選択1_Click()

End Sub

Private Sub 選択2_Click()

End Sub

Private Sub 選択3_Click()

End Sub

Private Sub 戻る_Click()
    Unload Me
    セーブ読み出し.Show
End Sub
