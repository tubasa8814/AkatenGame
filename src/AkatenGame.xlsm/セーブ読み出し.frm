VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} セーブ読み出し 
   Caption         =   "赤点回避シミュレータ"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "セーブ読み出し.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "セーブ読み出し"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub つづきから_Click()
    Unload Me
    セーブ選択.Show
End Sub

Private Sub はじめから_Click()
    Unload Me
    ゲーム1.Show
End Sub

Private Sub 戻る_Click()
    Unload Me
    スタートメニュー.Show
End Sub
