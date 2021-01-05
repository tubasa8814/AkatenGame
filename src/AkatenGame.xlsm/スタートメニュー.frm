VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} スタートメニュー 
   Caption         =   "赤点回避シミュレータ[スタートメニュー]"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "スタートメニュー.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "スタートメニュー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub スタート_Click()
    Unload スタートメニュー
    セーブ読み出し.Show
End Sub

Private Sub 終了_Click()
  確認.Show
End Sub
