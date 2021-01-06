VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} メニュー1 
   Caption         =   "赤点回避シミュレータ[スタートメニュー]"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "メニュー1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "メニュー1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub スタート_Click()
    Unload Me
    メニュー2.Show
End Sub

Private Sub 終了_Click()
  ウィンドウ1.Show
End Sub
