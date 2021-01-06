VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ウィンドウ3 
   Caption         =   "赤点回避シミュレーター[結果発表]"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5940
   OleObjectBlob   =   "ウィンドウ3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ウィンドウ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Label1.Caption = Kokugo & "点"
    Label2.Caption = Sugaku & "点"
    Label3.Caption = Eigo & "点"
End Sub

Private Sub 終了_Click()
    Unload Me
    メニュー1.Show
End Sub
