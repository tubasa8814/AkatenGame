VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ゲーム1 
   Caption         =   "赤点回避シミュレーター"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "ゲーム1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ゲーム1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '〇モジュール変数
    '＃＃＃＃＃パラメータ＃＃＃＃＃
    Dim Yaruki As Integer 'やる気パラメータ
    Dim Kokugo As Integer '国語パラメータ
    Dim Sugaku As Integer '数学パラメータ
    Dim Eigo As Integer '英語パラメータ
    Dim Neru As Integer '寝る確率
    '＃＃＃＃＃値保存＃＃＃＃＃
    Dim Hi As Integer '日
    Dim Zikan As Integer '時間
    Dim Skyoka As Integer '教科値
    Dim TBen As Boolean '2時間勉強
    Dim x As Integer
    Dim y As Integer
    '＃＃＃＃＃文字保存＃＃＃＃＃
    Dim Youbi(7) As String '曜日
    Dim Hzikan(14) As String '平日時間表示
    Dim Kzikan(10) As String '休日時間表示
    Dim Kyoka(5) As String '教科
    '〇グローバル変数
    Public Kamoku As Integer '科目保存
    Public Ktai As Boolean 'スマホ使用
    
Private Sub UserForm_Initialize()
    '＃＃＃＃＃ゲーム説明の表示＃＃＃＃＃
    説明.Show
    '＃＃＃＃＃文字列の保存＃＃＃＃＃
    Youbi(1) = "月曜日": Youbi(2) = "火曜日": Youbi(3) = "水曜日": Youbi(4) = "木曜日": Youbi(5) = "金曜日": Youbi(6) = "土曜日": Youbi(7) = "日曜日"
    Hzikan(1) = "1時限目": Hzikan(2) = "2時限目": Hzikan(3) = "3時限目": Hzikan(4) = "4時限目": Hzikan(5) = "5時限目": Hzikan(6) = "6時限目": _
        Hzikan(7) = "17:00": Hzikan(8) = "18:00": Hzikan(9) = "19:00": Hzikan(10) = "20:00": Hzikan(11) = "21:00": Hzikan(12) = "22:00": Hzikan(13) = "24:00": Hzikan(14) = "02:00"
    Kzikan(1) = "8:00": Kzikan(2) = "10:00": Kzikan(3) = "昼食タイム！": Kzikan(4) = "15:00": Kzikan(5) = "17:00": _
        Kzikan(6) = "19:00": Kzikan(7) = "21:00": Kzikan(8) = "22:00": Kzikan(9) = "24:00": Kzikan(10) = "02:00"
    Kyoka(1) = "国語": Kyoka(2) = "数学": Kyoka(3) = "英語": Kyoka(4) = "科学": Kyoka(5) = "日本史"
    '＃＃＃＃＃初期値の保存＃＃＃＃＃
    '日付の保存
    Hi = 1
    通知1.Caption = "今日は" & Hi & "日目、" & Youbi(Hi Mod 8) & "です"
    '時間の保存
    Zikan = 1
    通知2.Caption = "今は" & Hzikan(Zikan) & "です"
    '教科の保存
    Call 教科
    'パラメータ表示の保存
    やる気.Caption = "やる気 ☆☆☆☆☆☆☆☆☆☆"
    国語.Caption = "国語 ☆☆☆☆☆☆☆☆☆☆"
    数学.Caption = "数学 ☆☆☆☆☆☆☆☆☆☆"
    英語.Caption = "英語 ☆☆☆☆☆☆☆☆☆☆"
    '確率初期値
    Neru = 20 '寝る確率（初期値20）
End Sub

Private Sub 次へ_Click()
    '本プログラム
    If Hi < 6 Then '平日
        Zikan = Zikan + 1
        If Zikan < 7 Then '学校内
            TBen = False
            Call 学校学習
            Call 教科
            通知2.Caption = "今は" & Hzikan(Zikan) & "です"
        ElseIf Zikan < 11 Then '（常）家勉
            TBen = False
            Call 家庭学習
            通知3.Caption = ""
            通知2.Caption = "今は" & Hzikan(Zikan) & "です"
            If Zikan = 10 Then '睡眠選択準備
                Call 睡眠選択
            End If
        Else '（特）家勉
            If Zikan = 10 + y Then '寝る時間
                Call 日付進行
            End If
            TBen = True
            Call 家庭学習
            通知2.Caption = "今は" & Hzikan(Zikan) & "です"
        End If
    ElseIf Hi < 8 Then '休日
        Zikan = Zikan + 1
        通知3.Caption = ""
        If Zikan < 7 Then '（常）家勉
            TBen = True
            Call 家庭学習
            通知2.Caption = "今は" & Kzikan(Zikan) & "です"
            If Zikan = 6 Then
                Call 睡眠選択
            End If
        Else '（特）家勉
            If Zikan = 6 + y Then '寝る時間
                Call 日付進行
            End If
            TBen = True
            Call 家庭学習
            通知2.Caption = "今は" & Kzikan(Zikan) & "です"
        End If
    ElseIf Hi = 8 Then 'テスト当日
        Call 日付進行
    End If
End Sub

Private Sub 学校学習()
    '変数宣言
    Dim Sumaho As Integer 'スマホ使用保存
    Dim KNeru As Integer '寝る確率結果
    Dim Naisyoku As Integer '内職ばれる
    Dim Gakusen As Integer '学習選択保存
    '本プログラム
    For x = 1 To 4 '学習選択
        If Me.Controls("勉強" & x).Value = True Then
            Gakusen = x
        End If
    Next x
    If Me.Controls("スマホ").Value = True Then 'スマホ使用確認
        Ktai = True
    End If
    If Ktai = True Then 'スマホを使う
        Randomize 'スマホがばれる確率（初期値2分の1）
        Sumaho = Int((2 - 1 + 1) * Rnd + 1)
        If Sumaho = 1 Then 'スマホがばれない
            Randomize '寝る確率（初期値20）
            KNeru = Int((100 - 1 + 1) * Rnd + 1)
            If KNeru > Neru Then '寝ない
                Yaruki = Yaruki + 10
                If Gakusen = Skyoka Or Gakusen = 4 Then '時間割と選択科目が同様の場合
                    Select Case Gakusen
                        Case "1"
                            Kokugo = Kokugo + 10
                        Case "2"
                            Sugaku = Sugaku + 10
                        Case "3"
                            Eigo = Eigo + 10
                        Case "4"
                            Yaruki = Yaruki + 5
                    End Select
                Else '時間割と選択科目が同様でない場合
                    Randomize '内職確率（初期値2分の1）
                    Naisyoku = Int((2 - 1 + 1) * Rnd + 1)
                    If Naisyoku = 1 Then '内職がばれない
                        Select Case Gakusen
                            Case "1"
                                Kokugo = Kokugo + 10
                            Case "2"
                                Sugaku = Sugaku + 10
                            Case "3"
                                Eigo = Eigo + 10
                            Case "4"
                                Yaruki = Yaruki + 5
                        End Select
                    Else '内職がばれる
                        Yaruki = Yaruki - 20
                    End If
                End If
            End If
        Else  'スマホがばれる
            Yaruki = Yaruki - 35
        End If
    Else 'スマホを使わない
        Randomize '寝る確率（初期値20）
        KNeru = Int((100 - 1 + 1) * Rnd + 1)
        If KNeru > Neru Then '寝ない
            If Gakusen = Skyoka Or Gakusen = 4 Then '時間割と選択科目が同様の場合
                Select Case Gakusen
                    Case "1"
                        Kokugo = Kokugo + 10
                    Case "2"
                        Sugaku = Sugaku + 10
                    Case "3"
                        Eigo = Eigo + 10
                    Case "4"
                        Yaruki = Yaruki + 5
                End Select
            Else '時間割と選択科目が同様でない場合
                Randomize '内職確率（初期値2分の1）
                Naisyoku = Int((2 - 1 + 1) * Rnd + 1)
                If Naisyoku = 1 Then '内職がばれない
                    Select Case Gakusen
                        Case "1"
                            Kokugo = Kokugo + 10
                        Case "2"
                            Sugaku = Sugaku + 10
                        Case "3"
                            Eigo = Eigo + 10
                        Case "4"
                            Yaruki = Yaruki + 5
                    End Select
                Else '内職がばれる
                    Yaruki = Yaruki - 20
                End If
            End If
        End If
    End If
    Label1.Caption = Yaruki
    Label2.Caption = Kokugo
    Label3.Caption = Sugaku
    Label4.Caption = Eigo
End Sub

Private Sub 家庭学習()
    '変数宣言
    Dim KNeru As Integer '寝る確率結果
    Dim Gakusen As Integer '学習選択保存
    '本プログラム
    For x = 1 To 4 '学習選択
        If Me.Controls("勉強" & x).Value = True Then
            Gakusen = x
        End If
    Next x
    If Me.Controls("スマホ").Value = True Then 'スマホ使用確認
        Ktai = True
    End If
    If Ktai = True Then 'スマホを使う
        Randomize '寝る確率（初期値20）
        KNeru = Int((100 - 1 + 1) * Rnd + 1)
        If KNeru > Neru Then '寝ない
            Yaruki = Yaruki + 10
            Select Case Gakusen
                Case "1"
                    Kokugo = Kokugo + 10
                    If TBen = True Then
                        Kokugo = Kokugo + 10
                    End If
                Case "2"
                    Sugaku = Sugaku + 10
                    If TBen = True Then
                        Sugaku = Sugaku + 10
                    End If
                Case "3"
                    Eigo = Eigo + 10
                    If TBen = True Then
                        Eigo = Eigo + 10
                    End If
                Case "4"
                    Yaruki = Yaruki + 5
            End Select
        Else '寝る
            Yaruki = Yaruki - 25
        End If
    Else 'スマホを使わない
        Select Case Gakusen
            Case "1"
                Kokugo = Kokugo + 10
                If TBen = True Then
                    Kokugo = Kokugo + 10
                End If
            Case "2"
                Sugaku = Sugaku + 10
                If TBen = True Then
                    Sugaku = Sugaku + 10
                End If
            Case "3"
                Eigo = Eigo + 10
                If TBen = True Then
                    Eigo = Eigo + 10
                End If
            Case "4"
                Yaruki = Yaruki + 5
        End Select
    End If
    Label1.Caption = Yaruki
    Label2.Caption = Kokugo
    Label3.Caption = Sugaku
    Label4.Caption = Eigo
End Sub

Private Sub 睡眠選択()
    For x = 1 To 4
        If Me.Controls("睡眠" & x).Value = True Then
            y = x
        End If
    Next x
    Select Case y
    Case "1"
        If Neru > 0 Then
            Neru = Neru - 5
        End If
    Case "2"
    Case "3"
        If Neru < 100 Then
            Neru = Neru + 5
        End If
    Case "4"
        If Neru < 95 Then
            Neru = Neru + 10
        End If
    End Select
End Sub

Private Sub 日付進行()
    Hi = Hi + 1
    If Hi < 8 Then '月曜日～日曜日
       通知1.Caption = "今日は" & Hi & "日目、" & Youbi(Hi Mod 8) & "です"
    ElseIf Hi = 8 Then 'テスト当日
        通知1.Caption = "今日は" & "テストです！"
    ElseIf Hi > 8 Then '結果発表
        Unload Me
        結果.Show
    End If
    Zikan = 1
    If Hi < 8 Then
        Call 教科
    End If
End Sub

Private Sub 教科()
    Randomize
    Skyoka = Int((5 - 1 + 1) * Rnd + 1)
    通知3.Caption = Kyoka(Skyoka)
End Sub

Private Sub 終了_Click()
    確認.Show
End Sub

Private Sub 寝る時間_Click()

End Sub
