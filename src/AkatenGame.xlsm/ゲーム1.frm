VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ゲーム1 
   Caption         =   "赤点回避シミュレーター"
   ClientHeight    =   12000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19200
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
    '＃＃＃＃＃値保存＃＃＃＃＃
    Dim Hi As Integer '日
    Dim Zikan As Integer '時間
    Dim TBen As Boolean '2時間勉強
    Dim Skyoka As Integer '教科値
    Dim X As Integer
    Dim Y As Integer
    '＃＃＃＃＃文字保存＃＃＃＃＃
    Dim Youbi(7) As String '曜日
    Dim Hzikan(14) As String '平日時間表示
    Dim Kzikan(10) As String '休日時間表示
    Dim Kyoka(5) As String '教科

Private Sub UserForm_Initialize()
    '＃＃＃＃＃ゲーム説明の表示＃＃＃＃＃
    ウィンドウ2.Show
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
    通知2.Caption = "次は" & Hzikan(Zikan) & "です"
    '教科の保存
    Call 教科
    '確率初期値
    Neru = 20 '寝る確率（初期値20）
    'コメント・名前
    コメント.Caption = "「今日から一週間で赤点を回避してみせる！赤点を回避する必要があるのは国語と数学、そして英語であとは捨て教科だ！」"
    '背景
    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室1.jpg")
End Sub

Private Sub セーブ_Click()

End Sub

Private Sub ロード_Click()

End Sub

Private Sub 次へ_Click()
    '本プログラム
    If Hi < 6 Then '平日
        Zikan = Zikan + 1
        If Zikan < 7 Then '学校内
            TBen = False
            Call 学校学習
            Call 教科
            通知2.Caption = "次は" & Hzikan(Zikan) & "です"
        ElseIf Zikan < 11 Then '（常）家勉
            TBen = False
            Call 家庭学習
            通知3.Caption = ""
            通知2.Caption = "次は" & Hzikan(Zikan) & "です"
            If Zikan = 10 Then '睡眠選択準備
                ゲーム3.Show
            End If
        Else '（特）家勉
            If Zikan = 10 + Nzikan Then '寝る時間
                Call 日付進行
            End If
            TBen = True
            Call 家庭学習
            通知2.Caption = "次は" & Hzikan(Zikan) & "です"
        End If
    ElseIf Hi < 8 Then '休日
        Zikan = Zikan + 1
        通知3.Caption = ""
        If Zikan < 7 Then '（常）家勉
            TBen = True
            Call 家庭学習
            通知2.Caption = "次は" & Kzikan(Zikan) & "です"
            If Zikan = 6 Then
                ゲーム3.Show
            End If
        Else '（特）家勉
            If Zikan = 6 + Nzikan Then '寝る時間
                Call 日付進行
            End If
            TBen = True
            Call 家庭学習
            通知2.Caption = "次は" & Kzikan(Zikan) & "です"
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
    '本プログラム
    ゲーム2.Show
    If Ktai = True Then 'スマホを使う
        Randomize 'スマホがばれる確率（初期値2分の1）
        Sumaho = Int((2 - 1 + 1) * Rnd + 1)
        If Sumaho = 1 Then 'スマホがばれない
            Randomize '寝る確率（初期値20）
            KNeru = Int((100 - 1 + 1) * Rnd + 1)
            If KNeru > Neru Then '寝ない
                Yaruki = Yaruki + 10
                If Kamoku = Skyoka Or Kamoku = 6 Then '時間割と選択科目が同様の場合
                    Select Case Kamoku
                        Case "1"
                            Kokugo = Kokugo + 10
                        Case "2"
                            Sugaku = Sugaku + 10
                        Case "3"
                            Eigo = Eigo + 10
                        Case "6"
                            Yaruki = Yaruki + 5
                    End Select
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室1.jpg")
                    コメント.Caption = "「この時間は計画通りに勉強できた！スマホもばれなくて勉強って楽しいな！」"
                Else '時間割と選択科目が同様でない場合
                    Randomize '内職確率（初期値2分の1）
                    Naisyoku = Int((2 - 1 + 1) * Rnd + 1)
                    If Naisyoku = 1 Then '内職がばれない
                        Select Case Kamoku
                            Case "1"
                                Kokugo = Kokugo + 10
                            Case "2"
                                Sugaku = Sugaku + 10
                            Case "3"
                                Eigo = Eigo + 10
                            Case "6"
                                Yaruki = Yaruki + 5
                        End Select
                        背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室4.jpg")
                        コメント.Caption = "「この時間は計画通りに勉強できた！スマホも内職もばれなかった！この調子で勉強を進めていこう！」"
                    Else '内職がばれる
                        Yaruki = Yaruki - 20
                        背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室2.jpg")
                        コメント.Caption = "「しまった、さっきの時間は内職がばれて勉強どころじゃなかった…」"
                    End If
                End If
            Else '寝る
                Yaruki = Yaruki - 20
                背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室3.jpg")
                コメント.Caption = "主人公は寝てしまった"
            End If
        Else  'スマホがばれる
            Yaruki = Yaruki - 35
            背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室2.jpg")
            コメント.Caption = "「最悪だ、スマホが没収されてしまった。新しいのを買わないと…」"
        End If
    Else 'スマホを使わない
        Randomize '寝る確率（初期値20）
        KNeru = Int((100 - 1 + 1) * Rnd + 1)
        If KNeru > Neru Then '寝ない
            If Kamoku = Skyoka Or Kamoku = 6 Then '時間割と選択科目が同様の場合
                Select Case Kamoku
                    Case "1"
                        Kokugo = Kokugo + 10
                    Case "2"
                        Sugaku = Sugaku + 10
                    Case "3"
                        Eigo = Eigo + 10
                    Case "6"
                        Yaruki = Yaruki + 5
                End Select
                背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室1.jpg")
                コメント.Caption = "「この時間は計画通りに勉強できた。やっぱり勉強はまじめにやるべきだ！」"
            Else '時間割と選択科目が同様でない場合
                Randomize '内職確率（初期値2分の1）
                Naisyoku = Int((2 - 1 + 1) * Rnd + 1)
                If Naisyoku = 1 Then '内職がばれない
                    Select Case Kamoku
                        Case "1"
                            Kokugo = Kokugo + 10
                        Case "2"
                            Sugaku = Sugaku + 10
                        Case "3"
                            Eigo = Eigo + 10
                        Case "6"
                            Yaruki = Yaruki + 5
                    End Select
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室4.jpg")
                    コメント.Caption = "「この時間は計画通りに勉強できた。内職もばれなかった！この調子で勉強を進めていこう！」"
                Else '内職がばれる
                    Yaruki = Yaruki - 20
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室2.jpg")
                    コメント.Caption = "「しまった、さっきの時間は内職がばれて勉強どころじゃなかった…」"
                End If
            End If
        Else '寝る
            Yaruki = Yaruki - 20
            背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室3.jpg")
            コメント.Caption = "主人公は寝てしまった"
        End If
    End If
    Label1.Caption = Kokugo
    Label2.Caption = Sugaku
    Label3.Caption = Eigo
    Label4.Caption = Yaruki
    Label11.Caption = Kamoku
    Label9.Caption = Ktai
    Label13.Caption = KNeru
    Label15.Caption = Neru
End Sub

Private Sub 家庭学習()
    '変数宣言
    Dim Sumaho As Integer 'スマホ使用保存
    Dim KNeru As Integer '寝る確率結果
    '本プログラム
    ゲーム2.Show
    If Ktai = True Then 'スマホを使う
        Randomize '時間を無駄にする（初期値20）
        KNeru = Int((100 - 1 + 1) * Rnd + 1)
        If KNeru > Neru Then '時間を無駄にしない
            Yaruki = Yaruki + 10
            Select Case Kamoku
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
                Case "6"
                    Yaruki = Yaruki + 10
            End Select
            Select Case Zikan
                Case Is < 3
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋1-4.jpg")
                Case "3"
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋4-4.jpg")
                Case Is < 6
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室1-4.jpg")
                Case "6"
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋2-4.jpg")
                Case Is > 6
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋3-4.jpg")
            End Select
            コメント.Caption = "「この時間は計画通りに勉強できた！スマホを使って勉強って楽しくて捗るな！」"
        Else '時間を無駄にした
            Yaruki = Yaruki - 25
            Select Case Zikan
                Case Is < 3
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋1-2.jpg")
                Case "3"
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋4-2.jpg")
                Case Is < 6
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室1-2.jpg")
                Case "6"
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋2-2.jpg")
                Case Is > 6
                    背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋3-2.jpg")
            End Select
            コメント.Caption = "「最悪だ、スマホを使ったせいで時間を浪費してしまった。」"
        End If
    Else 'スマホを使わない
        Select Case Kamoku
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
            Case "6"
                Yaruki = Yaruki + 10
        End Select
        Select Case Zikan
            Case Is < 3
                背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋1-1.jpg")
            Case "3"
                背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋4-1.jpg")
            Case Is < 6
                背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\教室1-1.jpg")
            Case "6"
                背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋2-1.jpg")
            Case Is > 6
                背景1.Picture = LoadPicture(ThisWorkbook.Path & "\..\gfx\セット\部屋3-1.jpg")
        End Select
        コメント.Caption = "「この時間は計画通りに勉強できた。やっぱり勉強はまじめにやるべきだ！」"
    End If
    Label1.Caption = Kokugo
    Label2.Caption = Sugaku
    Label3.Caption = Eigo
    Label4.Caption = Yaruki
    Label11.Caption = Kamoku
    Label9.Caption = Ktai
    Label13.Caption = KNeru
    Label15.Caption = Neru
End Sub

Private Sub 日付進行()
    Hi = Hi + 1
    If Hi < 8 Then '月曜日～日曜日
       通知1.Caption = "今日は" & Hi & "日目、" & Youbi(Hi Mod 8) & "です"
    ElseIf Hi = 8 Then 'テスト当日
        通知1.Caption = "今日は" & "テストです！"
    ElseIf Hi > 8 Then '結果発表
        Unload Me
        ウィンドウ3.Show
    End If
    Zikan = 1
    If Hi < 8 Then
        Call 教科
    End If
End Sub

Private Sub 教科()
    '変数宣言

    'プログラム
    Randomize
    Skyoka = Int((5 - 1 + 1) * Rnd + 1)
    通知3.Caption = "次は" & Kyoka(Skyoka)
End Sub

Private Sub 終了_Click()
    ウィンドウ1.Show
End Sub
