[← 🏠 プロフィールに戻る](https://github.com/paison58223-lang)

# TechnologyNote(Excel)
## 📌 概要
このページは ExcelVBA のツール作成で学んだことをまとめたNoteです。

<details>
<summary><b>📒 セル操作を発火とするマクロを作る方法</b></summary>

---

### 🔧 基本ポイント
- セル操作を行うシートの「シートモジュール」にコードを書く  
- ※標準モジュールでは動作しないことに注意！

<br>

### 🔍 手順：イベントを選ぶ
下記の表示の場所で動作させたい種類を選ぶ。  
今回はダブルクリックを発火点としたいので  
**『BeforeDoubleClick』** を選択。

<img width="1913" height="260" alt="image" src="https://github.com/user-attachments/assets/ffdc8203-67d1-494b-af83-17b7c567adc1" />

### ⚙️ 実装：発火点の調整方法

```vb
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  ' D1 以外のダブルクリックは無効
  If Not (Target.Row = 1 And Target.Column = 4) Then Exit Sub
```
Targetが発火するセルに該当するとTrueが返される。
この際、直接 Target = Range("D1") とすることはできない。


**<b>🐔＜　Range同士比較は不正確・複数セルで壊れる　コケっ！ </b>**


従って該当セルのRow(行)、Column(列)を用いて指定しなければならない。

---
</details>
<details>
<summary><b>🕰️ 時間で発動するマクロを組む方法</b></summary>

---

### 🔧 基本ポイント
- 時間を指定して繰り返し同じ処理を行う。  
- 繰り返した結果で一定条件を満たしたときに終了する処理をいれる。
- ※実際に走らせてみると処理が重すぎて正確な時間で動作させるのが難しい場合があり、注意！

### 1️⃣ 用いる関数

```vb
Application.OnTime Now + TimeValue("00:00:01"), "UpdateCountDownTimer"
```
これは現在から1秒経過したら『UpdateCountDownTimer』というプロシージャーを動かすように設定しています。<br>
アラーム等、1回だけ動作する目的ならこれでOKですが、  
繰り返し動作させたい(カウントダウンタイマーを作成したい)等の場合には<br>
『UpdateCountDownTimer』のプロシージャー内で再び、<br>
Application.OnTime Now + TimeValue("00:00:01"), "UpdateCountDownTimer" <br>
を用いる必要があります。

### 2️⃣ 参考コード(10秒後にMsgBoxを出すマクロ)

```vb
Public TargetTime As Date
Public counting As Boolean
Sub Timestart()
    
    counting = True
    TargetTime = Now + TimeValue("00:00:10") '目標時間を10秒後に設定
    
    Application.OnTime Now + TimeValue("00:00:01"), "UpdateCountDownTimer" '1秒後にUpdateCountDownTimerを動作させる

End Sub

Sub UpdateCountDownTimer()

    Dim currentTime As Date
    
    ' ★ 中断してたら即終了（誤発火防止）
    If counting = False Then Exit Sub

    currentTime = TargetTime - Now '現在時刻から残り時間を算出

    ' ★ 自然終了
    If counting = True And currentTime <= 0 Then '残り時間0以下かつ動作継続中

        counting = False
        MsgBox "10秒経過しました!"
        Exit Sub

    End If

    ' ★ カウント継続
    If currentTime > 0 Then '残り時間がまだある場合
        'Range("C3") = currentTime 'もし特定のセル等に残り時間を表示したければ
        Application.OnTime Now + TimeValue("00:00:01"), "UpdateCountDownTimer" '1秒後に再度このマクロを動作させる
    End If

End Sub

```

### 3️⃣ 注意点
- この参考コードもPCが重いと毎秒動作する判定がでないため、正確な動作を求めるなら動作間隔を広げてみたりしてください。
- 厳密に10秒計測したい場合にはお勧めしません。
- Excel上でカウントダウンを雰囲気で出したい等の用途にご利用ください。

</details>
