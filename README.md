[← 🏠 プロフィールに戻る](https://github.com/paison58223-lang)

# TechnologyNote(Excel)
## 📌 概要
このページは ExcelVBA のツール作成で学んだことをまとめたNoteです。

<details open>
<summary><b>📒 セル操作を発火とするマクロを作る方法</b></summary>


<b>🔧 基本ポイント<br></b>
・ セル操作を行うシートの「シートモジュール」にコードを書く  <br>
・ ※標準モジュールでは動作しないことに注意！<br>
<br>
<b>🔍 手順：イベントを選ぶ</b><br>
下記の表示の場所で動作させたい種類を選ぶ。  <br>
今回はダブルクリックを発火点としたいので  <br>
<b>『BeforeDoubleClick』</b> を選択。<br>

<img width="1913" height="130" alt="image" src="https://github.com/user-attachments/assets/ffdc8203-67d1-494b-af83-17b7c567adc1" />
<br>
<b>⚙️ 実装：発火点の調整方法</b><br>

<pre><code class="language-vb">
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' D1 以外のダブルクリックは無効
    If Not (Target.Row = 1 And Target.Column = 4) Then Exit Sub
End Sub
</code></pre>

<br>
Targetが発火するセルに該当するとTrueが返される。<br>
この際、直接 Target = Range("D1") とすることはできない。<br>


**<b>🐔＜　Range同士比較は不正確・複数セルで壊れる　コケっ！ </b>**<br>


従って該当セルのRow(行)、Column(列)を用いて指定しなければならない。<br>

</details>
<details open>
<summary><b>🕰️ 時間で発動するマクロを組む方法</b></summary>

<b>🔧 基本ポイント</b>
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
<details>
<summary><b>😒 いまさら聞けないVBAの変数宣言あれこれ</b></summary>

  ---

### 🔧 基本ポイント（この記事で分かること）
- `Dim` で宣言した変数が、なぜ他のプロシージャから見えないのか  
- `Call ` ＋ `ByRef` で値を受け渡す方法  
- `Public` で「どこからでも使える変数」を作る方法と、その落とし穴  

### Lv1:Dim
`Dim` による変数宣言は **そのプロシージャーの中だけで有効（ローカル変数）** です。
そのため、別のプロシージャーから同じ名前で参照しても、中身は共有されません。

```vb
Sub Test()

    Dim x As Long
    
    x = 10
        
End Sub

Sub Test2()

    x = x + 1
    
    Debug.Print x

End Sub
```

<img width="291" height="374" alt="image" src="https://github.com/user-attachments/assets/24e0d476-10d7-4f26-bf27-5b3a5260e5c7" /><br>
↑Testで定義したx=10が反映されず、x=0として計算された結果1が出力されている。

例えば繰り返しでi等を使いまわす際等、プロシージャー毎で変数の扱いをリセットしたい場合はDimが有効です。

一応この状態でもTest2にx=10を渡す方法もあります。

```vb
Sub Test()

    Dim x As Long
    
    x = 10
    
    Call Test2(x)
        
End Sub

Sub Test2(ByRef x As Long)

    x = x + 1
    
    Debug.Print x

End Sub
```
<img width="269" height="379" alt="image" src="https://github.com/user-attachments/assets/ef89f354-1bf9-4ffe-8a91-63feb75e6f9b" /><br>
↑2つのプロシージャーを連結し、変数の結果を渡す。

実は恥ずかしながら、変数を渡すにはこの方法しかないと思い込んでいましたが、もっとスマートに変数を受け渡す方法があります。

### Lv2:Public
`Public` で宣言した変数は **モジュール全体で共有されるグローバル変数** になります。

```vb
' 標準モジュールの先頭に書く
Public x As Long

Sub Test()
    
    x = 10
    
    Call Test2
            
End Sub

Sub Test2()

    x = x + 1
    
    Debug.Print x　' → 11 が表示される

End Sub
```
<img width="269" height="358" alt="image" src="https://github.com/user-attachments/assets/215ce7cc-128d-4fca-a69f-ee76b5d7c87e" />
<br>'Public x As Long' は プロシージャの外側（モジュールの先頭） に書く必要があります
<br>じゃあDimなんて使わずに全部Publicにすればいいんじゃない？と思うかもしれませんが、
<br>一度宣言すると、同じ標準モジュール内のすべてのプロシージャで同じ x を共有します

<details>
<summary><b>おまけ：Publicの落とし穴（実務編）</b></summary>
Public は便利ですが、安易に使うと次のような問題を起こします。

❌ ① 値が「勝手に書き換わる」事故が起きる

Public で宣言した変数はすべてのプロシージャから読み書きできるため、
どこか別の処理が意図せず上書きしてしまうことがあります。

```vb
Public cnt As Long

Sub A()
    cnt = 10
End Sub

Sub B()
    cnt = 999   ' ← Aが設定した値が消える
End Sub
```

→ <b>🐔< 大規模マクロになるほどバグ源になるコケ。</b>

❌ ② マクロが“並列で動く処理”と相性最悪（タイマー・イベント系）

OnTime や Worksheet_Change を使う場合は特に危険。

複数のタイマーが同時に走ると、
Public変数を 同時に上書き してしまい、値が破壊されることがある。

Public remain As Long  ' カウントダウンの残り時間

' 2つの違うOnTimeが同じ変数を触ってバグるパターン


→ <b>🐔< この手の事故は“再現が難しい”から最悪コケ。</b>

❌ ③ シートモジュールとの Public は反映されない

Public なのに片方のモジュールでしか有効にならないという罠。

標準モジュール → シートモジュールからも見える

シートモジュール → 標準モジュールからは見えない（見えそうで見えない罠）

→ <b>🐔< “Public にしたのに値が渡らない” という混乱ポイントになるコケ。</b>

❌ ④ バグが起きたとき「どこが原因か追えない」

Public 変数が複数のプロシージャで書き換えられるため、
どこで値が壊れたのか調査が非常に難しい。

Dim なら
「このプロシージャの中でしか使われない」
と分かるので原因箇所を即特定できる。

<b>🐔＜ 結論コケ！ </b>

Public は <b>設定値・最終結果・定数に近い役割</b> のときだけ安全。
処理中の “途中経過” や “カウント用の一時変数” は Public にすると事故るコケ！

</details>
</details>
