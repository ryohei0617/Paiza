# VBA ベストプラクティス

VBAコードを作成・レビューする際は、以下の注意点を必ず確認してください。

---

## 頻発エラー TOP3（必ず最初に確認）

### Integer オーバーフロー

`Integer` の上限は **32,767**。行数・件数・金額など少し大きい値でも即クラッシュする。
**`Integer` は使わず、常に `Long` を使う。**

```vb
' 悪い例 — 32,768行目で実行時エラー6（オーバーフロー）
Dim i As Integer
For i = 1 To 100000  ' ← 即死

' 良い例
Dim i As Long
For i = 1 To 100000
```

> ループカウンタ・行番号・配列インデックスはすべて `Long` で統一すること。

---

### 文字リテラルのダブルクォートエスケープ

VBAに `\"` は存在しない。文字列中の `"` は **`""`（二重化）** でエスケープする。

```vb
' 悪い例 — コンパイルエラー
Dim msg As String
msg = "彼は\"はい\"と言った"

' 良い例
msg = "彼は""はい""と言った"
' → 出力: 彼は"はい"と言った

' 長い場合は Chr(34) を使うと読みやすい
msg = "彼は" & Chr(34) & "はい" & Chr(34) & "と言った"
```

---

### Set 漏れ（オブジェクト変数への代入）

オブジェクト型変数への代入で `Set` を忘れると **実行時エラー91**（オブジェクト変数がセットされていません）が発生する。

```vb
' 悪い例 — エラー91
Dim ws As Worksheet
ws = ThisWorkbook.Worksheets("Sheet1")  ' Set がない

' 良い例
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Sheet1")
```

`Set` が必要なオブジェクト型の代表例：

| 型 | 例 |
|----|----|
| `Worksheet` | `Set ws = ThisWorkbook.Worksheets("Sheet1")` |
| `Workbook` | `Set wb = Workbooks.Open(path)` |
| `Range` | `Set rng = ws.Range("A1:C10")` |
| `Collection` | `Set col = New Collection` |
| クラスインスタンス | `Set obj = New MyClass` |

> **プリミティブ型**（`Long`, `String`, `Boolean` など）には `Set` は不要。
> 間違えて付けるとコンパイルエラーになる。

---

## 1. 宣言・スコープ

- **`Option Explicit` を全モジュールの先頭に記述する**  
  未宣言変数によるバグを防ぐ。VBEのツール→オプションで「変数の宣言を強制する」をONにすること。
- 変数は使用する直前で宣言し、スコープを最小限に保つ。
- モジュールレベル変数は必要な場合のみ使用し、`Public` 変数は原則避ける。

## 2. 命名規則

| 対象 | 規則 | 例 |
|------|------|----|
| 変数 | camelCase | `totalCount`, `userName` |
| 定数 | UPPER_SNAKE_CASE | `MAX_ROWS`, `DEFAULT_PATH` |
| プロシージャ | PascalCase + 動詞 | `GetSheetData`, `UpdateRecord` |
| 引数 | camelCase + 接頭辞 `p` | `pFilePath`, `pRowIndex` |
| 型付き変数 | ハンガリアン不要（型は宣言で明示） | — |

## 3. エラーハンドリング

```vb
Sub DoSomething()
    On Error GoTo ErrHandler
    
    ' 処理本体
    
    Exit Sub
ErrHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Err.Clear
End Sub
```

- `On Error Resume Next` は局所的にのみ使用し、直後に `Err.Number` を確認してから `On Error GoTo 0` で解除する。
- エラー番号とメッセージは必ずログまたはメッセージで出力する。

## 4. パフォーマンス

画面更新・自動計算の停止は長い処理の必須パターン：

```vb
Sub FastProcess()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo Cleanup
    
    ' 処理本体
    
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub
```

- セルへのアクセスはループ内で1セルずつではなく、**配列に一括読み込み**してから処理する。
- `Select` / `Activate` は使わず、オブジェクト参照を直接操作する。

```vb
' 悪い例
Sheets("Data").Select
Range("A1").Select
Selection.Value = "test"

' 良い例
Sheets("Data").Range("A1").Value = "test"
```

## 5. オブジェクト参照

- **オブジェクト変数への代入は必ず `Set` を付ける**（漏れると実行時エラー91 → 「頻発エラー TOP3」参照）。
- `Set` で代入したオブジェクトは処理後に `Set obj = Nothing` で解放する。
- ワークブック・シートは名前で参照せず、変数に格納して使う。

```vb
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Sheet1")
' ... 使用 ...
Set ws = Nothing
```

## 6. 型の明示

- 変数には必ず型を宣言する（`As Variant` は意図がある場合のみ）。
- **数値には必ず `Long` を使い、`Integer` は使わない**（上限32,767でオーバーフロー → 「頻発エラー TOP3」参照）。
- 文字列結合には `&` を使い、`+` は避ける。

## 7. マジックナンバー・マジック文字列の排除

```vb
' 悪い例
If status = 3 Then ...

' 良い例
Const STATUS_COMPLETED As Long = 3
If status = STATUS_COMPLETED Then ...
```

## 8. プロシージャの設計

- 1プロシージャは1つの責務のみ（単一責任の原則）。
- 行数の目安は **50行以内**。超える場合はサブプロシージャに分割。
- `Function` は値を返すだけにし、副作用（シート操作など）を含めない。

## 9. シート・ブック操作の安全対策

```vb
' シートの存在確認
Function SheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
```

- ファイルパスは `Dir()` で存在確認してから開く。
- `Workbooks.Open` 後は必ず該当ブックの変数を保持する。

## 10. コメント・ドキュメント

- プロシージャの先頭に目的・引数・戻り値を1〜3行で記述する。
- WHYが自明でない箇所のみコメントを書く（WHATはコード自体が示す）。
- 「修正履歴コメント」はGitのコミットメッセージに任せ、コード内には残さない。

---

これらを踏まえてVBAコードを作成・レビューしてください。
