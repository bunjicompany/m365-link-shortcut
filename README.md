# M365リンクをアイコン化

SharePoint、Teams、OneDrive の共有リンクを、ローカルのショートカットアイコンとして保存する Windows 用ツールです。

## リリース情報

- バージョン: 1.0.0
- 更新日: 2026-07-04
- 配布ファイル: `dist\M365LinkShortcut.exe`
- SHA-256: `1C4FF9E056FE129FD4014C729AFCE7047E730A9AF641150CBC36FC46FE33A8FC`
- ライセンス: [LICENSE.txt](LICENSE.txt)

## ビルド

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Build-Exe.ps1
```

ビルド成果物は `dist` フォルダに作成されます。

```text
dist\M365LinkShortcut.exe
```

WebView2 関連DLLは exe に埋め込んでいます。別PCへ配布する場合は、`dist\M365LinkShortcut.exe` だけをコピーしてください。

## 開発メモ

ソースコードは `src` フォルダに分割しています。主な役割は以下です。

- `src\Program.cs`: 起動処理、右クリックメニュー登録、クリップボードからのショートカット作成
- `src\LinkParser.cs`: URL判定、クリップボード解析、リンク名候補の取得
- `src\TeamsNameResolver.cs`: Teams会議・チャット名の取得と整形
- `src\BrowserNameResolver.cs`: WebView2によるSharePoint / OneDrive名取得
- `src\ShortcutWriter.cs`: アイコン判定、`.url` / `.lnk` 作成、ファイル名処理

`Build-Exe.ps1` を実行すると、まず `tests\Tests.cs` の自前テストをビルドして実行します。テストが失敗した場合、配布用 exe は更新しません。

テスト成功後に `dist\M365LinkShortcut.exe` を作成し、`dist\SHA256SUMS.txt`、README、配布ページのSHA-256を自動更新します。`dist`、`obj`、`backups` は `.gitignore` で除外しています。

## 右クリックメニュー

`dist\M365LinkShortcut.exe` をダブルクリックすると、右クリックメニューの登録と解除を切り替えます。

PowerShell から明示的に実行する場合:

```powershell
.\dist\M365LinkShortcut.exe --install
.\dist\M365LinkShortcut.exe --uninstall
```

登録後、エクスプローラーでフォルダの余白、またはフォルダ自体を右クリックすると「M365リンクをアイコン化」が表示されます。

## デバッグログ

通常はデバッグログを出力しません。Teams の会議名やチャット名の取得を調査したい場合だけ、登録時に `--debug` を付けます。

```powershell
.\dist\M365LinkShortcut.exe --install --debug
```

この状態で右クリックメニューから実行すると、exe と同じフォルダに `teams-title-debug-*.log` を出力します。通常運用へ戻す場合は、`--debug` なしで登録し直します。

```powershell
.\dist\M365LinkShortcut.exe --install
```

## 使い方

1. SharePoint、Teams、OneDrive でリンクをコピーします。
2. ショートカットを作成したいフォルダをエクスプローラーで開きます。
3. フォルダの余白、またはフォルダ自体を右クリックします。
4. 「M365リンクをアイコン化」を選択します。

ファイルリンクは `.url`、フォルダリンクはフォルダアイコン付きの `.lnk` として作成します。Excel、Word、PowerPoint、Visio は、対応する Office プロトコルが登録されている環境ではデスクトップアプリで開くリンクに変換します。未登録の場合は通常の https リンクのまま作成します。

## Teamsリンクのショートカット名

Teams の会議リンク、チャットリンク、チャネル投稿リンクもショートカット化できます。Teamsリンクの場合は、裏側で Teams アプリを起動して、ウィンドウタイトルや画面上の要素から会議名、チャット名、投稿者、投稿日時を取得します。

ショートカット名はおおむね以下の形式になります。

| リンクの種類 | ショートカット名の例 |
|---|---|
| Teams会議 | `Teams会議 【サンプル定例会】` |
| 会議名が取得できず、日時だけ取得できたTeams会議 | `Teams会議 【7月3日(金)16時00分-17時00分】` |
| チャット全体 | `Teamsチャット 【田中 太郎 Taro Tanaka】` |
| チャネル全体 | `Teamsチャット 【サンプルチーム - 一般】` |
| 個別投稿 | `Teamsチャット 【開発チーム_レビュー用チャット】 2026年7月3日(金)19時16分 - 田中 太郎 Taro Tanaka` |
| チャネルの個別投稿 | `Teamsチャット 【サンプルチーム - 一般】 2026年7月3日(金)19時16分 - 田中 太郎 Taro Tanaka` |

Teams会議は、名前取得後に起動した会議ウィンドウを閉じます。Teamsチャットは、作成後も Teams アプリを閉じません。名前が取得できない場合は、手入力ダイアログを表示します。

## ファイル名の取得

ファイル名は WebView2 でリンク先を非表示で開き、ページタイトル、URL、画面上のテキストから拡張子付きの候補を探して取得します。

WebView2 は通常は画面外で非表示実行し、ユーザーデータは `%LOCALAPPDATA%\M365LinkShortcut\WebView2` に保存します。Microsoft 365 のログイン画面が必要な場合だけ WebView2 を表示し、ログイン後に名前を取得できたら自動で閉じます。埋め込み済みの `WebView2Loader.dll` は初回起動時に `%LOCALAPPDATA%\M365LinkShortcut\Embedded` へ展開します。Windows/Office の組織アカウントでサインイン済みの場合は、OS のプライマリアカウントを使った SSO を試します。

ログイン画面やサインインが必要な状態を検知した場合は、ログイン操作用の WebView2 を表示します。ログイン後もファイル名が取得できない場合だけ、名前入力ダイアログを表示します。SharePoint / OneDrive の名前取得では、クリップボードやURLからの推定名、文字化け修復辞書、Shift-JIS 逆変換による補正は使いません。Teams の会議名・チャット名では、Teams アプリから取得した候補を整えるために、最小限の文字化け補正を行う場合があります。

## アイコン

ショートカット名の拡張子を優先して、アイコンと起動アプリを判定します。名前が取得できない場合や拡張子がない場合のみ、SharePoint URL の種類 (`/:x:/`、`/:w:/`、`/:p:/`、`/:t:/`、`/:i:/`、`/:f:/`) から推定します。

テキスト系や画像系のアイコンは Windows の関連付け情報を優先します。関連付けアイコンが取得できない場合でも、`.url` で解釈が不安定になりやすい負数の `IconIndex` は使わず、正の index のフォールバックアイコンを使います。
