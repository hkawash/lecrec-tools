[(English)](README-en.md)

- <a href="#compress">音声付きパワーポイントスライドの圧縮</a>
- <a href="#hideaudio">挿入音声の「自動再生＆アイコン隠す」を一括設定</a>

<a id="compress"></a>

# 音声付きパワーポイントスライドの圧縮

PowerPointで録音されるオーディオファイルは「音声」としてはビットレートが高いため、どうしてもサイズが大きくなります．このツールは内部のオーディオデータのビットレートを下げて，pptx/ppsxのファイルサイズを小さくします．「挿入 > オーディオ > オーディオの録音」で作成された音声付きパワーポイントを想定しています。

- 内部に含まれる m4a の音声ファイルのビットレートを64kbpsに圧縮します。
- 画像や動画は圧縮しません（それぞれ「書式/図の形式 > 図の圧縮」や「ファイル > 情報 > メディアの圧縮」などを別途利用できます）。

## インストール

![Windowsでの流れ](fig/flow-win-ja.png)

1. 適当なフォルダ（`lecrec-tools`とします）を作成し、中に`ppt-in`というフォルダを作成します。さらに、以下をダウンロードして置いておきます（右クリック > 名前を付けてリンク先を保存）。
   - Windows: [compress_pptaudio-win.bat](https://github.com/hkawash/lecrec-tools/raw/master/compress_pptaudio-win.bat)
   - macOS/Linux: [compress_pptaudio-mac.sh](https://github.com/hkawash/lecrec-tools/raw/master/compress_pptaudio-mac.sh)
   - 別の簡単な方法として、[このプロジェクト(lecrec-tools)のzip](https://github.com/hkawash/lecrec-tools/archive/master.zip)をダウンロードし、zipを展開してもよいです。
2. ffmpeg を[こちらのサイト](https://ffmpeg.zeranoe.com/builds/)からダウンロードして展開します。
   - Version はリリース版（4.2.2など）でよいです。
   - Architecture でOSを選びます。
   - Linking は Static のままとしておきます。
   - ダウンロードサイトは[ffmpeg の元サイト](https://www.ffmpeg.org/download.html)からたどってもよいです。
3. 展開してできたフォルダをたどり、bin というフォルダの下にある ffmpeg.exe (macOSではffmpeg) を、1で作成（もしくは展開）したフォルダの中に入れておきます。

すでに ffmpeg をインストール済みの場合、2, 3のステップは不要です。環境変数でパスを設定していないならば .bat や .sh 内の `PATH` を編集してパスを通しておきます。面倒な場合は 3のステップのようにファイルをコピーしておきます。

## 使い方

1. 音声付きpptxもしくはppsxファイルを`ppt-in`フォルダに入れておきます。
   - 複数のファイルを入れてもOKです。
   - 念のため pptxやppsxはバックアップを取っておいてください。
2. スクリプトを実行します（<a href="#note1">実行時の注意点</a>も参考にしてください）。
   - Windows: `compress_pptaudio-win.bat` をダブルクリックします。
   - macOS (Linux, Windowsのbash): ターミナルを開いて `compress_pptaudio-mac.sh` を実行します。
3. `ppt-out`フォルダに、圧縮された pptx や ppsx ファイルが出力されます。


<a id="note1"></a>

### 実行時の注意点

#### Windows

- **ダウンロードしてから初めて実行するときは、「WindowsによってPCが保護されました」と表示されるので、「詳細情報」をクリックしてから「実行」を選んでください。**
   ![Windowsでの警告](fig/warning-win-ja.png)
- 黒い画面が開きますが、勝手に閉じるまで待ってください。
- 同時に何回も実行しないでください。

#### macOS (および Linux や Windowsのbash)

1. bash の使えるターミナルを開きます。macOSは Launchpad や Spotlightで、ter.. ぐらいを打ち込むとターミナルを選べます。）
1. スクリプトのあるフォルダへ移動します。macOSで、[書類] の下に `lecrec-tools` というフォルダで展開したのであれば、以下を打ち込んでフォルダを移動します。
    ```
    cd /Users/<ユーザ名>/Documents/lecrec-tools
    ```
    （上のコマンドを打ち込んだあと、最後に return を押します。途中でtabキーを押すとファイル名などが補完できます。`<ユーザ名>`のところには、PCで使っている自分のユーザ名を入れます。）
1. ターミナルで以下を打ち込んで return キーを押し、スクリプトを実行します。
    ```
    $ chmod 755 compress_pptaudio-mac.sh
    $ ./compress_pptaudio-mac.sh
    ```
   （一行目の`chmod`はダウンロード後に一回だけ必要）

### その他

- このスクリプトでは、内部の m4a ファイルを 64kbps にします。これ以外のビットレートにするには、スクリプト内の `BITRATE` を変更してください。
- m4a 以外のオーディオファイルが挿入されていて、これを圧縮対象とするには、スクリプト内の `m4a` のあたりを変更してください。

<a id="hideaudio"></a>

# 挿入音声の「自動再生＆アイコン隠す」を一括設定

Powerpoint の「挿入 > オーディオ > オーディオの録音」で作成し、動画(mp4)でエクスポートする（もしくは音声付きパワーポイントにする）には、各ページで、挿入されたスピーカーアイコンを選択し、以下の設定をする必要があります。これは結構面倒なので、全スライド一括で設定するスクリプトを用意します。こちらは PowerShell を使うので Windows用のみです。（Macでも PowerShell を入れれば動くかもしれません。）

- 「再生 > 開始: 自動」を選択
- 「再生 > スライドショーを実行中にサウンドのアイコンを隠す」にチェック

## インストール

以下を同じフォルダへダウンロードします。二通りの使い方があり、それぞれ別の .bat ファイルになっています（.bat ファイルは使う方だけでOK）。

- [hideaudio.ps1](https://github.com/hkawash/lecrec-tools/raw/master/hideaudio.ps1)：どちらを使う場合でも必要
- [hideaudio-dnd.bat](https://github.com/hkawash/lecrec-tools/raw/master/hideaudio-dnd.bat)：pptxファイルをドラッグアンドドロップする場合
- [hideaudio-folder.bat](https://github.com/hkawash/lecrec-tools/raw/master/hideaudio-folder.bat)：pptxファイルをppt-inフォルダに入れておく場合

## 使い方

（注意）MIT ライセンス (No warranty) です。結果の pptx は通常の方法で編集したものと異なり、壊れる可能性もゼロではありませんのでご了承ください。また、入力の pptx も念のためバックアップをとっておいてください。

### ドラッグアンドドロップを使う場合 (hideaudio-dnd.bat)

   1. ファイルエクスプローラーで、pptx ファイル（複数可）を選択
   2. hideaudio-dnd.bat にドラッグアンドドロップ

### フォルダにいれておく場合 (hideaudio-folder.bat)

   1. ppt-in という名前のフォルダを作成
   2. ppt-in フォルダに pptx ファイルを入れる（複数可）
   3. hideaudio-folder.bat をダブルクリック

「セキュリティ警告」などが出ますが、青い画面は<a href="#note1">こちら</a>を参考に、黒い画面では、「R」を入れて実行してください。
いずれの場合も、ppt-out というフォルダに変換後の pptx ファイルが出力されます。
（毎回「R」を入れるのが面倒な場合は、PowerShell を開いて、ファイルのあるフォルダへ `cd` し、`Unblock-File .\hideaudio.ps1` を実行してください。）

## 仕様

- 録音後に「再生」の設定を変更していないオーディオが対象になります。
- 複数オーディオがある場合は最初に見つかったものの設定が変更されます。
