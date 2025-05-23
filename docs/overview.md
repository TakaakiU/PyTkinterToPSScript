# OVERVEIW - PyTkinterToPSScript

## はじめに

本資料は、PyTkinterToPSScriptでパッケージ化処理やチェック処理、モード変更、設定変更を機能を説明したもの。

使用しているAdpack.exeの引数指定では、パッケージ化処理のことをパック（Pack）、チェック処理のことをアンパック（Unpack）と定義。

以降でもその表現に合わせている箇所もあり。

## インストーラー

### 制約・前提条件

- Windows 10/11 の64bit版であること

- Cドライブが存在すること

- Microsoft Office 2019以降がインストールされていること

### inno setup 設定ファイル

「開発環境の概要」を参照。

### 処理の流れ

1. セットアップファイルをダブルクリックして実行

2. インストールボタンをクリック
   
   - インストールフォルダー「`C:¥PyTkinterToPSScript`」を作成
   
   - インストールフォルダー配下にE-DataPackgerのリソースを配置
   
   - デスクトップに「`C:¥E-DataPakcager¥exe/PyTkinterToPSScript.exe`」のショートカットを配置

3. 完了

## パッケージ化処理

### 制約・前提条件

- パッケージ化対象のファイルに1.5GB以上のデータがないこと（ファイル数の制限はなし）

- XXXXシステムのデータであること

- [AdPackツール](https://www.ossal.org/salproj/adpack.html)の実行ファイルを個別にダウンロードし配置

### 処理の流れ

1. 入力データのチェック
   
   - 各入力項目のチェック
     
     - 対象フォルダーの入力チェック
       → 未入力の場合はエラー。
     
     - 対象フォルダーの存在チェック
       → 存在しない場合はエラー。
     
     - 作業対象の入力チェック
       → 未入力の場合はエラー。
     
     - 作業日付の論理チェック
       → 日付ではない場合はエラー。
     
     - 部署名・作業者名の入力チェック
       → 未入力の場合はエラー。
     
     - 作業端末名の入力チェック
       → 未入力の場合はエラー。
   
   - 入力データをCSVファイルに保存
     「`C:/PyTkinterToPSScript/input/Packagelist_HeaderValues.csv`」に出力。（すでに存在する場合は上書き保存）
   
   - ファイル数とファイルサイズをチェック
     → 対象フォルダー内のデータが設定ファイルの値以上だった場合にエラー

2. パッケージ化
   
   - 下記の引数を指定しPowerShellスクリプトを実行
     
     - 実行するPowerShellスクリプトは、「`C:/PyTkinterToPSScript/script/AdpackController.ps1`」
     
     - 設定ファイルのハッシュアルゴリズムを読み取る（規定値は、`s256`）
     
     - 入力データは、対象フォルダーの入力値
     
     - 出力データは、対象フォルダーに「`.zip`」を付与したファイル

3. 出力処理
   
   - 下記の引数を指定しPowerShellスクリプトを実行
     
     - 実行するPowerShellスクリプトは、「`C:/PyTkinterToPSScript/script/PrintController.ps1`」
     
     - パック処理後の出力処理を意味する「`PackForm`」を指定
     
     - 出力先は、入力フォルダ―と同じ階層のフォルダーを指定
     
     - E-DataPackerのインストール先「`C:/PyTkinterToPSScript`」
     
     - ヘッダーと本文の位置データファイルと実データファイルを指定

## チェック処理

### 制約・前提条件

- モード変更のパスワード認証によりチェックモードに変更していること

- 同じバージョンのPyTkinterToPSScriptでパッケージ化したデータが対象であること

- ファイルサイズやファイル数の制限はなし（機能としては実装しており設定ファイルで定義可能）

### 処理の流れ

1. 入力データのチェック
   
   パッケージ化処理の入力データのチェックと同様。ファイルサイズ と ファイル数 はチェック処理用の設定項目を参照。

2. チェック処理
   
   1. 下記の引数を指定しPowerShellスクリプトを実行
      
      - 実行するPowerShellスクリプトは、「`C:/PyTkinterToPSScript/script/AdpackController.ps1`」
      
      - 入力データは、対象フォルダーの入力値
      
      - 出力データは、「`C:/PyTkinterToPSScript/input/Checklist_ZipFileList.csv`」
   
   2. 入力フォルダーにあるチェック対象のZIPファイルを帳票本文の実データファイルとして生成
      入力データは、前段のPowerShellスクリプトの事項結果、「`C:/PyTkinterToPSScript/input/Checklist_ZipFileList.csv`」を使用。
      出力データは、「`C:/PyTkinterToPSScript/input/Checklist_BodyValues.csv`」。
      

3. 出力処理
   
   下記の引数を指定しPowerShellスクリプトを実行
   
   - 実行するPowerShellスクリプトは、「`C:/PyTkinterToPSScript/script/PrintController.ps1`」
   
   - パック処理後の出力処理を意味する「`CheckForm`」を指定
     
     - 出力先は、入力フォルダーのひとつ上の階層に入力フォルダー名と同じ名前のPDFファイルを指定
       （例：入力フォルダーが`C:/Users/"ユーザー名"/YYYYMMDD_チェック`だった場合、出力ファイルは`C:/Users/"ユーザー名"/Desktop/YYYYMMDD_チェック.pdf` となる
   
   - E-DataPackerのインストール先「`C:/PyTkinterToPSScript`」
   
   - ヘッダーと本文の位置データファイルと実データファイルを指定

## モード変更（System003）

### 制約・前提条件

- モード変更の認証時は、設定ファイルに平文で保存されたパスワードを参照

- 設定ファイルのパスワードは、PyTkinterToPSScriptの起動時に読み取る

### 処理の流れ

1. パスワード認証

2. モード切り替え処理

## 設定変更

### 制約・前提条件

- ファイルの形式はXMLファイル

- 設定ファイルは、PyTkinterToPSScriptの起動時に読み取る

- クライアントアプリケーションであるため、環境毎に設定ファイルを変更可能

- 運用では、各環境の設定ファイルを管理すること

### 処理の流れ

1. ツール起動時 設定ファイル読込み

2. 設定変更処理

## バッチ・スクリプト

### AdpackController.ps1（パック・アンパックで使用）

引数一覧

- Pack or Unpack

- Hash
  
  - adpackに遵守（s256/s384/s512/sha1）
  
  - 自作関数を呼び出す場合（SHA256/SHA384/SHA512/SHA1）

- InputPath

- OutputPath

- Check or NoCheck（Unpack指定時のみ）

---

1. 使用するモジュールファイルの読み込み
   
   - 「`C:¥PyTkinterToPSScript¥script/PSModules/CommonModules.psm1`」
   
   - 「`C:¥PyTkinterToPSScript¥script/PSModules/AdpackModules.psm1`」

2. 引数の論理チェック
   
   - Pack と Unpack 両方指定時はエラー
   
   - Pack と Unpack どちらも指定していない場合はエラー
   
   - Pack 指定時に Check または NoCheck を指定するとエラー
   
   - Unpack 指定時に Check と NoCheck の両方を指定するとエラー

3. 引数毎に他の引数を論理チェック
   
   - Pack 指定
     
     - 入力データのパスがフォルダーであることを確認
     
     - 出力データのパスは入力データを元に自動的に設定
       入力データのフォルダーパスが `C:/Users/"ユーザー名"/Desktop/1234_XXXX` の場合は、`C:/Users/"ユーザー名"/Desktop/1234_XXXX.zip` となる。）
   
   - Unpack 指定
     
     - 入力データのパスがZIPファイルであることを確認
     
     - 出力データのパスは入力データを元に自動的に設定
       入力データのフォルダーパスが `C:/Users/"ユーザー名"/Desktop/1234_XXXX` の場合は、`C:/Users/"ユーザー名"/Desktop/1234_XXXX.pdf` となる。）

~~4. PowerShellバージョンの確認~~ 自作関数のAdpackを作成した場合は実装が必要

5. Adpack実行ファイルの存在チェック

6. 引数に応じてAdpackModules.psm1の関数を実行
   
   - Pack指定：Compress-Package_Adpackの実行
   - Unpack指定のみ：Expand-Package_Adpackを通常モードで実行
   - Unpack指定＋Check：Expand-Package_Adpackをチェックのみのモードで実行
   - Unpack指定＋NoCheck：Expand-Package_Adpackをチェック無しのモードで実行

### MultiCheckController.ps1（チェック処理で使用）

引数一覧

- InputPath

- OutputPath

---

1. 使用するモジュールファイルの読み込み
   
   - 「`C:¥PyTkinterToPSScript¥script/AdpackController.ps1」
   
   - 「`C:¥PyTkinterToPSScript¥script/PSModules/CommonModules.psm1`」
   
   - 「`C:¥PyTkinterToPSScript¥script/PSModules/AdpackModules.psm1`」

2. 引数の論理チェック
   
   - InputPath がフォルダーでない場合はエラー

~~3. PowerShellバージョンの確認~~ 自作関数のAdpackを作成した場合は実装が必要

4. Adpack実行ファイルの存在チェック

5．入力フォルダ―配下にあるZIPファイルのリストを作成

6. Adpackを使ったアンパック処理をZIPファイルのリスト分、実行
    ZIPファイル毎の実行結果をオブジェクトに格納。異常が発生した場合は中断。

7. 実行結果のオブジェクトをテキストベースのファイルに出力
    「`C:/PyTkinterToPSScript/input/Checklist_ZipFileList.csv`」

### PrintController.ps1（パック・アンパック後のPDF出力で使用）

### 前段の処理

- パック処理時
  パック処理で出力されるリスト、「`/META-INF/Manifest.xml`」からCSVファイル「`C:/PyTkinterToPSScript/input/Packagelist_BodyValues.csv'`」を生成

- アンパック処理時
  パック処理で出力されるリスト、「`C:/PyTkinterToPSScript/input/Checklist_ZipFileList.csv`」からCSVファイル「`C:/PyTkinterToPSScript/input/Checklist_BodyValues.csv`」を生成

### PowerShellの処理

引数一覧

- PackForm or CheckForm
  パック処理後の場合は、PackForm

- OutputPath

- RootPath：`C:/PyTkinterToPSScript`

- DataMapping_Header：`DataMapping_Header.csv`

- DataMapping_Body：`DataMapping_Body.csv`

- PackForm_Template：`Template_PackFormLists.xlsx` ※パック処理時

- PackForm_HeaderValues：`Packagelist_HeaderValues.csv` ※パック処理時

- PackForm_BodyValues：`Packagelist_BodyValues.csv` ※パック処理時

- CheckForm_Template：`Template_CheckFormLists.xlsx` ※アンパック処理時

- CheckForm_HeaderValues：`Checklist_HeaderValues.csv` ※アンパック処理時

- CheckForm_BodyValues：`Checklist_BodyValues.csv` ※アンパック処理時

### メイン処理

1. 使用するクラス・定数・モジュールファイルの読み込み
   
   - 「`C:¥PyTkinterToPSScript¥script/PSModules/CommonModules.psm1`」
   
   - 「`C:¥PyTkinterToPSScript¥script/PSModules/PrintModules.psm1`」

2. 引数の論理チェック
   
   - PackForm と CheckForm 両方を指定しているとエラー
   
   - PackForm と CheckForm どちらも指定していない場合にエラー

3. 各種フォルダー・ファイルの存在チェック
   
   - 共通部分
     
     - インストールフォルダー「`C:/PyTkinterToPSScript`」
     
     - テンプレートフォルダー「`C:/PyTkinterToPSScript/template`」
     
     - 帳票テンプレート - ヘッダー位置情報ファイル「`C:/PyTkinterToPSScript/template/DataMapping_Header.csv`」
     
     - 帳票テンプレート - 本文位置情報ファイル「`C:/PyTkinterToPSScript/template/DataMapping_Body.csv`」
   
   - パック処理時
     
     - 帳票テンプレートファイル「`C:/PyTkinterToPSScript/template/Template_PackFormLists.xlsx`」
     
     - 帳票テンプレート - ヘッダー実データファイル「`C:/PyTkinterToPSScript/input/Packagelist_HeaderValues.csv`」
     
     - 帳票テンプレート - 本文の実データファイル「`C:/PyTkinterToPSScript/input/Packagelist_BodyValues.csv`」
   
   - アンパック処理時
     
     - 帳票テンプレートファイル「`C:/PyTkinterToPSScript/template/Template_CheckFormLists.xlsx`」
     
     - 帳票テンプレート - ヘッダー実データファイル「`C:/PyTkinterToPSScript/input/Packagelist_HeaderValues.csv`」
     
     - 帳票テンプレート - 本文の実データファイル「`C:/PyTkinterToPSScript/input/Packagelist_BodyValues.csv`」

4. 帳票テンプレートファイル内のシート存在チェック
      1ページ目のシートと2ページ目以降のシートの存在を確認。

5. 帳票テンプレートファイルを一時フォルダ―にコピー
   コピー先は、`$Env:TEMP/E-DataPackger_ExcelTemplaete.xlsx`で以降の処理ではこのファイルを使用。
   なお、コピーする前に削除するが何らかの要因で削除ができ鳴った場合は、コピー先のファイル名が一意になるようファイル名を変更してコピーする。

6. ヘッダー情報と本文情報を読み込み
   読み込んだ後、それぞれの位置情報ファイルと実データファイルのオブジェクトを比較し構成が一致していることを確認

7. 帳票テンプレートファイルに値をセット
   各実データファイルを使用。

8. 帳票テンプレートファイルのテンプレート用のシートを削除

9. 帳票テンプレートファイルを開き、PDF出力
   実行する前に出力先をチェックし、存在する場合は削除
