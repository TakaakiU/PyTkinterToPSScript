# SETUP - PyTkinterToPSScript

## はじめに

◯◯

### 開発環境の構成

#### ハードウェア・仮想環境の情報

Oracle VirtualBoxを使った仮想環境となる。

##### ホストOS情報

- OS：Windows 10 Pro

- CPU：Intel(R) Core(TM) i3-8100T CPU @ 3.10GHz

- メモリ：16GB

##### ゲストOS情報

- OS：Windows 10 Pro

- CPU：2CPU（割り当てCPUのプロセッサー数）

- メモリ：8GB割り当て（8,192MB）

#### ソフトウェア

- ホストOS
  
  - Oracle VirtualBox 7.1.6

- ゲストOS
  
  - Inno Setup 6.4.2
  
  - Microsoft Office Standard 2019　※ Excelテンプレート操作、およびPDF印刷で使用。
  
  - Visual Studio Code（VS Code）　※ Python・PowerShellの開発に使用。

#### ソースコード（ゲストOS内）

- PowerShellのコード「`"ルートリポジトリ"classes\script`」

- Pythonのコード「拡張子が`*.py`のファイル」
  ※ 詳細は下記のツリーを参照。

```
.\PyTkinterToPSScript
├──　< .vscode >
│　　└──　launch.json                                         ・・・Visual Studio Codeの設定ファイル
│　　　　　
│　　
├──　< classes >
│　　│　　
│　　├──　< config >                                          ・・・インストーラーのパッケージ対象
│　　│　　└──　settings.xml                                  ・・・設定ファイルの原本
│　　│　　
│　　├──　< control >
│　　│　　├──　__init__.py
│　　│　　├──　ctrlBatch.py                                  ・・・PowerShellの実行
│　　│　　├──　ctrlCommon.py                                 ・・・Pythonで使用する共通モジュール
│　　│　　├──　ctrlConfig.py                                 ・・・設定ファイルの操作
│　　│　　├──　ctrlCsv.py                                    ・・・CSVファイルの操作
│　　│　　├──　ctrlExcel.py                                  ・・・Excelファイルの操作
│　　│　　├──　ctrlMessage.py                                ・・・エラーメッセージの制御
│　　│　　└──　ctrlString.py                                 ・・・文字列操作のモジュール
│　　│　　
│　　├──　< exe >                                             ・・・インストーラーのパッケージ対象
│　　│　　├──　adpack_exe.md                                 ・・・adpack.exeのダミーファイル（大容量のため）
│　　│　　└──　PyTkinterToPSScript.exe                      ・・・Pyinstallerでexe化した実行ファイル
│　　│　　
│　　├──　< form >
│　　│　　├──　< __pycache__ >
│　　│　　│　　
│　　│　　├──　__init__.py
│　　│　　├──　formAuth.py                                   ・・・パスワード認証画面
│　　│　　├──　formBase.py                                   ・・・メイン画面の継承元
│　　│　　└──　formPackageMain.py                            ・・・メイン画面
│　　│　　
│　　├──　< image >
│　　│　　│　　
│　　│　　└──　logo.png                                     ・・・メイン画面のロゴ画像ファイル
│　　│　　
│　　├──　< script >                                         ・・・インストーラーのパッケージ対象
│　　│　　├──　< PSClasses >
│　　│　　│　　└──　PrintListConfig.ps1                    ・・・PowerShell内で使うカスタムクラスの定義
│　　│　　│　　
│　　│　　├──　< PSModules >
│　　│　　│　　├──　AdpackModules.psm1                     ・・・Adpackを使うモジュール
│　　│　　│　　├──　CommonModules.psm1                     ・・・PowerShell共通モジュール
│　　│　　│　　└──　PrintModules.psm1                      ・・・PDF出力で使うモジュール
│　　│　　│　　
│　　│　　├──　< PSRefresh >
│　　│　　│　　├──　Refresh-CDrivePSScript.ps1             ・・・開発フォルダーのPowerShellを元にC:¥PyTkinterToPSScriptを洗い替える
│　　│　　│　　├──　Refresh-ExecutableFile.ps1             ・・・開発フォルダーの実行ファイルを元にC:¥PyTkinterToPSScriptを洗い替える
│　　│　　│　　└──　Refresh-SettingsXML.ps1                ・・・開発フォルダーの設定ファイルを元にC:¥PyTkinterToPSScriptを洗い替える
│　　│　　│　　
│　　│　　├──　AdpackController.ps1                         ・・・Adpack（パッケージ化）の実行スクリプト
│　　│　　├──　MultiCheckController.ps1                     ・・・チェック処理の実行スクリプト
│　　│　　├──　PrintController.ps1                          ・・・帳票印刷（PDF出力）の実行スクリプト
│　　│　　└──　RefreshController.ps1                        ・・・PSRefreshフォルダー配下をすべて実行するスクリプト
│　　│　　
│　　├──　< structure >
│　　│　　├──　< __pycache__ >
│　　│　　│　　
│　　│　　├──　__init__.py
│　　│　　└──　structureEntrydata.py                        ・・・メイン画面 入力情報の構造体
│　　│　　
│　　├──　< template >                                       ・・・インストーラーのパッケージ対象
│　　│　　├──　DataMapping_Body.csv                         ・・・画面データの基本情報とテンプレートファイルのデータを紐づけるファイル
│　　│　　├──　DataMapping_Header.csv                       ・・・画面データとテンプレートファイルの紐づけファイル
│　　│　　├──　Template_CheckFormLists.xlsx                 ・・・チェック処理の帳票テンプレートファイル
│　　│　　└──　Template_PackFormLists.xlsx                  ・・・パッケージ化処理の帳票テンプレートファイル
│　　│　　
│　　└──　__init__.py
│　　
├──　< dist >
│　　└──　PyTkinterToPSScript.exe                                ・・・Create-EdataPackagerEXE.ps1の実行で生成する実行ファイル
│　　
├──　__init__.py
├──　Create-EdataPackagerEXE.ps1                                ・・・実行ファイル生成用のスクリプト
├──　PyTkinterToPSScript.spec                                        ・・・Pythonシステムファイル
├──　main.py                                                    ・・・起動する画面を指定するメインPythonファイル
└──　RefreshController.ps1 - ショートカット.lnk                ・・・classes\script\RefreshController.ps1のショートカット
```

#### Pythonライブラリの一覧

```powershell
# Pythonバージョン
PS C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript> python -V
Python 3.10.5
PS C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript> 

# Pythonライブラリ一覧
PS C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript> pip list
Package                   Version
---------------------------------------
altgraph                  0.17.4        
arrow                     1.3.0
autopep8                  2.3.2
babel                     2.17.0        
binaryornot               0.4.4
cachetools                5.5.2
certifi                   2025.1.31     
chardet                   5.2.0
charset-normalizer        3.4.1
click                     8.1.8
colorama                  0.4.6
cookiecutter              2.6.0
et_xmlfile                2.0.0
flake8                    7.1.2
future                    1.0.0
idna                      3.10
Jinja2                    3.1.6
jinja2-time               0.2.0
lxml                      5.3.1
markdown-it-py            3.0.0
MarkupSafe                3.0.2
mccabe                    0.7.0
mdurl                     0.1.2
MouseInfo                 0.1.3
mypy                      1.15.0
mypy-extensions           1.0.0
numpy                     2.2.4
openpyxl                  3.1.5
packaging                 24.2
pandas                    2.2.3
pefile                    2023.2.7
pillow                    11.1.0
pip                       25.0.1
pip-review                1.3.0
py2exe                    0.13.0.2
pyasn1                    0.6.1
PyAutoGUI                 0.9.54
pycodestyle               2.12.1
pycryptodome              3.22.0
pyflakes                  3.2.0
PyGetWindow               0.0.9
Pygments                  2.19.1
pyinstaller               6.12.0
pyinstaller-hooks-contrib 2025.2
PyMsgBox                  1.0.9
pyperclip                 1.9.0
PyRect                    0.2.0
PyScreeze                 1.0.1
pysmb                     1.2.10
python-dateutil           2.9.0.post0
python-slugify            8.0.4
pytweening                1.2.0
pytz                      2025.1
pywin32                   310
pywin32-ctypes            0.2.3
PyYAML                    6.0.2
requests                  2.32.3
rich                      13.9.4
setuptools                77.0.3
six                       1.17.0
text-unidecode            1.3
tkcalendar                1.6.1
tkinterdnd2               0.4.3
toml                      0.10.2
tomli                     2.2.1
tqdm                      4.67.1
ttkthemes                 3.2.2
types-python-dateutil     2.9.0.20241206
typing_extensions         4.12.2
tzdata                    2025.2
urllib3                   2.3.0
PS C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript> 
```

#### Visual Studio Code 拡張機能の一覧

```powershell
# Visual Studio Codeのバージョン
PS C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript> code -v
1.99.3
17baf841131aa23349f217ca7c570c76ee87b957
x64
PS C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript> 

# Visual Studio Code 拡張機能のバージョン
PS C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript> code --list-extensions --show-versions
grapecity.gc-excelviewer@4.2.63
ms-ceintl.vscode-language-pack-ja@1.99.2025041609
ms-python.debugpy@2025.6.0
ms-python.python@2025.4.0
ms-python.vscode-pylance@2025.4.1
ms-vscode.powershell@2025.0.0
PS C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript> 
```

#### 設定ファイル（settings.xml）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<settings>
	<basicsettings>
		<hash_algorithm>s256</hash_algorithm>                ・・・圧縮や解凍ともにファイル容量に制限がadpackにあり、リリースタイミングで自作関数を作成するか判断予定
		<package_maxsize>0</package_maxsize>                 ・・・パッケージ化処理では、ファイルサイズを1.5GB（1,610,612,736 Byte = 1.5 * (1,024 * 1,024 * 1,024）
		<package_maxfiles>0</package_maxfiles>
		<check_maxsize>0</check_maxsize>                     ・・・チェック処理では、制限をなしとする。
		<check_maxfiles>0</check_maxfiles>
		<password>password</password>                        ・・・チェックモードに切り替える際のパスワード ※要件により平文で保存
	</basicsettings>
	<combosettings>                                          ・・・コンボボックスの選択肢 ※ソフトの再起動で反映。
		<targetrange>
			<value>System A</value>
			<value>System B</value>
		</targetrange>
		<workername>
			<value>System Department 1 - Taro A</value>
			<value>System Department 2 - Biko</value>
		</workername>
		<terminalname>
			<value>Management Number 001</value>
			<value>Management Number 002</value>
		</terminalname>
	</combosettings>
</settings>
```

#### インストール先の構成

`C:¥PyTkinterToPSScript` の構成

```
C:\PyTkinterToPSScript
├──　< config >
│　　└──　settings.xml                              ・・・起動時に参照する設定ファイル
│　　
├──　< exe >
│　　├──　adpack.exe                                ・・・AdDataPackagerの実行ファイル
│　　└──　PyTkinterToPSScript.exe                  ・・・PyTkinterToPSScript本体、実行ファイル
│　　
├──　< input >                                       ・・・画面で入力中間ファイルの保存フォルダー
│　　├──　Checklist_BodyValues.csv                  ・・・チェックリスト用の中間ファイル
│　　├──　Checklist_HeaderValues.csv                ・・・同上
│　　├──　Checklist_ZipFileList.csv                 ・・・チェック処理実行後に出力する中間ファイル
│　　├──　Packagelist_BodyValues.csv                ・・・パッケージリスト用の中間ファイル
│　　└──　Packagelist_HeaderValues.csv              ・・・同上
│　　
├──　< output >                                      ・・・画面で出力する中間ファイルの保存フォルダー
│　　├──　CheckLists.pdf
│　　└──　PackageLists.pdf
│　　
├──　< script >                                      ・・・ソースコードで記載している為、割愛
│　　├──　< PSClasses >
│　　│　　└──　PrintListConfig.ps1
│　　│　　
│　　├──　< PSModules >
│　　│　　├──　AdpackModules.psm1
│　　│　　├──　CommonModules.psm1
│　　│　　└──　PrintModules.psm1
│　　│　　
│　　├──　AdpackController.ps1
│　　├──　MultiCheckController.ps1
│　　├──　PrintController.ps1
│　　└──　RefreshController.ps1
│　　
└──　< template >                                   ・・・帳票作成時に使用するテンプレートファイル
    │                                                 　　　※ スクリプト実行時は、$Env:TEMP（C:\Users\ADMINI~1\AppData\Local\Temp）に一時ファイルとしてコピー。
    │                                                 　　　　 その後、編集およびPDF出力、一時ファイルの削除を実行。
　　├──　DataMapping_Body.csv
　　├──　DataMapping_Header.csv
　　├──　Template_CheckFormLists.xlsx
　　└──　Template_PackFormLists.xlsx
```

## 開発・クライアント環境の作業手順

### 【開発環境】仮想環境の開発環境の起動と表示

1. Oracle VM VirtualBox マネージャーを起動

2. 開発環境の「対象の仮想マシン」を選択し、「起動(T)」ボタンをクリック

3. 仮想の開発環境を起動できたを確認

※ **シャットダウンの手順**は、ゲストOSのWindows でシャットダウンを操作してください。

### 【開発環境】開発リソースをインストールフォルダーにコピーする手順

開発環境内でテストする際に対応する手順。ソースコード内にあるリソースをインストールフォルダーにスクリプトを使ってコピーする手順です。

1. ソースフォルダーを開く
   Winキー ＋ R → 「`C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript`」

2. `RefreshController.ps1 - ショートカット` を右クリックして「PowerShell で実行」

3. 上記によりソースフォルダーにある `PyTkinterToPSScript.exe` と `adpack.exe` がインストールフォルダー（`C:¥PyTkinterToPSScript`）の配下に上書きコピーされたことを確認

4. `RefreshController.ps1` を右クリックして「PowerShell で実行」

5. 上記によりソースフォルダーのリソースがインストールフォルダーにコピーされてことを確認

6. これで開発環境のインストールフォルダーが最新の状態となりました。

### 【開発環境】インストーラー作成手順

1. ソースフォルダー内のインストーラーフォルダーを開く
   Winキー ＋ R → 「`C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript\installer`」
2. Inno Setup Scriptファイル「`inno-setup_installer.iss`」をダブルクリックで開く
3. 必要に応じて設定内容を変更する
   インストーラーのバージョン「`#define MyAppVersion "0.0.1"`」
4. コンパイルボタン（もしくは Ctrl + F9 を押す）
5. インストーラーフォルダー配下にインストーラーファイル「`edatapackager-0.0.1.exe`」が生成されたことを確認

```inno-setup_installer.iss
; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "PyTkinterToPSScript"
#define MyAppVersion "0.0.1"
#define MyAppPublisher "Bozo Research Center Inc"
#define MyAppURL "https://www.bozo.co.jp"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{3564F68E-ACC3-4E0F-B743-28102309D014}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName=C:\PyTkinterToPSScript
DisableDirPage=yes
DefaultGroupName={#MyAppName}
; Uncomment the following line to run in non administrative install mode (install for current user only).
;PrivilegesRequired=lowest
OutputDir=C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript\installer
OutputBaseFilename=edatapackager-0.0.1
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Dirs]
Name: "{app}\input"
Name: "{app}\output"

[Files]
Source: "C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript\classes\config\*"; DestDir: "{app}\config"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript\classes\exe\*"; DestDir: "{app}\exe"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript\classes\script\*"; DestDir: "{app}\script"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript\classes\template\*"; DestDir: "{app}\template"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\"ユーザー名"\Documents\Git\python\PyTkinterToPSScript\installer\PyTkinterToPSScript.exe - ショートカット.lnk"; DestDir: "{commondesktop}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
```

### 【クライアント環境】 PyTkinterToPSScriptインストール手順

※ PyTkinterToPSScriptは、**Microsoft Office 2019以降**のインストールが必須。

1. インストーラーファイル「`edatapackager-X.X.X.exe`」をローカルにダウンロード

2. インストーラーファイルをダブルクリック

3. ユーザーアカウント制御が表示された場合は、「はい」を選択

4. インストールボタンをクリック（インストール先は"C:¥PyTkinterToPSScript"の固定）

5. インストール準備完了できたこと

6. 下記のポイントを確認し、インストールが完了したことを確認
   
   - `C:\PyTkinterToPSScript`配下にリソースが配置されていること
   
   - デスクトップにPyTkinterToPSScript実行ファイル(`PyTkinterToPSScript.exe`)のショートカットが配置されたこと

### 【クライアント環境】 PyTkinterToPSScriptアンインストール手順

1. Windows 設定のアプリを開く

2. PyTkinterToPSScriptの選択し「アンインストール」をクリック

3. ユーザーアカウント制御が表示された場合は、「はい」をクリック

4. 削除の確認で「はい」をクリック

5. 正常に削除（アンインストール）されたことを確認

6. 場合によって帳票の中間ファイルがインストール先に残存しているので、必要に応じて削除
