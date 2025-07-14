[Setup]
AppName=城北中央公園テニスコート予約システム
AppVersion=1.1.0
DefaultDirName={autopf}\JohokuTennisApp
DefaultGroupName=城北中央公園テニスコート予約システム
OutputDir=installer
OutputBaseFilename=JohokuTennisApp_Setup_v1.1.0
Compression=lzma
SolidCompression=yes
SetupIconFile=app_icon.ico
PrivilegesRequired=admin
DisableProgramGroupPage=yes
WizardStyle=modern
ShowLanguageDialog=no

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成(&D)"; GroupDescription: "追加のアイコン:"; Flags: unchecked

[Files]
; メインアプリケーション（更新版）
Source: "dist\JohokuTennisApp.exe"; DestDir: "{app}"; Flags: ignoreversion
; アイコンファイル
Source: "app_icon.ico"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
; スタートメニューのショートカット
Name: "{group}\城北中央公園テニスコート予約システム"; Filename: "{app}\JohokuTennisApp.exe"; IconFilename: "{app}\app_icon.ico"
Name: "{group}\アンインストール"; Filename: "{uninstallexe}"
; デスクトップのショートカット
Name: "{autodesktop}\城北中央公園テニスコート予約システム"; Filename: "{app}\JohokuTennisApp.exe"; IconFilename: "{app}\app_icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\JohokuTennisApp.exe"; Description: "アプリケーションを起動"; Flags: nowait postinstall skipifsilent

[Code]
// Chrome がインストールされているかチェック
function IsChromeInstalled: Boolean;
var
  ChromePath: String;
begin
  Result := False;
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe', '', ChromePath) then
    Result := FileExists(ChromePath)
  else if RegQueryStringValue(HKCU, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe', '', ChromePath) then
    Result := FileExists(ChromePath)
  else if FileExists(ExpandConstant('{pf}\Google\Chrome\Application\chrome.exe')) then
    Result := True
  else if FileExists(ExpandConstant('{pf32}\Google\Chrome\Application\chrome.exe')) then
    Result := True;
end;

// インストール前の準備
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if not IsChromeInstalled then
  begin
    if MsgBox('Google Chrome がインストールされていません。' + #13#10 +
              'このアプリケーションを正常に動作させるには Google Chrome が必要です。' + #13#10#13#10 +
              'インストールを続行しますか？', 
              mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := 'Google Chrome が必要です。';
    end;
  end;
end;