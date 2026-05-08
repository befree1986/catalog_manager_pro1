; Script generated for Inno Setup
; Scarica Inno Setup da: https://jrsoftware.org/isdl.php

#define MyAppName "Catalogo Manager Pro"
#define MyAppVersion "1.1.6b"
#define MyAppPublisher "BeFree"
#define MyAppURL "https://github.com/befree1986/catalog_manager_pro1"
#define MyAppExeName "CatalogoApp.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
AppId={{A1B2C3D4-E5F6-7890-1234-56789ABCDEF0}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
; Installa nella cartella locale dell'utente per evitare problemi di permessi di scrittura (per i file json/db)
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; Assicura che venga mostrata la pagina per scegliere la cartella di installazione
DisableDirPage=no
; Richiede solo permessi utente standard (non amministratore)
PrivilegesRequired=lowest
OutputDir=.
OutputBaseFilename=CatalogoManager_Setup
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
LicenseFile=license.txt

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Assicurati di aver compilato prima il progetto con: pyinstaller CatalogoApp.spec
Source: "dist\CatalogoApp.exe"; DestDir: "{app}"; Flags: ignoreversion
; Includi il redistributable scaricato. Viene copiato nella cartella temporanea e rimosso dopo l'installazione.
Source: "VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
; Nota: I file .db e .json verranno creati automaticamente dall'app al primo avvio

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Esegue l'installazione del VC Redist in modalità silenziosa prima di completare il setup
Filename: "{tmp}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; \
    Check: VCRedistNeedsInstall; StatusMsg: "Installazione dei componenti di sistema (Microsoft Visual C++ Redistributable)..."

Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[Code]
function VCRedistNeedsInstall: Boolean;
var
  Version: String;
begin
  // Verifica se la versione 2015-2022 (v14.x) è già presente nel registro di sistema
  if RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64', 'Version', Version) then
  begin
    // Se trova il valore, confronta o semplicemente considera che sia installato
    Log('VC Redist v14 trovato: ' + Version);
    Result := False;
  end
  else
  begin
    Log('VC Redist non trovato, procedo con l''installazione.');
    Result := True;
  end;
end;