; Script generated for Inno Setup
; Scarica Inno Setup da: https://jrsoftware.org/isdl.php

#define MyAppName "Catalogo Manager Pro"
#define MyAppVersion "1.1.2"
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
; Nota: I file .db e .json verranno creati automaticamente dall'app al primo avvio

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent