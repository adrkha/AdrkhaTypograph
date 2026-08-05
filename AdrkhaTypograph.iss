; ============================================================
;  Adrkha Typograph - Inno Setup Script
;  الإصدار: 1.0.0
;  يبني مثبت Windows احترافي لإضافة PowerPoint
; ============================================================

#define AppName      "Adrkha Typograph"
#define AppVersion   "1.0.0"
#define AppPublisher "Adrkha"
#define AppURL       "https://powerpoint.adrkha.com"
#define AppExeName   "AdrkhaTypograph"
; معرّف الإضافة في Registry (يجب أن يطابق ProgId في VSTO)
#define AddInProgId  "AdrkhaTypograph"

[Setup]
AppId={{E3A7B2C1-1234-5678-ABCD-9F0E1D2C3B4A}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=Installer
OutputBaseFilename={#AppExeName}_v{#AppVersion}_Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\icon.ico
; لغة RTL عربية (اختياري - يحتاج ملف لغة خارجي)
; LanguageDetectionMethod=uilanguage

; معلومات الناشر للملف التنفيذي (Metadata)
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Setup
VersionInfoVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoCopyright=Copyright (C) 2026 {#AppPublisher}

[Languages]
Name: "arabic"; MessagesFile: "compiler:Default.isl"

[Files]
; ملفات الإضافة الرئيسية
Source: "bin\Release\AdrkhaTypograph.dll";              DestDir: "{app}"; Flags: ignoreversion
Source: "PublishFiles\AdrkhaTypograph.dll.manifest";     DestDir: "{app}"; Flags: ignoreversion
Source: "PublishFiles\AdrkhaTypograph.vsto";             DestDir: "{app}"; Flags: ignoreversion

; مكتبات HarfBuzz و SkiaSharp
Source: "bin\Release\HarfBuzzSharp.dll";              DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\SkiaSharp.dll";                  DestDir: "{app}"; Flags: ignoreversion

; المجلدات الفرعية للمكتبات الأصلية (مهم جداً للعمل على بيئات 64 بت و 32 بت)
Source: "bin\Release\x64\libSkiaSharp.dll";           DestDir: "{app}\x64"; Flags: ignoreversion
Source: "bin\Release\x64\libHarfBuzzSharp.dll";       DestDir: "{app}\x64"; Flags: ignoreversion
Source: "bin\Release\x86\libSkiaSharp.dll";           DestDir: "{app}\x86"; Flags: ignoreversion
Source: "bin\Release\x86\libHarfBuzzSharp.dll";       DestDir: "{app}\x86"; Flags: ignoreversion
Source: "bin\Release\arm64\libSkiaSharp.dll";         DestDir: "{app}\arm64"; Flags: ignoreversion skipifsourcedoesntexist
Source: "bin\Release\arm64\libHarfBuzzSharp.dll";     DestDir: "{app}\arm64"; Flags: ignoreversion skipifsourcedoesntexist

; مكتبات System المساعدة
Source: "bin\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "bin\Release\System.Memory.dll";              DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "bin\Release\System.Buffers.dll";             DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "bin\Release\System.Numerics.Vectors.dll";    DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "bin\Release\System.Runtime.CompilerServices.Unsafe.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Registry]
; تسجيل الإضافة في PowerPoint (لكل مستخدم حالي)
; LoadBehavior=3 → تحميل تلقائي عند بدء PowerPoint
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInProgId}"; \
    ValueType: string;  ValueName: "Description";   ValueData: "تحويل خصائص الخط إلى شكل"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInProgId}"; \
    ValueType: string;  ValueName: "FriendlyName";  ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInProgId}"; \
    ValueType: dword;   ValueName: "LoadBehavior";  ValueData: "3"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInProgId}"; \
    ValueType: string;  ValueName: "Manifest";      ValueData: "{code:GetManifestUrl}"

[Code]
// دالة لتحويل مسار المجلد المحلي إلى صيغة URL (file:///) المطلوبة من VSTO
function GetManifestUrl(Param: String): String;
var
  AppDir: String;
begin
  AppDir := ExpandConstant('{app}');
  StringChangeEx(AppDir, '\', '/', True);
  Result := 'file:///' + AppDir + '/AdrkhaTypograph.vsto|vstolocal';
end;

// تحقق من وجود VSTO Runtime في جميع مساراته المحتملة
function VSTORuntimeInstalled(): Boolean;
var
  S: String;
begin
  Result :=
    // VS 2022 / Office 365 (64-bit path)
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Version', S) or
    // VS 2019 / 2017 (32-bit WOW path)
    RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R', 'Version', S) or
    // نسخ أقدم
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4', 'Version', S) or
    RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4', 'Version', S) or
    // Office مثبت مع VSTO مدمج
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Office\15.0\Common\VSTO Runtime Setup', 'Version', S) or
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Office\16.0\Common\VSTO Runtime Setup', 'Version', S);
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if not VSTORuntimeInstalled() then
  begin
    if MsgBox('لم يتم العثور على Microsoft Visual Studio Tools for Office Runtime.' + #13#10 +
              'هذا المكون مطلوب لتشغيل الإضافة. يمكنك تثبيته مجاناً من موقع Microsoft.' + #13#10 + #13#10 +
              'هل تريد الاستمرار في التثبيت على أي حال؟',
              mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;
