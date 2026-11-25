; electron-builder NSIS custom installer script (Windows)
; Place in your Electron project and reference via build.nsis.include in package.json

!macro customInstall
  ; 1. Trust the local Root CA in the Windows Machine Root Store
  nsExec::ExecToLog 'certutil -addstore -f "Root" "$INSTDIR\resources\certs\rootCA.crt"'

  ; 2. Copy manifest.xml to a shared local Trusted Catalog folder
  CreateDirectory "$COMMONPROGRAMDATA\MarketDataAddin\excel-catalog"
  CopyFiles /SILENT "$INSTDIR\resources\manifest.xml" "$COMMONPROGRAMDATA\MarketDataAddin\excel-catalog\manifest.xml"

  ; 3. Register Trusted Add-in Catalog in Windows Registry for Excel
  ; Replace GUID below with a fixed GUID for your catalog
  WriteRegStr HKCU "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A4E8F120-9D34-4B1E-8C22-112233445566}" "Id" "{A4E8F120-9D34-4B1E-8C22-112233445566}"
  WriteRegStr HKCU "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A4E8F120-9D34-4B1E-8C22-112233445566}" "Url" "$COMMONPROGRAMDATA\MarketDataAddin\excel-catalog"
  WriteRegDWORD HKCU "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A4E8F120-9D34-4B1E-8C22-112233445566}" "Flags" 1
!macroend

!macro customUnInstall
  DeleteRegKey HKCU "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A4E8F120-9D34-4B1E-8C22-112233445566}"
  RMDir /r "$COMMONPROGRAMDATA\MarketDataAddin\excel-catalog"
!macroend
