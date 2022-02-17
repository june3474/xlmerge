;-------------------------------------------------------------------------------
; Includes
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "WinVer.nsh"
!include "x64.nsh"

;-------------------------------------------------------------------------------
; Constants
!define PRODUCT_NAME "xlmerge"
!define PRODUCT_DESCRIPTION "Merge Excel sheets"
!define COPYRIGHT "MIT License 2022 dks"
!define PRODUCT_VERSION "1.0.0.0"
!define SETUP_VERSION 1.0.0.0
# additional constants
!define ADDIN_DIR "$APPDATA\Microsoft\AddIns"
!define UI_DIR "$LOCALAPPDATA\Microsoft\Office"
!define RELEASE_VERSION "1.0"


;-------------------------------------------------------------------------------
; Attributes
Name "xlmerge"
OutFile "${PRODUCT_NAME}-${RELEASE_VERSION}_Setup.exe"
InstallDir "$LOCALAPPDATA\${PRODUCT_NAME}"
# Let us not pollute registry
;InstallDirRegKey HKCU "Software\${PRODUCT_NAME}" ""
RequestExecutionLevel user ; user|highest|admin

;-------------------------------------------------------------------------------
; Version Info
VIProductVersion "${PRODUCT_VERSION}"
VIAddVersionKey "ProductName" "${PRODUCT_NAME}"
VIAddVersionKey "ProductVersion" "${PRODUCT_VERSION}"
VIAddVersionKey "FileDescription" "${PRODUCT_DESCRIPTION}"
VIAddVersionKey "LegalCopyright" "${COPYRIGHT}"
VIAddVersionKey "FileVersion" "${SETUP_VERSION}"

;-------------------------------------------------------------------------------
; Modern UI Appearance
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "${NSISDIR}\Contrib\Graphics\Header\win.bmp"
!define MUI_WELCOMEFINISHPAGE_BITMAP "${NSISDIR}\Contrib\Graphics\Wizard\win.bmp"
!define MUI_FINISHPAGE_NOAUTOCLOSE

;-------------------------------------------------------------------------------
; Installer Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

;-------------------------------------------------------------------------------
; Uninstaller Pages
!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

;-------------------------------------------------------------------------------
; Languages
!insertmacro MUI_LANGUAGE "Korean"

;-------------------------------------------------------------------------------
; Installer Sections - Main
Section "Main" main_sec
    SectionIn RO
	SetOutPath $INSTDIR
	File /r "dist\${PRODUCT_NAME}\*"
	WriteUninstaller "$INSTDIR\Uninstall.exe"
    # Uninstall registry keys. But let's not pollute registry
    ;WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "DisplayName" "${PRODUCT_NAME}"
    ;WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
SectionEnd

; Installer Sections - Addin
Section "Addin" addin_sec
	SetOutPath $INSTDIR
    ExecWait '"$INSTDIR\addin.exe" "$INSTDIR"'
    ;CreateDirectory ${ADDIN_DIR}
    CopyFiles "$INSTDIR\xlmerge.xlam" ${ADDIN_DIR}
    CopyFiles "$INSTDIR\Excel.officeUI" ${UI_DIR}
SectionEnd
    
;-------------------------------------------------------------------------------
; Section Descriptions
LangString DESC_main ${LANG_KOREAN} "Main module"
LangString DESC_addin ${LANG_KOREAN} "엑셀에서 xlmerge를 바로 실행할 수 있는 부가 모듈"

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
!insertmacro MUI_DESCRIPTION_TEXT ${main_sec} $(DESC_main)
!insertmacro MUI_DESCRIPTION_TEXT ${addin_sec} $(DESC_addin)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

;-------------------------------------------------------------------------------
; Uninstaller Sections
Section "Uninstall"
	Delete "$INSTDIR\Uninstall.exe"
	RMDir /r "$INSTDIR"
    Delete "${ADDIN_DIR}\xlmerge.xlam"
	Delete "${UI_DIR}\Excel.officeUI"
	;DeleteRegKey HKCU "Software\${PRODUCT_NAME}"
	;DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}    
SectionEnd

