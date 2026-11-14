; Generated from src/fileTypes.json by scripts/sync-file-types.mjs. Do not edit by hand.

; Uninstalling removes what Next Office registered with Windows at start-up ("Open with" entries and
; its page in Settings > Default apps). An update keeps it.
!macro NSIS_HOOK_POSTUNINSTALL
  ${If} $UpdateMode <> 1
    DeleteRegKey HKCU "Software\Classes\NextOffice.xlsx"
    DeleteRegValue HKCU "Software\Classes\.xlsx\OpenWithProgids" "NextOffice.xlsx"
    DeleteRegKey HKCU "Software\Classes\NextOffice.xlsm"
    DeleteRegValue HKCU "Software\Classes\.xlsm\OpenWithProgids" "NextOffice.xlsm"
    DeleteRegKey HKCU "Software\Classes\NextOffice.csv"
    DeleteRegValue HKCU "Software\Classes\.csv\OpenWithProgids" "NextOffice.csv"
    DeleteRegKey HKCU "Software\Classes\NextOffice.tsv"
    DeleteRegValue HKCU "Software\Classes\.tsv\OpenWithProgids" "NextOffice.tsv"
    DeleteRegKey HKCU "Software\Classes\NextOffice.nxsheet"
    DeleteRegValue HKCU "Software\Classes\.nxsheet\OpenWithProgids" "NextOffice.nxsheet"
    DeleteRegKey HKCU "Software\Classes\NextOffice.pptx"
    DeleteRegValue HKCU "Software\Classes\.pptx\OpenWithProgids" "NextOffice.pptx"
    DeleteRegKey HKCU "Software\Classes\NextOffice.ppsx"
    DeleteRegValue HKCU "Software\Classes\.ppsx\OpenWithProgids" "NextOffice.ppsx"
    DeleteRegKey HKCU "Software\Classes\NextOffice.potx"
    DeleteRegValue HKCU "Software\Classes\.potx\OpenWithProgids" "NextOffice.potx"
    DeleteRegKey HKCU "Software\Classes\NextOffice.pptm"
    DeleteRegValue HKCU "Software\Classes\.pptm\OpenWithProgids" "NextOffice.pptm"
    DeleteRegKey HKCU "Software\Classes\NextOffice.nxslides"
    DeleteRegValue HKCU "Software\Classes\.nxslides\OpenWithProgids" "NextOffice.nxslides"
    DeleteRegKey HKCU "Software\Classes\NextOffice.docx"
    DeleteRegValue HKCU "Software\Classes\.docx\OpenWithProgids" "NextOffice.docx"
    DeleteRegKey HKCU "Software\Classes\NextOffice.docm"
    DeleteRegValue HKCU "Software\Classes\.docm\OpenWithProgids" "NextOffice.docm"
    DeleteRegKey HKCU "Software\Classes\NextOffice.dotx"
    DeleteRegValue HKCU "Software\Classes\.dotx\OpenWithProgids" "NextOffice.dotx"
    DeleteRegKey HKCU "Software\Classes\NextOffice.doc"
    DeleteRegValue HKCU "Software\Classes\.doc\OpenWithProgids" "NextOffice.doc"
    DeleteRegKey HKCU "Software\Classes\NextOffice.dot"
    DeleteRegValue HKCU "Software\Classes\.dot\OpenWithProgids" "NextOffice.dot"
    DeleteRegKey HKCU "Software\Classes\NextOffice.nxdoc"
    DeleteRegValue HKCU "Software\Classes\.nxdoc\OpenWithProgids" "NextOffice.nxdoc"
    DeleteRegKey HKCU "Software\Classes\Applications\Next Office.exe"
    DeleteRegKey HKCU "Software\NextOffice"
    DeleteRegValue HKCU "Software\RegisteredApplications" "Next Office"
    System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
  ${EndIf}
!macroend
