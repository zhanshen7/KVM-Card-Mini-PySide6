$ErrorActionPreference = "Stop"

Push-Location $PSScriptRoot
try {
    & pyside6-lupdate.exe -no-obsolete .\main.py .\ui\main.ui .\ui\device_setup_dialog.ui -ts .\trans_cn.ts
    & pyside6-linguist.exe .\trans_cn.ts
    & pyside6-lrelease.exe .\trans_cn.ts -qm .\trans_cn.qm
}
finally {
    Pop-Location
}
