# -*- mode: python ; coding: utf-8 -*-

datas = []
datas += [("dist\\DocGen.exe", "payload")]
datas += [("dist\\DocGenUninstall.exe", "payload")]
datas += [("安装包使用说明.txt", "payload")]
import glob
import os
import zipfile

_spec_dir = globals().get("SPECPATH") or os.getcwd()
_payload_build_dir = os.path.join(_spec_dir, "build", "payload")
os.makedirs(_payload_build_dir, exist_ok=True)
_templates_zip = os.path.join(_payload_build_dir, "word_template.zip")
with zipfile.ZipFile(_templates_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
    for fp in glob.glob(os.path.join(_spec_dir, "word_template", "*.docx")):
        zf.write(fp, arcname=os.path.join("word_template", os.path.basename(fp)))
datas += [(_templates_zip, "payload")]

hiddenimports = []
hiddenimports += ["PySide6"]

a = Analysis(
    ["docgen_installer.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="DocGenInstaller_Setup",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
