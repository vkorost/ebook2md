# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['convert_to_md.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'torchvision', 'torchaudio', 'scipy', 'numpy', 'pandas',
              'sklearn', 'matplotlib', 'PIL', 'tensorflow', 'transformers',
              'pytest', 'pygments', 'sqlalchemy', 'cryptography', 'bcrypt',
              'onnxruntime', 'lightning', 'yt_dlp', 'av', 'fsspec',
              'uvicorn', 'websockets', 'anyio', 'jsonschema', 'opentelemetry',
              'grpc', 'openpyxl', 'win32com'],
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
    name='convert_to_md',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
