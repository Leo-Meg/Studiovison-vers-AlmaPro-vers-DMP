# sv_export.spec
# PyInstaller spec file pour creer ExportStudioVision_AlmaPro.exe
# Icone composite : moitie gauche StudioVision / moitie droite AlmaPro

block_cipher = None

a = Analysis(
    ['sv_export.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('export_sv_alma.ico', '.'),
    ],
    hiddenimports=[
        'pyodbc',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'tkinter',
        'tkinter.scrolledtext',
        'tkinter.messagebox',
        'xml.etree.ElementTree',
        'xml.dom.minidom',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=['pandas','numpy','matplotlib','scipy','PIL'],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExportStudioVision_AlmaPro',
    debug=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='export_sv_alma.ico',
)
