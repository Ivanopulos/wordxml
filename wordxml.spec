# -*- mode: python ; coding: utf-8 -*-


import sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p)
block_cipher = None

a = Analysis(['wordxml.py'],
             pathex=["C:\\Users\\IMatveev\\PycharmProjects\\wordxml\\exe"],
             binaries=[],
             datas=[(path.join(site_packages,"docx","templates"), "docx/templates")],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts, 
          exclude_binaries=True,
          name='wordxml',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas, 
               strip=False,
               upx=True,
               upx_exclude=[],
               name='wordxml')
