# -*- mode: python -*-

block_cipher = None


a = Analysis(['central-line-event-calculator.py'],
             pathex=['d:\\projects\\med\\central-line-event-calculator'],
             binaries=None,
             datas=None,
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='central-line-event-calculator',
          debug=False,
          strip=False,
          upx=True,
          console=False )
