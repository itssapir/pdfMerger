# -*- mode: python -*-

block_cipher = None


a = Analysis(['pdfMerger.pyw'],
             pathex=['C:\\Users\\itssa\\Desktop'],
             binaries=[],
             datas= [('C:\\Users\\itssa\\AppData\\Local\\Programs\\Python\\Python37-32\\Lib\\site-packages\\ttkthemes', 'ttkthemes')],
             hiddenimports=['ttkthemes'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='pdfMerger',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
