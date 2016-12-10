# -*- mode: python -*-

block_cipher = None


a = Analysis(['images_rc.pyc', 'C:/Users/Jwesner/Documents/AsciiToWord_Project/filesToWordConverter.py'],
             pathex=['C:\\Users\\Jwesner\\Documents\\AsciiToWord_Project'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=['.'],
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
          name='images_rc',
          debug=False,
          strip=False,
          upx=True,
          console=True )
