# -*- mode: python -*-

block_cipher = None


a = Analysis(['EXCP_Main.py'],
             pathex=['G:\\Excel-Com\\Excel_comp'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
a.datas +=[("Res/Logo.jpg","G:/Excel-Com/Excel_comp/Res/Logo.jpg","DATA")]
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='EXCP_Main',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
