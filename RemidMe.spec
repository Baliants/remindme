# -*- mode: python -*-

block_cipher = None


a = Analysis(['RemidMe.py'],
             pathex=['k:\\project'],
             hiddenimports=['win32timezone'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

for d in a.datas:
    if 'pyconfig' in d[0]:
        a.datas.remove(d)
        break
		
a.datas += [ ('remind.ico', '.\\remind.ico', 'DATA')]
a.datas += [ ('open.gif', '.\\open.gif', 'DATA')]
a.datas += [ ('about.gif', '.\\about.gif', 'DATA')]
	
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)		
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='RemidMe',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='K:\\project\\remind.ico')
