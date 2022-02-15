# -*- mode: python ; coding: utf-8 -*-

block_cipher = None
python_dir = 'C:\\Development\\Anaconda3.9\\envs\\python3.9\\'

# exclude Tk, Tcl completely
import sys
sys.modules['FixTk'] = None

xlmerge_a = Analysis(['xlmerge\\xlmerge.py'],
                     pathex=['C:\\Users\\dks\\Programming\\python\\xlmerge\\xlmerge',
                             python_dir + 'Library\\bin',
                             python_dir + 'Lib\\site-packages\\PyQt5\\Qt5\\bin'],
                     binaries=[],
                     datas=[],
                     hiddenimports=[],
                     hookspath=[],
                     runtime_hooks=[],
                     excludes=['FixTk', 'tcl', 'tk', '_tkinter', 'tkinter', 'Tkinter',
                               'docutils', 'babel', 'sphinx', 'numpy', 'pandas',
                               'matplotlib', 'notebook', 'ipython'],
                     win_no_prefer_redirects=False,
                     win_private_assemblies=False,
                     cipher=block_cipher,
                     noarchive=False)
xlmerge_pyz = PYZ(xlmerge_a.pure, xlmerge_a.zipped_data,
                  cipher=block_cipher)
xlmerge_exe = EXE(xlmerge_pyz,
                  xlmerge_a.scripts,
                  [],
                  exclude_binaries=True,
                  name='xlmerge',
                  debug=False,
                  bootloader_ignore_signals=False,
                  strip=True,
                  upx=False,
                  console=False )

addin_a = Analysis(['xlmerge\\addin.py'],
                   pathex=['C:\\Users\\dks\\Programming\\python\\xlmerge',
                           python_dir + 'Library\\bin',
                           python_dir + 'Lib\\site-packages\\PyQt5\\Qt5\\bin'],
                   binaries=[],
                   datas=[('xlmerge\\addin.template', '.'),
                          ('xlmerge\\Excel.officeUI', '.')],
                   hiddenimports=[],
                   hookspath=[],
                   runtime_hooks=[],
                   excludes=['FixTk', 'tcl', 'tk', '_tkinter', 'tkinter', 'Tkinter',
                             'docutils', 'babel', 'sphinx', 'numpy', 'pandas',
                             'matplotlib', 'notebook', 'ipython'],
                   win_no_prefer_redirects=False,
                   win_private_assemblies=False,
                   cipher=block_cipher,
                   noarchive=False)
addin_pyz = PYZ(addin_a.pure, addin_a.zipped_data,
                cipher=block_cipher)
addin_exe = EXE(addin_pyz,
                addin_a.scripts,
                [],
                exclude_binaries=True,
                name='addin',
                debug=False,
                bootloader_ignore_signals=False,
                strip=True,
                upx=True,
                console=False )

coll = COLLECT(xlmerge_exe,
               xlmerge_a.binaries,
               xlmerge_a.zipfiles,
               xlmerge_a.datas,
               addin_exe,
               addin_a.binaries,
               addin_a.zipfiles,
               addin_a.datas,
               strip=False,
               upx=False,
               upx_exclude=[],
               name='xlmerge')
