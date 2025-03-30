# build_exe.py
import os
import PyInstaller.__main__

# Convert PNG to ICO if needed (recommended for Windows)
try:
    from PIL import Image
    img = Image.open('icon.png')
    img.save('icon.ico', sizes=[(32,32), (48,48), (64,64), (128,128)])
    icon_file = 'icon.ico'
except ImportError:
    print("Pillow not installed, using PNG directly")
    icon_file = 'icon.png'
except Exception as e:
    print(f"Couldn't convert icon: {e}")
    icon_file = 'icon.png'

# Build command
PyInstaller.__main__.run([
    'excel_processor.py',
    '--onefile',
    '--windowed',
    '--clean',
    '--name=ExcelProcessor',
    f'--icon={icon_file}',
    '--add-data=icon.png;.',
    '--noconfirm'
])