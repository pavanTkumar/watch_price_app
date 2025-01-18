import PyInstaller.__main__

PyInstaller.__main__.run([
    'watch_app.py',
    '--onefile',
    '--windowed',
    '--name=WatchBatteryPricing',
    '--noconsole',
])