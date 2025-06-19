c:\workbench\projects\tests\genai\venv\Scripts\pyinstaller.exe --onefile -w ..\src\zkteco_machine_proxy.py
move /y dist\zkteco_machine_proxy.exe .\
rmdir /S /Q build
rmdir /S /Q dist
del /q *.spec