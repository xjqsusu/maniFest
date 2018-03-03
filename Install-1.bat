

cd /d %~dp0

msiexec /i python-2.7.14.msi /qn+ ADDLOCAL=ALL TargetDir=c:\Python27

copy pywin32-222-cp27-cp27m-win32.whl c:\Python27\Scripts

cd /d c:\Python27\Scripts

pip install pywin32-222-cp27-cp27m-win32.whl

cd /d C:\Python27\Lib\site-packages\pywin32_system32

copy pywintypes27.dll C:\Python27\Lib\site-packages\win32

copy pythoncom27.dll C:\Python27\Lib\site-packages\win32

PAUSE