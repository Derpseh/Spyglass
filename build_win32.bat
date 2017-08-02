@ECHO OFF
ECHO Building Spyglass-win32 binary using cx_freeze (python 2.7)
python2 C:\Python27\Scripts\cxfreeze Spyglass.py
MOVE /Y dist Spyglass-bin
COPY /Y UpdTime.py Spyglass-bin\src\UpdTime.py
COPY /Y Spyglass.py Spyglass-bin\src\Spyglass.py
ECHO Done.
PAUSE