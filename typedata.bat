@echo on

call "C:\ProgramData\Anaconda3\Scripts\activate.bat"
call conda activate typedata
call jupyter-notebook --notebook-dir="C:\Users\STJ2TW\Python"  

pause
