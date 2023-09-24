@ECHO OFF
@REM ECHO "%~1"  &:: 드래그 앤 드롭으로 갖다놓은 파일명

@REM CALL conda activate ppt_idx
@REM python main.py "%~1"

%USERPROFILE%\anaconda3\envs\ppt_idx\python.exe main.py "%~1"

@REM PAUSE
