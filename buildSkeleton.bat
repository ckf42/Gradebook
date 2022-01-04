@echo off
setlocal
if not exist build mkdir build
cd skeleton
tar -a -c -f ..\build\skeleton.zip *
move /-Y ..\build\skeleton.zip ..\build\skeleton.xlsm
