@echo off

@echo off

SET /P selected="VBAコードをエクスポートしますか？(Y/N)？"

if /i {%selected%}=={y} (goto :yes)
if /i {%selected%}=={yes} (goto :yes)

exit

:yes
pushd %0\..
cscript //nologo vbac.wsf decombine

exit
