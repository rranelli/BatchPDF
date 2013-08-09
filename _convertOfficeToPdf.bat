: #############################################
: Arquivo em lotes para convers�o autom�tica dos arquivos Word e Excel em .pdf
: O javascript usa o m�todo "saveaspdf" do objeto application.word.
: Implementado por Renan Ranelli e Paulo Roberto Polastri, 04/12, CHT/SP
: #############################################


echo off

echo ##########################################################
echo ####  CHEMTECH's MSoffice to .pdf batch converter v0.2 ###
echo ##########################################################

: ###############
: O cscript usa o m�todo "saveaspdf" do objeto application.word.


: este loop varre todos os arquivos *.doc* e executa a convers�o para pdf.
echo .
echo =======================================
echo "Converting .doc* files in the folder "
echo =======================================
echo .

for  %%a in (*.doc) do cscript.exe //nologo //E:jscript _SaveDOCasPDF.js "%%a"

: este loop varre todos os arquivos *.xls* e executa a convers�o para pdf.
echo .
echo =======================================
echo "Converting .xls* files in the folder "
echo =======================================
echo .

for  %%a in (*.xls) do cscript.exe //nologo //E:jscript _SaveXLSasPDF.js "%%a"


:#####
: fim
pause