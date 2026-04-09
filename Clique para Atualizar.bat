@echo off
:: Garante que o terminal mude para a pasta onde o .bat está salvo
cd /d "%~dp0"

:: Executa o python chamando o app.py da mesma pasta
python app.py

exit