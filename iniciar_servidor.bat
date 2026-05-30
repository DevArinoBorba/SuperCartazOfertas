@echo off
echo ==============================================
echo Iniciando Servidor de Cartazes na Rede Local
echo ==============================================
echo.

echo 1. Ativando o ambiente virtual e instalando dependencias...
call venv\Scripts\activate.bat
pip install -r requirements.txt --quiet
pip install waitress --quiet

echo.
echo 2. Iniciando a aplicacao...
echo.
echo Para aplicar mudancas nos arquivos, use R + Enter nesta janela.
echo Para encerrar o servidor, use S + Enter.
echo.
python server_manager.py

pause
