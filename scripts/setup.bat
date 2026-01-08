@echo off
echo Criando ambiente virtual...
python -m venv .venv

echo Ativando ambiente virtual...
call .venv\Scripts\activate

echo Atualizando pip...
python -m pip install --upgrade pip

echo Instalando dependências do requirements.txt...
pip install -r requirements.txt

echo ---
echo Ambiente configurado! Para rodar o app:
echo.
echo     call .venv\Scripts\activate
echo.
pause

