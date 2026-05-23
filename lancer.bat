@echo off
echo Installation des dependances (premiere fois seulement)...
pip install -r requirements.txt -q
echo.
echo Lancement de l'application...
echo L'interface va s'ouvrir dans ton navigateur.
echo Pour arreter : appuie sur CTRL+C dans cette fenetre.
echo.
streamlit run app.py
pause
