@echo off
title Album Poster Generator Setup
echo Installing/Updating dependencies...
py -m pip install python-pptx pillow scikit-learn numpy requests streamlit

echo Launching Album Poster Generator...
py -m streamlit run src/app.py

pause