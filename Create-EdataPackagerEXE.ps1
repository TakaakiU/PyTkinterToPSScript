# pyinstaller -n PyTkinterToPSScript --add-data "classes/config/settings.xml;classes/config" --add-data "classes/image/logo.png;classes/image" --hidden-import babel.numbers --clean --onefile --noconsole --collect-data tkinterdnd2 --paths="C:\Users\Administrator\Documents\Git\python\PyTkinterToPSScript" main.py
pyinstaller -n PyTkinterToPSScript --add-data "classes/image/logo.png;classes/image" --hidden-import babel.numbers --clean --onefile --noconsole --collect-data tkinterdnd2 --paths="C:\Users\Administrator\Documents\Git\python\PyTkinterToPSScript" main.py
# Update after release
# pyinstaller .\PyTkinterToPSScript.spec

# Copy the created executable file
$copyFrom = ".\dist\PyTkinterToPSScript.exe"
$copyTo = ".\classes\exe\PyTkinterToPSScript.exe"
Copy-Item -Path $copyFrom -Destination $copyTo -Force
