# projektCitius
# setup
# Download and install Python silently with PATH added
$pythonURL = "https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe"
$installerPath = "$env:TEMP\python_installer.exe"

Invoke-WebRequest -Uri $pythonURL -OutFile $installerPath
Start-Process -Wait -FilePath $installerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_launcher=0" -PassThru

# Refresh PATH to detect Python immediately
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

# Install required libraries
py -m pip install --upgrade pip
py -m pip install -r requirements.txt

# Verify
py -c "import docx; print('Python and all libraries installed successfully!')"
# intalling depencies and libaries use this command