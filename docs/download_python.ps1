$url = "https://www.python.org/ftp/python/3.13.13/python-3.13.13-amd64.exe"
$dest = "C:\Users\topge\Downloads\python-3.13.13-amd64.exe"
Write-Host "Downloading Python 3.13.13..."
Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
Write-Host "Download complete."
Write-Host "SHA256:"
(Get-FileHash $dest -Algorithm SHA256).Hash
