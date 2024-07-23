$exclude = @("venv", "extrair_dolar.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "extrair_dolar.zip" -Force