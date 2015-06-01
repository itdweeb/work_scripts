$tempfolders = @("C:\Windows\Temp\*", "C:\Users\*\Appdata\Local\Temp\*")

Stop-Process -Name iexplore -Force
Stop-Process -Name winword -Force
Stop-Process -Name excel -Force
Stop-Process -Name AcroRd32 -Force
Stop-Process -Name outlook -Force
Stop-Process -Name pdfDocs -Force

Remove-Item $tempfolders -force -recurse

