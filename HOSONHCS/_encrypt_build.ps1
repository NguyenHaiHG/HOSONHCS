param([string]$projectDir, [string]$outputPath)

$key = [byte[]]@(0x4E,0x67,0x75,0x79,0x65,0x6E,0x48,0x61,0x69,0x48,0x47,0x5F,0x4B,0x65,0x79,0x32,0x30,0x32,0x35,0x5F,0x53,0x65,0x63,0x75,0x72,0x65,0x44,0x61,0x74,0x61,0x58,0x59)
$iv  = [byte[]]@(0x48,0x6F,0x53,0x6F,0x4E,0x48,0x43,0x53,0x5F,0x49,0x56,0x5F,0x32,0x30,0x32,0x35)

$jsonPath = Join-Path $projectDir "toanquoc.json"
$encSrc   = Join-Path $projectDir "toanquoc.enc"
$encOut   = Join-Path $outputPath  "toanquoc.enc"

$bytes = [System.IO.File]::ReadAllBytes($jsonPath)
$aes   = [System.Security.Cryptography.Aes]::Create()
$aes.Key = $key; $aes.IV = $iv
$aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
$aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
$enc = $aes.CreateEncryptor()
$ms  = New-Object System.IO.MemoryStream
$cs  = New-Object System.Security.Cryptography.CryptoStream($ms, $enc, [System.Security.Cryptography.CryptoStreamMode]::Write)
$cs.Write($bytes, 0, $bytes.Length)
$cs.FlushFinalBlock()
$result = $ms.ToArray()
$cs.Dispose(); $ms.Dispose(); $aes.Dispose()

[System.IO.File]::WriteAllBytes($encSrc, $result)
[System.IO.File]::WriteAllBytes($encOut, $result)
Remove-Item $jsonPath -Force

Write-Host "[EncryptBuild] toanquoc.enc updated ($([math]::Round($result.Length/1MB,1)) MB)"
