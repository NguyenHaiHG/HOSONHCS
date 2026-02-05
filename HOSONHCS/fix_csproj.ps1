$content = Get-Content "HOSONHCS.csproj" -Raw
$content = $content -replace '<Compile Include="Models\.cs" />`r`n    <Compile Include="Program\.cs" />', '<Compile Include="Models.cs" />
    <Compile Include="Program.cs" />'
Set-Content "HOSONHCS.csproj" -Value $content -NoNewline
