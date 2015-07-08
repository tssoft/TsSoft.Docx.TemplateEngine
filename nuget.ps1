function Get-CurrentLineNumber { 
    $MyInvocation.ScriptLineNumber 
}

del TsSoft.Docx.TemplateEngine.*.nupkg
echo Get-CurrentLineNumber
del *.nuspec
echo Get-CurrentLineNumber
del .\TsSoft.Docx.TemplateEngine\*\bin\*.nuspec
echo Get-CurrentLineNumber
function GetNodeValue([xml]$xml, [string]$xpath)
{
	return $xml.SelectSingleNode($xpath).'#text'
}

function SetNodeValue([xml]$xml, [string]$xpath, [string]$value)
{
	$node = $xml.SelectSingleNode($xpath)
	if ($node) {
		$node.'#text' = $value
	}
}

Remove-Item .\TsSoft.Docx.TemplateEngine\*\bin -Recurse 
echo Get-CurrentLineNumber
Remove-Item .\TsSoft.Docx.TemplateEngine\*\obj -Recurse
echo Get-CurrentLineNumber

$build = "c:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe ""TsSoft.Docx.TemplateEngine\TsSoft.Docx.TemplateEngine.csproj"" /p:Configuration=Release" 
Invoke-Expression $build
echo Get-CurrentLineNumber

$Artifact = (resolve-path ".\TsSoft.Docx.TemplateEngine\bin\Release\TsSoft.Docx.TemplateEngine.dll").path

.nuget\nuget spec -F -A $Artifact

Copy-Item TsSoft.Docx.TemplateEngine.nuspec.xml .\TsSoft.Docx.TemplateEngine\bin\Release\TsSoft.Docx.TemplateEngine.nuspec

$GeneratedSpecification = (resolve-path ".\TsSoft.Docx.TemplateEngine.nuspec").path
$TargetSpecification = (resolve-path ".\TsSoft.Docx.TemplateEngine\bin\Release\TsSoft.Docx.TemplateEngine.nuspec").path

[xml]$srcxml = Get-Content $GeneratedSpecification
[xml]$destxml = Get-Content $TargetSpecification
$value = GetNodeValue $srcxml "//version"
SetNodeValue $destxml "//version" $value;
$value = GetNodeValue $srcxml "//description"
SetNodeValue $destxml "//description" $value;
$value = GetNodeValue $srcxml "//copyright"
SetNodeValue $destxml "//copyright" $value;
$destxml.Save($TargetSpecification)

.nuget\nuget pack $TargetSpecification

del *.nuspec
del .\TsSoft.Docx.TemplateEngine\bin\Release\*.nuspec

exit
