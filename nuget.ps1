del TsSoft.Docx.TemplateEngine.*.nupkg
del *.nuspec
del .\TsSoft.Docx.TemplateEngine\*\bin\*.nuspec

Remove-Item .\TsSoft.Docx.TemplateEngine\*\bin -Recurse 
Remove-Item .\TsSoft.Docx.TemplateEngine\*\obj -Recurse

$build = "c:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe ""TsSoft.Docx.TemplateEngine\TsSoft.Docx.TemplateEngine.csproj"" /p:Configuration=Release" 
Invoke-Expression $build

$Artifact = (resolve-path ".\TsSoft.Docx.TemplateEngine\bin\Release\TsSoft.Docx.TemplateEngine.dll").path

.nuget\nuget spec -F -A $Artifact

Copy-Item TsSoft.Docx.TemplateEngine.nuspec.xml .\TsSoft.Docx.TemplateEngine\bin\Release\TsSoft.Docx.TemplateEngine.nuspec

$GeneratedSpecification = (resolve-path ".\TsSoft.Docx.TemplateEngine.nuspec").path
$TargetSpecification = (resolve-path ".\TsSoft.Docx.TemplateEngine\bin\Release\TsSoft.Docx.TemplateEngine.nuspec").path

[xml]$srcxml = Get-Content $GeneratedSpecification
[xml]$destxml = Get-Content $TargetSpecification

function EnsureMetadataNodeExists([xml]$destination, [string]$name)
{
	if (!$destination.package.metadata.$name)
	{
		$destination.package.metadata.AppendChild($destination.CreateElement($name))
	}
}

function CopyMetadata([xml]$source, [xml]$destination, [string]$name) 
{
	EnsureMetadataNodeExists $destination $name
	$destination.package.metadata.$name = $source.package.metadata.$name
	return $null;
}

echo 'funcs declared'

CopyMetadata $srcxml $destxml 'version'

echo 'version copied'

CopyMetadata $srcxml $destxml 'description'

echo 'description copied'

CopyMetadata $srcxml $destxml 'copyright'

echo 'copyright copied'

$destxml.Save($TargetSpecification)

echo 'saved'

.nuget\nuget pack $TargetSpecification

echo 'packed'

del *.nuspec
del .\TsSoft.Docx.TemplateEngine\bin\Release\*.nuspec

echo 'deleted'

exit
