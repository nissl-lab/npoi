param($installPath, $toolsPath, $package, $project)
$asms = $package.AssemblyReferences | %{$_.Name}
foreach ($reference in $project.Object.References)
{
    if ($asms -contains $reference.Name + ".dll")
    {
        if ($reference.Name -ne "nunit.framework")
        {
            $reference.CopyLocal = $false;
        }
    }
}
