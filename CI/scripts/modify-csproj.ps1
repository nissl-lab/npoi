# Get a list of all csproj files in the repository
$csprojFiles = Get-ChildItem -Path . -Filter *.csproj -Recurse

# Loop through each csproj file
foreach ($csprojFile in $csprojFiles) {
    # Read the contents of the csproj file into a string
    $csproj = Get-Content $csprojFile.FullName

    # Use a regular expression to match the PackageReference element
    $csproj = $csproj -Replace '<PackageReference Include="IronDevTools.PerformanceAnalyzer" Version="\d+.\d+.\d+" />', '<!-- <PackageReference Include="IronDevTools.PerformanceAnalyzer" Version="\d+.\d+.\d+" /> -->'

    # Write the modified string back to the csproj file
    Set-Content $csprojFile.FullName $csproj
}