$folder="a";
foreach ($file in Get-ChildItem $folder)
{
.\JavaToCSharp "$folder\$file";
}