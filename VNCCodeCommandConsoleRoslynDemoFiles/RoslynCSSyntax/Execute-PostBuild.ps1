$targetFolder = "C:\temp\VNCCodeCommandConsoleTestFiles"

$sourceFolders = @(
    "DemoFiles"
    , "DesignChecks"
    , "PerformanceChecks"
    , "QualityChecks"
    )
    
foreach ($folder in $sourceFolders)
{
    "Copying files from $folder to $targetFolder"
    Copy-Item -Path $folder -Recurse -Force -Destination  $targetFolder
}