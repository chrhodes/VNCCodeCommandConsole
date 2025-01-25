
$folders = @(
    # "VNC.ActiveDirectoryHelper"
    # , "VNC.AssemblyHelper\VNC.AssemblyHelper"
    # , "VNC.AZDO"
    # , "VNC.AZDOHelper"
    # , "VNC.CodeAnalysis\VNC.CodeAnalysis"
    # , "VNC.Core"
    # , "VNC.HttpHelper"
    # , "VNC.Logging\VNC.Logging"
    # , "VNC.VNC.SMOHelper"
    # , "VNC.SPHelper\VNC.SPHelper"
    # , "VNC.TFSHelper\VNC.TFSHelper"
    # , "VNC.WPF.Presentation"
    # , "VNC.WPF.Presentation.Dx"
    )

foreach ($folder in $folders)
{
    "Removing obj\ and bin\ folder contents in $folder"

    if (Test-Path -Path $folder\obj)
    {
        remove-item $folder\obj -Recurse -Force
    }

    if (Test-Path -Path $folder\bin)
    {
        remove-item $folder\bin -Recurse -Force
    }

}

Read-Host -Prompt "Press Enter to Exit"