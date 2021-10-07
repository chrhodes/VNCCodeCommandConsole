<# 

.SYNOPSIS 
A brief description of the function or script. 
This keyword can be used only once in each topic.

.DESCRIPTION
A detailed description of the function or script.
This keyword can be used only once in each topic
		
.PARAMETER firstNamedArgument
The description of a parameter. You can include a Parameter keyword for
each parameter in the function or script syntax.

The Parameter keywords can appear in any order in the comment block, but
the function or script syntax determines the order in which the parameters
(and their descriptions) appear in Help topic. To change the order,
change the syntax.
 
You can also specify a parameter description by placing a comment in the
function or script syntax immediately before the parameter variable name.
If you use both a syntax comment and a Parameter keyword, the description
associated with the Parameter keyword is used, and the syntax comment is
ignored.

.PARAMETER secondNamedArgument
blah blah about secondNamedArgument

.EXAMPLE
A sample command that uses the function or script, optionally followed
by sample output and a description. Repeat this keyword for each example.
.EXAMPLE
Example2

.INPUTS
Inputs to this cmdlet (if any)

.OUTPUTS
Output from this cmdlet (if any)

.NOTES
Additional information about the function or script.

.LINK
The name of a related topic. Repeat this keyword for each related topic.

This content appears in the Related Links section of the Help topic.

The Link keyword content can also include a Uniform Resource Identifier
(URI) to an online version of the same Help topic. The online version 
opens when you use the Online parameter of Get-Help. The URI must begin
with "http" or "https".

.COMPONENT
The technology or feature that the function or script uses, or to which
it is related. This content appears when the Get-Help command includes
the Component parameter of Get-Help.

.ROLE
The user role for the Help topic. This content appears when the Get-Help
command includes the Role parameter of Get-Help.

.FUNCTIONALITY
The intended use of the function. This content appears when the Get-Help
command includes the Functionality parameter of Get-Help.


<ScriptName - Consider Verb-Noun>.ps1

SCC:
	This script is under source code control.  Modifications should be 
	checked into the TFS repository located at 
		<Team Project Collection>
	under a project 
		$<TeamProject>/<Path>

Last Update:

v1.0.0 <Author>, <Date>, <Company>

Be sure to leave two blank lines after end of block comment.
#>

##############################################    
# Script Level Parameters.
##############################################

param
(
# <TODO: Add script level parameters>
    # [switch] $SwitchArg1,
    # [switch] $SwitchArg2,
    [string] $Configuration, 
    [string] $Platform,
	[string] $TargetName,
    [switch] $Contents,
    [switch] $Verbose
)

##############################################    
# Script Level Variables
##############################################

<#
$ScriptVar1 = "Foo"
$ScriptVar2 = 42
$ScriptVar3 = @(
    ("Apple")
    ,("Pear")
    ,("Yoghurt")
)
#>

$UsePLLog = $false

$SCRIPTNAME = $MyInvocation.MyCommand.Name
$SCRIPTPATH = & { $myInvocation.ScriptName }
$CURRENTDIRECTORY = $PSScriptRoot

##############################################
# Main function
##############################################

function Main
{
    if ($SCRIPT:Verbose)
    {
        "SCRIPTNAME         = $SCRIPTNAME"
		"SCRIPTPATH         = $SCRIPTPATH"
		"CURRENTDIRECTORY   = $CURRENTDIRECTORY"
		
        "Configuration      = $Configuration"
        "Platform           = $Platform"
        "TargetName         = $TargetName"
<#
        "ScriptVar2         = $ScriptVar2"
        "ScriptVar3         = $ScriptVar3"
        ""
#>
		"`$Verbose           = $Verbose"
    }
    
    if ( ! (VerifyPrerequisites))
    {
        LogMessage "Error Verifying Prerequisites" "Main" "Error"
        exit
    }
    else
    {
        LogMessage "Prerequisites OK" "Main" "Info"
    }

    $message = "Beginning " + $SCRIPTNAME + ": " + (Get-Date)
    LogMessage $message "Main" "Info"
    
# <TODO: Add code, functional calls here to do something cool>

cd $CURRENTDIRECTORY

    Func1
    
#    Func2
    
#    Func3
    
    $message = "Ending   " + $SCRIPTNAME + ": " + (Get-Date)
    LogMessage $message "Main" "Info"
}

##############################
# Internal Functions
##############################

function Func1()
{
#TODO - Get this from env
    $message = $SCRIPTNAME
    LogMessage $message $SCRIPTNAME "Info"

    # TODO
    # Maybe switch and handle unexpected config

    # if ($Configuration -eq "Debug")
    # {
        # $destinations = @(
	        # "..\Common\Debug"
	        # )
    # }
    # else
    # {
        # $destinations = @(
	        # "..\Common"
	        # )
    # }

    # $targets = @(
	    # ".\bin\$Configuration\VNC.Prism.Infrastructure.dll"
	    # ".\bin\$Configuration\VNC.Prism.Infrastructure.pdb"
	    # )
	
    # "pushing new targets to destinations"

    # foreach ($destination in $destinations)
    # {
	    # $destination
	
	    # foreach ($target in $targets)
	    # {
		    # $target
		    # copy-item -path $target -destination $destination
	    # }
    # }
}

function Func2()
{
    $message = "Doing Something cool in Func2"
    LogMessage $message "Func2" "Info"
}

function Func3()
{
    $message = "Doing Something cool in Func3"
    LogMessage $message "Func3" "Info"
}

##############################
# Internal Support Functions
##############################

# function Write-Status($message, $color)
# {
    # $message | Write-Host -ForegroundColor $color;
# }

# function LogMessage()
# {
    # param
    # (
        # [string] $message,
        # [string] $method,
        # [string] $logLevel
    # )
    
	# # <TODO: Each case can be modified to do the appropriate type of console/PLLog logging.
    
    # switch ($logLevel)
    # {
        # "Trace"
        # { 
            # if ($SCRIPT:Verbose) { Write-Status $message  "Yellow"}
            # if ($SCRIPT:UsePLLog) { Call-PLLog -Trace   -message $message -class "Process-DLPFiles" -method $method }
            # break
        # }    
        
        # "Info"
        # { 
            # # if ($SCRIPT:Verbose) { Write-Host $message }   
			# Write-Status $message "Green"
            # if ($SCRIPT:UsePLLog) { Call-PLLog -Info    -message $message -class "Process-DLPFiles" -method $method }
            # break
        # }
        
        # "Warning"
        # { 
            # Write-Status $message  "Orange"
            # if ($SCRIPT:UsePLLog) { Call-PLLog -Warning -message $message -class "Process-DLPFiles" -method $method }
            # break
        # }
		
        # "Error"
        # { 
            # Write-Status $message  "Red"
            # if ($SCRIPT:UsePLLog) { Call-PLLog -Error   -message $message -class "Process-DLPFiles" -method $method }
            # break
        # }
        
        # "None"
        # { 
            # if ($SCRIPT:Verbose) { Write-Status $message }        
            # break
        # }
        
        # default
        # {
            # Write-Status $message        
            # if ($SCRIPT:UsePLLog) {  Call-PLLog -Error "Unexpected log level" + $logLevel -class "Process-DLPFiles" -method "LogMessage" }
            # break
        # }
    # }
# }

function VerifyFunc1()
{

    $message = "  VerifyFunc1()"
    LogMessage $message "VerifyFunc1" "Trace"

    # Verify something
    
    # if ( ! (Test-Path $InputFolder))
    # {
        # $message = "InputFolder: " + $InputFolder + " does not exist"
        # LogMessage $message "VerifyInputFiles" "Error"

        # return $false
    # }
    # else
    # {
        # foreach ($file in $EDMFileNames)
        # {
            # $inputFile = ($InputFolder + "\" + $file + ".csv")
            
            # if ( ! (Test-Path $inputFile))
            # {
                # $message = "    Missing Input file: " + $inputFile
                # LogMessage $message "VerifyInputFiles" "Error"

                # return $false
            # }
        # }
    # }
    
    return $true
}
    
function VerifyFunc2()
{
    $message = "  VerifyFunc2()"
    LogMessage $message "VerifyFunc2" "Trace"

    return $true
}
    
function VerifyPrerequisites()
{
    $message = "VerifyPrerequisites()"
    LogMessage $message "VerifyPrerequisites" "Trace"

    if ( ! (VerifyFunc1))
    {
        return $false
    }
    
    if ( ! (VerifyFunc2))
    {
        return $false
    }
            
    return $true
}

if ($SCRIPT:Contents)
{
	$myInvocation.MyCommand.ScriptBlock
	exit
}
	
# Call the main function.  Use Dot Sourcing to ensure executed in Script scope.

. Main

#
# End New-ScriptTemplate1.ps1
#