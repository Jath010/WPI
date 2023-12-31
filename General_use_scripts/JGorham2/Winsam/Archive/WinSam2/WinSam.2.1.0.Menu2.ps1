#Requires -version 2.0

<#
 -----------------------------------------------------------------------------
 Script: Demo-ConsoleMenu.ps1
 Version: 1.0
 Author: Jeffery Hicks
    http://jdhitsolutions.com/blog
    http://twitter.com/JeffHicks
    http://www.ScriptingGeek.com
 Date: 12/30/2011
 Keywords: Read-Host, Menu, Switch
 Comments:
 The best way to create a menu is to use a here string
 
 Use -ClearScreen if you want to run CLS before displaying the menu

 "Those who forget to script are doomed to repeat their work."

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************
 -----------------------------------------------------------------------------
 #>
 
Function Show-Menu {

Param(
[Parameter(Position=0,Mandatory=$True,HelpMessage="Enter your menu text")]
[ValidateNotNullOrEmpty()]
[string]$Menu,
[Parameter(Position=1)]
[ValidateNotNullOrEmpty()]
[string]$Title="Menu",
[switch]$ClearScreen
)

if ($ClearScreen) {Clear-Host}

#build the menu prompt
$menuPrompt=$title
#add a return
$menuprompt+="`n"
#add an underline
$menuprompt+="-"*$title.Length
$menuprompt+="`n"
#add the menu
$menuPrompt+=$menu

Read-Host -Prompt $menuprompt

} #end function

#define a menu here string
$menu=@"
1 Show info about a computer
2 Show info about someones mailbox
3 Restarts the print spooler
Q Quit

Select a task by number or Q to quit
"@

#Keep looping and running the menu until the user selects Q (or q).
Do {
    #use a Switch construct to take action depending on what menu choice
    #is selected.
    Switch (Show-Menu $menu "My Help Desk Tasks" -clear) {
     "1" {Write-Host "run get info code" -ForegroundColor Yellow
         sleep -seconds 2
         } 
     "2" {Write-Host "run show mailbox code" -ForegroundColor Green
          sleep -seconds 5
          }
     "3" {Write-Host "restart spooler" -ForegroundColor Magenta
         sleep -seconds 2
         }
     "Q" {Write-Host "Goodbye" -ForegroundColor Cyan
         Return
         }
     Default {Write-Warning "Invalid Choice. Try again."
              sleep -milliseconds 750}
    } #switch
} While ($True)

