<# Paths
$pathChrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$pathExcel = "C:\Program Files\Microsoft Office\root\Office16\excel.exe"
$pathOutlook = "C:\Program Files\Microsoft Office\root\Office16\outlook.exe"
$pathPowerPoint = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
$pathWord = "C:\Program Files\Microsoft Office\root\Office16\winword.exe"
$pathCherwell = "C:\Program Files (x86)\Cherwell Software\Cherwell Service Management\Trebuchet.App.exe"
$pathMaple = "C:\Program Files\Maple 2022\bin.X86_64_WINDOWS\maplew.exe"
$pathMathcad = "C:\Program Files (x86)\Mathcad\Mathcad 15\mathcad.exe"
$pathSolstice = "C:\Program Files\Mersive\SolsticeClient\SolsticeClient.exe"
$pathZoom = "C:\Program Files\Zoom\bin\Zoom.exe"
#>

function Repair-ASRShortcuts {
    [CmdletBinding()]
    param (
        $hostname
    )
    
    begin {
        
        $StartMenuFolder = "\\$hostname\c$\ProgramData\Microsoft\Windows\Start Menu\Programs"
        $recoverySource = "\\storage.wpi.edu\dept\Information Technology\CCC\Helpdesk\fc_helpdesk\ShortcutRepair"
        $paths = @{
            Chrome        = "\\$hostname\c$\Program Files\Google\Chrome\Application\chrome.exe";
            Excel         = "\\$hostname\c$\Program Files\Microsoft Office\root\Office16\excel.exe";
            Outlook       = "\\$hostname\c$\Program Files\Microsoft Office\root\Office16\outlook.exe";
            PowerPoint    = "\\$hostname\c$\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE";
            Word          = "\\$hostname\c$\Program Files\Microsoft Office\root\Office16\winword.exe";
            Cherwell      = "\\$hostname\c$\Program Files (x86)\Cherwell Software\Cherwell Service Management\Trebuchet.App.exe";
            Maple         = "\\$hostname\c$\Program Files\Maple 2022\bin.X86_64_WINDOWS\maplew.exe";
            Mathcad       = "\\$hostname\c$\Program Files (x86)\Mathcad\Mathcad 15\mathcad.exe";
            Solstice      = "\\$hostname\c$\Program Files\Mersive\SolsticeClient\SolsticeClient.exe";
            Zoom          = "\\$hostname\c$\Program Files\Zoom\bin\Zoom.exe";
            VLC           = "\\$hostname\c$\Program Files\VideoLAN\VLC\vlc.exe";
            Notepad       = "\\$hostname\c$\Program Files\Notepad++\notepad++.exe";
            Blender2      = "\\$hostname\c$\Program Files\Blender Foundation\Blender 3.2\blender-launcher.exe";
            Blender3      = "\\$hostname\c$\Program Files\Blender Foundation\Blender 3.3\blender-launcher.exe";
            AcrobatDist   = "\\$hostname\c$\Program Files\Adobe\Acrobat DC\Acrobat\acrodist.exe";
            Acrobat       = "\\$hostname\c$\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe";
            AfterEffects  = "\\$hostname\c$\Program Files\Adobe\Adobe After Effects 2023\Support Files\AfterFX.exe";
            Animate       = "\\$hostname\c$\Program Files\Adobe\Adobe Animate 2023\Animate.exe";
            Audition      = "\\$hostname\c$\Program Files\Adobe\Adobe Audition 2023\Adobe Audition.exe";
            Bridge        = "\\$hostname\c$\Program Files\Adobe\Adobe Bridge 2023\Adobe Bridge.exe";
            CharacterRem  = "\\$hostname\c$\Program Files\Adobe\Adobe Character Animator 2023\Support Files\Character Animator.exe";
            AdobeCC       = "\\$hostname\c$\Program Files\Adobe\Adobe Creative Cloud\ACC\Creative Cloud.exe";
            Dreamweaver   = "\\$hostname\c$\Program Files\Adobe\Adobe Dreamweaver 2021\Dreamweaver.exe";
            Illustrator   = "\\$hostname\c$\Program Files\Adobe\Adobe Illustrator 2023\Support Files\Contents\Windows\Illustrator.exe";
            InCopy        = "\\$hostname\c$\Program Files\Adobe\Adobe InCopy 2023\InCopy.exe";
            InDesign      = "\\$hostname\c$\Program Files\Adobe\Adobe InDesign 2023\InDesign.exe";
            Lightroom     = "\\$hostname\c$\Program Files\Adobe\Adobe Lightroom Classic\Lightroom.exe";
            MediaEncoder  = "\\$hostname\c$\Program Files\Adobe\Adobe Media Encoder 2023\Adobe Media Encoder.exe";
            Photoshop     = "\\$hostname\c$\Program Files\Adobe\Adobe Photoshop 2023\Photoshop.exe";
            Prelude       = "\\$hostname\c$\Program Files\Adobe\Adobe Prelude 2022\Adobe Prelude.exe";
            PremierePro   = "\\$hostname\c$\Program Files\Adobe\Adobe Premiere Pro 2023\Adobe Premiere Pro.exe";
            PremiereRush  = "\\$hostname\c$\Program Files\Adobe\Adobe Premiere Rush 2.0\Adobe Premiere Rush.exe";
            AdobeDesign   = "\\$hostname\c$\Program Files\Adobe\Adobe Substance 3D Designer\Adobe Substance 3D Designer.exe";
            AdobeModeler  = "\\$hostname\c$\Program Files\Adobe\Adobe Substance 3D Modeler\Adobe Substance 3D Modeler.exe";
            AdobePainter  = "\\$hostname\c$\Program Files\Adobe\Adobe Substance 3D Painter\Adobe Substance 3D Painter.exe";
            AdobeSampler  = "\\$hostname\c$\Program Files\Adobe\Adobe Substance 3D Sampler\Adobe Substance 3D Sampler.exe";
            AdobeStager   = "\\$hostname\c$\Program Files\Adobe\Adobe Substance 3D Stager\Adobe Substance 3D Stager.exe";
            Dimension     = "\\$hostname\c$\Program Files\Adobe\Adobe Dimension\Dimension.exe";
            Echo360       = "\\$hostname\c$\Program Files\Echo360\UniversalCapture\Echo360 Capture\Echo360 Capture.exe";
            EpicGames     = "\\$hostname\c$\Program Files (x86)\Epic Games\Launcher\Portal\Binaries\Win32\EpicGamesLauncher.exe";
            Keyshot11     = "\\$hostname\c$\WPIAPPS\Keyshot 11\bin\keyshot.exe";
            Max8          = "\\$hostname\c$\Program Files\Cycling '74\Max 8\Max.exe";
            Maya          = "\\$hostname\c$\Program Files\Autodesk\Maya2023\bin\maya.exe";
            VisualStudio  = "\\$hostname\c$\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.exe";
            VSCode        = "\\$hostname\c$\Program Files\Microsoft VS Code\Code.exe";
            Matlab        = "\\$hostname\c$\Program Files\MATLAB\R2022b\bin\matlab.exe";
			VStwo         = "\\$hostname\c$\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe";
			DrRacket      = "\\$hostname\c$\Program Files\Racket\DrRacket.exe";
			Inkscape      = "\\$hostname\c$\Program Files\Inkscape\bin\inkscape.exe";
			ZBrush        = "\\$hostname\c$\Program Files\Maxon ZBrushCoreMini 2021\ZBrushCoreMini.exe"
			
        }
        $shortcutFiles    = @{
            Chrome        = "$recoverySource\Google Chrome.lnk";
            Excel         = "$recoverySource\Excel.lnk";
            Outlook       = "$recoverySource\Outlook.lnk";
            PowerPoint    = "$recoverySource\PowerPoint.lnk";
            Word          = "$recoverySource\Word.lnk";
            Cherwell      = "$recoverySource\Cherwell Service Management.lnk";
            Maple         = "$recoverySource\Maple 2022.lnk";
            Mathcad       = "$recoverySource\Mathcad 15.lnk";
            Solstice      = "$recoverySource\Mersive Solstice.lnk";
            Zoom          = "$recoverySource\Zoom.lnk";
            VLC           = "$recoverySource\VLC media player.lnk";
            Notepad       = "$recoverySource\Notepad++.lnk";
            Blender2      = "$recoverySource\Blender 3.2.lnk";
            Blender3      = "$recoverySource\Blender 3.3.lnk";
            AcrobatDist   = "$recoverySource\Adobe Acrobat Distiller.lnk";
            Acrobat       = "$recoverySource\Adobe Acrobat.lnk";
            AfterEffects  = "$recoverySource\Adobe After Effects 2023.lnk";
            Animate       = "$recoverySource\Adobe Animate 2023.lnk";
            Audition      = "$recoverySource\Adobe Audition 2023.lnk";
            Bridge        = "$recoverySource\Adobe Bridge 2023.lnk";
            CharacterRem  = "$recoverySource\Adobe Character Animator 2023.lnk";
            AdobeCC       = "$recoverySource\Adobe Creative Cloud.lnk";
            Dreamweaver   = "$recoverySource\Adobe Dreamweaver 2021.lnk";
            Illustrator   = "$recoverySource\Adobe Illustrator 2023.lnk";
            InCopy        = "$recoverySource\Adobe InCopy 2023.lnk";
            InDesign      = "$recoverySource\Adobe InDesign 2023.lnk";
            Lightroom     = "$recoverySource\Adobe Lightroom Classic.lnk";
            MediaEncoder  = "$recoverySource\Adobe Media Encoder 2023.lnk";
            Photoshop     = "$recoverySource\Adobe Photoshop 2023.lnk";
            Prelude       = "$recoverySource\Adobe Prelude 2022.lnk";
            PremierePro   = "$recoverySource\Adobe Premiere Pro 2023.lnk";
            PremiereRush  = "$recoverySource\Adobe Premiere Rush.lnk";
            AdobeDesign   = "$recoverySource\Adobe Substance 3D Designer.lnk";
            AdobeModeler  = "$recoverySource\Adobe Substance 3D Modeler.lnk";
            AdobePainter  = "$recoverySource\Adobe Substance 3D Painter.lnk";
            AdobeSampler  = "$recoverySource\Adobe Substance 3D Sampler.lnk";
            AdobeStager   = "$recoverySource\Adobe Substance 3D Stager.lnk";
            Dimension     = "$recoverySource\Dimension.lnk";
            Echo360       = "$recoverySource\Echo360 Universal Capture.lnk";
            EpicGames     = "$recoverySource\Epic Games Launcher.lnk";
            Keyshot11     = "$recoverySource\Keyshot 11.lnk";
            Max8          = "$recoverySource\Max 8 (64-bit).lnk";
            Maya          = "$recoverySource\Maya 2023.lnk";
            VisualStudio  = "$recoverySource\Visual Studio 2019.lnk";
            VSCode        = "$recoverySource\Visual Studio Code.lnk";
            Matlab        = "$recoverySource\MATLAB R2022b.lnk";
			VStwo         = "$recoverySource\Visual Studio 2022.lnk";
			DrRacket      = "$recoverySource\DrRacket.lnk";
			Inkscape      = "$recoverySource\Inkscape.lnk";
			ZBrush        = "$recoverySource\ZBrushCore.lnk"
        }
        $shortcut = @{
            Chrome        = "$StartMenuFolder\Google Chrome.lnk";
            Excel         = "$StartMenuFolder\Excel.lnk";
            Outlook       = "$StartMenuFolder\Outlook.lnk";
            PowerPoint    = "$StartMenuFolder\PowerPoint.lnk";
            Word          = "$StartMenuFolder\Word.lnk";
            Cherwell      = "$StartMenuFolder\Cherwell Service Management\Cherwell Service Management.lnk";
            Maple         = "$StartMenuFolder\Maple 2022.lnk";
            Mathcad       = "\\$hostname\c$\ProgramData\Microsoft\Windows\Start Menu\Mathcad 15.lnk";
            Solstice      = "$StartMenuFolder\Mersive Technologies, Inc\Solstice\Mersive Solstice.lnk";
            Zoom          = "$StartMenuFolder\Zoom\Zoom.lnk";
            VLC           = "$StartMenuFolder\VideoLAN\VLC\VLC media player.lnk";
            Notepad       = "$StartMenuFolder\Notepad++.lnk";
            Blender2      = "$StartMenuFolder\blender\Blender 3.2.lnk";
            Blender3      = "$StartMenuFolder\blender\Blender 3.3.lnk";
            AcrobatDist   = "$StartMenuFolder\Adobe Acrobat Distiller.lnk";
            Acrobat       = "$StartMenuFolder\Adobe Acrobat.lnk";
            AfterEffects  = "$StartMenuFolder\Adobe After Effects 2023.lnk";
            Animate       = "$StartMenuFolder\Adobe Animate 2023.lnk";
            Audition      = "$StartMenuFolder\Adobe Audition 2023.lnk";
            Bridge        = "$StartMenuFolder\Adobe Bridge 2023.lnk";
            CharacterRem  = "$StartMenuFolder\Adobe Character Animator 2023.lnk";
            AdobeCC       = "$StartMenuFolder\Adobe Creative Cloud.lnk";
            Dreamweaver   = "$StartMenuFolder\Adobe Dreamweaver 2021.lnk";
            Illustrator   = "$StartMenuFolder\Adobe Illustrator 2023.lnk";
            InCopy        = "$StartMenuFolder\Adobe InCopy 2023.lnk";
            InDesign      = "$StartMenuFolder\Adobe InDesign 2023.lnk";
            Lightroom     = "$StartMenuFolder\Adobe Lightroom Classic.lnk";
            MediaEncoder  = "$StartMenuFolder\Adobe Media Encoder 2023.lnk";
            Photoshop     = "$StartMenuFolder\Adobe Photoshop 2023.lnk";
            Prelude       = "$StartMenuFolder\Adobe Prelude 2022.lnk";
            PremierePro   = "$StartMenuFolder\Adobe Premiere Pro 2023.lnk";
            PremiereRush  = "$StartMenuFolder\Adobe Premiere Rush.lnk";
            AdobeDesign   = "$StartMenuFolder\Adobe Substance 3D Designer.lnk";
            AdobeModeler  = "$StartMenuFolder\Adobe Substance 3D Modeler.lnk";
            AdobePainter  = "$StartMenuFolder\Adobe Substance 3D Painter.lnk";
            AdobeSampler  = "$StartMenuFolder\Adobe Substance 3D Sampler.lnk";
            AdobeStager   = "$StartMenuFolder\Adobe Substance 3D Stager.lnk";
            Dimension     = "$StartMenuFolder\Dimension.lnk";
            Echo360       = "$StartMenuFolder\Echo360 Universal Capture.lnk";
            EpicGames     = "$StartMenuFolder\Epic Games Launcher.lnk";
            Keyshot11     = "$StartMenuFolder\Keyshot 11\Keyshot 11.lnk";
            Max8          = "$StartMenuFolder\Cycling '74\Max 8\Max 8 (64-bit).lnk";
            Maya          = "$StartMenuFolder\Autodesk Maya 2023\Maya 2023.lnk";
            VisualStudio  = "$StartMenuFolder\Visual Studio 2019.lnk";
            VSCode        = "$StartMenuFolder\Visual Studio Code.lnk";
            Matlab        = "$StartMenuFolder\MATLAB R2022b.lnk";
			VStwo         = "$StartMenuFolder\Visual Studio 2022.lnk";
			DrRacket      = "$StartMenuFolder\DrRacket.lnk";
			Inkscape      = "$StartMenuFolder\Inkscape\Inkscape.lnk";
			ZBrush        = "$StartMenuFolder\Maxon\Maxon ZBrushCoreMini 2021.6.5\ZBrushCore.lnk"
        }
    }
    
    process {
        
            foreach ($key in $paths.keys) {
                "Checking for $key"
                if (!(Test-Path $shortcut.$key)) {
                    "Shortcut missing. Checking for application."
                    if (test-path $paths.$key) {
                        "Application found. Restoring Shortcut."
                        Copy-Item -Path $shortcutFiles.$key -Destination $shortcut.$key
                    }
                }
            }
        
    }
    
    end {
        
    }
}

function Repair-PublicMachineShortcuts {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $publicMachines = get-adcomputer -filter * -SearchBase "OU=PUBLIC,OU=WPIWorkstations,DC=admin,DC=wpi,DC=edu"
    }
    
    process {
        foreach($computer in $publicMachines){
            Repair-ASRShortcuts -hostname $computer.name
        }
    }
    
    end {
        
    }
}