##################################################
#------------------Version 2.1--------------------
#-----------Script von Nicolas Caluori------------
#--------------Amt für Informatik-----------------
#-----------------Januar 2020---------------------
##################################################

########### Verstecken des PowerShell Fensters #############
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0) # Steht für Verstecken
}

Hide-Console #Aufruf der oben definierten Funktion

##############################################
############ Definition Variablen ############

$sync = [Hashtable]::Synchronized(@{}) #Die Variable, welche zwischen allen Runspaces synchronisiert ist
$sync.path = $PSScriptRoot
$sync.list = New-Object System.Collections.ArrayList
$defaultcsvpath = "$($sync.path)\server.csv"
$ServerCsvExists = $false #Existiert die Server.csv Datei?
$script:csvloaded = $false
$sync.status = "standby"
$sync.durchgang = 0
$sync.output = @()
$global:defaultoutputfile = "$($sync.path)\output.csv"

############ Wenn die Server.csv exisitert wird dies festgehalten, dass sie dann geladen wird ############
if (Test-Path -Path $defaultcsvpath){
    $ServerCsvExists = $true 
}
else {
    $ServerCsvExists = $false
}
#######################################################################################
############ Paint Funktion für das einfach erstellen der Form-Komponenten ############

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null 
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null 
[System.Windows.Forms.Application]::EnableVisualStyles()

function paint(${sync.form}, $ctrl, $TablIndex, $name, $Text, $x, $y, $Width, $Height){
    try{$sync.form.Controls.Add($ctrl)                                      }catch{}
    try{$ctrl.TabIndex = $TablIndex                                         }catch{}
    try{$ctrl.Text     = $Text                                              }catch{}
    try{$ctrl.name     = $name                                              }catch{}
    try{$ctrl.Location = New-Object System.Drawing.Size($x,$y)              }catch{}
    try{$ctrl.size     = New-Object System.Drawing.Size($Width,$Height)     }catch{}
    try{$ctrl.DataBindings.DefaultDataSourceUpdateMode = 0                  }catch{}
    try{$ctrl.MinimumSize = New-Object System.Drawing.Size($Width,$Height)  }catch{}
    $ctrl
}

################################################################
############ Windows Form und Komponenten erstellen ############
$sync.form                       = New-Object system.Windows.Forms.Form
$sync.form.ClientSize            = '1250,550'
$sync.form.MinimumSize           = '1300,640'
$sync.form.text                  = "Software Checker"
$sync.form.BackColor             = "#ffffff"
$sync.form.TopMost               = $false

$sync.progressbar1               = paint $sync.form (New-Object system.Windows.Forms.ProgressBar) 0 $null $null 18 540 931 30 
$sync.progressbar1.Maximum       = 100
$sync.progressbar1.Minimum       = 0
$sync.progressbar1.Step          = 5
$sync.progressbar1.Value         = 0
$sync.progressbar1.Style         = "Continuous"
$sync.progressbar1.Anchor        = 'bottom,right,left'

$sync.logbox                     = paint $sync.form (New-Object system.Windows.Forms.TextBox) 0 $null $null 18 70 368 340 
$sync.logbox.multiline           = $true
$sync.logbox.Enabled             = $true
$sync.logbox.ReadOnly            = $true
$sync.logbox.Scrollbars          = "Vertical"
$sync.logbox.Font                = 'Microsoft Sans Serif,10'
$sync.logbox.Anchor              = 'top,bottom,left'

$titel                           = paint $sync.form (New-Object system.Windows.Forms.Label) 0 $null "Software Checker" 530 5 25 10 
$titel.AutoSize                  = $true
$titel.Font                      = 'Microsoft Sans Serif,26'

$datagridtitel1                  = paint $sync.form (New-Object system.Windows.Forms.Label) 0 $null "Eingabe" 481 60 25 10 
$datagridtitel1.AutoSize         = $true
$datagridtitel1.Font             = 'Microsoft Sans Serif,20'

$datagridtitel2                  = paint $sync.form (New-Object system.Windows.Forms.Label) 0 $null "Ausgabe" 850 60 25 10 
$datagridtitel2.AutoSize         = $true
$datagridtitel2.Font             = 'Microsoft Sans Serif,20'

$sync.startbutton                = paint $sync.form (New-Object system.Windows.Forms.Button) 2 $null "Start" 789 493 161 43 
$sync.startbutton.Font           = 'Microsoft Sans Serif,19'
$sync.startbutton.Anchor         = 'bottom,right'

$groupBox                        = paint $sync.form (New-Object System.Windows.Forms.GroupBox) 1 $null "Diese Software prüfen:" 980 400 $null $null 
$groupBox.Anchor                 = 'bottom,right'  
###############################################################
############ Teil für das erstellen der Checkboxen ############

# Die verschiedenen Checkboxen benennen
$checkboxnamen = @()
$checkboxnamen += @{"Name"="Java"}
$checkboxnamen += @{"Name"="PHP"}
$checkboxnamen += @{"Name"="OpenJDK"}
$checkboxnamen += @{"Name"="IIS"}
$checkboxnamen += @{"Name"=".NET"}
$checkboxnamen += @{"Name"="Tomcat"}
$Checkboxes = @()
$y = 20 # Y-Wert der ersten Checkbox

foreach ($name in $checkboxnamen){    
    $Checkbox = New-Object System.Windows.Forms.CheckBox
    $Checkbox.Text = $name.Name
    $Checkbox.Location = New-Object System.Drawing.Size(10,$y) 
    $y += 30 # Erhöht den Y-Wert + 30 (30px tiefer als Checkbox darüber)
    $groupBox.Controls.Add($Checkbox) #Die Checkbox in die Groupbox hinzufügen
    $Checkboxes += $Checkbox #Die Checkbox in das Array Checkboxes hinzufügen --> Wird zum prüfen der Haken gebraucht
}
$groupBox.size = New-Object System.Drawing.Size(200,(35*$checkboxes.Count))  #Die Grösse der GroupBox festlegen 40px pro Checkbox

$checkboxes[5].Add_CheckedChanged({
    if ($Checkboxes[5].Checked){
        $checkboxes[0].Checked = $true
        $Checkboxes[0].Enabled = $false
    }
    elseif (!$Checkboxes[5].Checked){
        $Checkboxes[0].Enabled = $true
    }
})

####################################################
############ Teil fürs Input,Output CSV ############
function Sync-Csv{ #Sync-Csv lädt das Csv neu in das DataGridView
    $script:CsvData = New-Object System.Collections.ArrayList #CsvData auf eine leere Arraylist zurücksetzen
    $script:CsvData.AddRange(@(import-csv $script:csvPath)) #Das CSV importieren und im CsvData speichern
    if ($script:CsvData.Name -eq "" -or $null -eq $script:CsvData.Name){
        Remove-Item -Path $script:csvPath
        "{0}" -f '"Name"' | add-content -path $script:csvPath #Dem CSV wird eine leere Zeile hinzugefügt
        "{0}" -f '" "' | add-content -path $script:csvPath #Dem CSV wird eine leere Zeile hinzugefügt
        $script:CsvData = New-Object System.Collections.ArrayList #CsvData auf eine leere Arraylist zurücksetzen
        $script:CsvData.AddRange(@(import-csv $script:csvPath)) #Das CSV importieren und im CsvData speichern
    }
    $script:dataGridinput.DataSource = $script:CsvData; #Die Daten vom CSV(CsvData) in das DataGridView eintragen
    $sync.form.refresh() #Das Form aktualisieren
}
function Save-ToCsv(){ #DataGridview zu CSV speichern
    if($script:csvloaded -or $script:servercsvexists){
        $script:CsvData | export-csv -NoTypeInformation -path $script:csvPath #Die Werte vom DataGridView in das CSV schreiben
        Sync-Csv #Funktion Sync-Csv
    }
}

#Diese Funktion erstellt die Komponenten für den Input und Output der Daten
function InputOutput($script:csvPath) { 
    #CSV laden
        $script:CsvData = New-Object System.Collections.ArrayList #CsvData definieren
        if ($ServerCsvExists){
            $script:CsvData.AddRange(@(import-csv $script:csvPath)) #Wenn die "server.csv" Datei existiert werden diese Server importiert
        }
        ################# Wenn das Form geladen ist
        $sync.form.add_Load({
            if ($ServerCsvExists){
                $script:dataGridinput.DataSource = $script:CsvData; #Wenn die "server.csv" Datei existiert werden die Werte ins DataGridView eingetragen 
            }
            $sync.dataGridoutput.DataSource = $sync.output; #die Daten vom DataGridView Output sind in der Variable output
            $sync.form.refresh() #Das Form aktualisieren
        }) 
        #Neue Knöpfe für die DataGridViews
        ################# Speichern Knopf
        $buttonloadcsv = paint $sync.form (New-Object System.Windows.Forms.Button) 2 "button1" "Import CSV" 420 380 75 30 
                    $buttonloadcsv.UseVisualStyleBackColor = $True 
                    $buttonloadcsv.add_Click({
                        #Dialog in dem man das CSV auswählen kann
                        $newCSVdialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
                            Multiselect = $false #Man kann nur eine Datei auswählen
                            InitialDirectory = [Environment]::GetFolderPath('MyComputer') #Ordner, indem man startet
                            Filter = 'CSV Dateien (*.csv)|*' #Nur CSV-Dateien aber auch andere anzeigen
                        }
                        $null = $newCSVdialog.ShowDialog() #Zeige den Dialog
                        $newCSV = $newCSVdialog.filename #Das ist der Pfad zur Datei
                        if ($newCSV -ne ""){
                            $sync.logbox.AppendText("Import CSV vom: `r`n$newCSV`r`n")
                            $script:csvloaded = $true #Ein CSV wurde geladen
                            $script:csvPath = $newCSV #Den neuen CSV-Pfad setzen
                            Sync-Csv #Das CSV neu einlesen
                            Save-ToCsv #Einmal speichern um nachher Fehler zu vermeiden
                        }
                    })
        $buttonloadcsv.Anchor = 'bottom,left'
        ################# Export CSV Knopf
        $buttonExportCsv = paint $sync.form (New-Object System.Windows.Forms.Button) 3 "button3" "Export zu CSV" 670 380 100 30 
                    $buttonExportCsv.UseVisualStyleBackColor = $True 
                    $buttonExportCsv.add_Click({ 
                        $exportpath=New-Object System.Windows.Forms.SaveFileDialog #Dialog für das Speichern der Datei 
                        $exportpath.Filter = "Als CSV Datei speichern (*.csv)|*.csv|Als Text Datei speichern (*.txt)|*.txt" # Als CSV oder TXT
                        if($exportpath.ShowDialog() -eq 'Ok'){ #Wenn es keine Fehler gibt
                            Out-ToCsv -dateipfad $($exportpath.filename) -values $sync.output #Daten in das vorher ausgewählte CSV ausgeben (Funktion Out-ToCSv)
                        }
                        
                    })
        $buttonexportcsv.Anchor = 'bottom,left'    
        ################# Neue Zeile Knopf
        $buttonaddrow = paint $sync.form (New-Object System.Windows.Forms.Button) 4 "button2" "Zeile hinzufuegen" 510 380 130 30 
                    $buttonaddrow.UseVisualStyleBackColor = $True 
                    $buttonaddrow.add_Click({
                        if ($script:csvloaded -or $script:servercsvexists){
                            Save-ToCsv #Zuerst ins CSV speichern
                            "{0},{1}" -f '""','' | add-content -path $script:csvPath #Dem CSV wird eine leere Zeile hinzugefügt
                            Sync-Csv #Das CSV wird wieder eingelesen
                        }
                    }) 
        $buttonaddrow.Anchor = 'bottom,left'
        ################# Export CSV Knopf
        $sync.showoutput = paint $sync.Form (New-Object System.Windows.Forms.Button) 3 "button4" "Output anzeigen" 780 380 100 30 
                        $sync.showoutput.UseVisualStyleBackColor = $True 
                        $sync.showoutput.add_Click({ 
                            show-output
                        })  
        $sync.showoutput.Anchor = 'bottom,left' 
        ################# DataGridView 1 (Eingabe)
        $script:dataGridinput                       = paint $sync.form (New-Object System.Windows.Forms.DataGridView) 3 "dataGrid0" $Null 420 100 220 275 
        $script:dataGridinput.RowHeadersVisible     = $false
        $script:dataGridinput.AutoSizeColumnsMode   = 'Fill'
        $script:dataGridinput.selectionmode         = 'FullRowSelect'
        $script:dataGridinput.MultiSelect           = $false
        $script:dataGridinput.Anchor                = 'top,bottom,left'
        $script:dataGridinput.Add_CellValueChanged({Save-ToCsv}) #Sobald man etwas ändert, wird gespeichert

        ################# DataGridView 2 (Ausgabe)
        $sync.dataGridoutput                      = paint $sync.form (New-Object System.Windows.Forms.DataGridView) 4 "dataGrid2" $Null 670 100 600 275
        $sync.dataGridoutput.RowHeadersVisible    = $false
        $sync.dataGridoutput.AutoSizeColumnsMode  = 'Fill'
        $sync.dataGridoutput.selectionmode        = 'FullRowSelect'
        $sync.dataGridoutput.MultiSelect          = $false
        $sync.dataGridoutput.AllowUserToAddRows   = $false
        $sync.dataGridoutput.ReadOnly             = $true
        $sync.dataGridoutput.Anchor               = 'top,right,bottom,left' 
} 
if ($ServerCsvExists){
    $null = (InputOutput $defaultcsvpath) #Die Funktion InputOutput einmal ausführen
    Save-ToCsv #Einmal speichern um nachher Fehler zu vermeiden
}
else {
    $null = (InputOutput) #Wenn kein server.csv existiert wird die Funktion geladen, aber vorerst ohne Dateipfad
}

############################################################
############ Funktion die Java Version ausliest ############
function Get-JavaVersion([string]$thisserver,[int]$durchgang){

    $process = [PowerShell]::Create().AddScript({ ##### Erstellt einen Runspace und fügt ihm ein Scriptblock hinzu
        Param ($thisserver,$durchgang)
        while ($thisserver -ne $sync.server -and $sync.status -eq "running"){ #Wenn dieser Server am Zug ist und das Script nicht gestoppt wurde
            Start-Sleep -Milliseconds 500
            #$sync.logbox.AppendText("Warte.. auf $thisserver `r`n")
        }
        try{
            #Den folgenden Scriptblock auf dem Ziel-Server ausführen und die zurückgegebene Variable in $javaversion speichern
            $javaversion = Invoke-Command -ComputerName $thisserver -ScriptBlock {
                & { 
                    $process = Get-Process -Id $pid
                    $process.PriorityClass = 'BelowNormal' #Priorität tief, sodass andere wichtige Prozesse davor ausgefürt werden 
                    # 2 Zeilen werden dem Array hinzugefügt, da sonst Fehler auftreten
                    $javaversions = @()
                    $javaversions +=,@("Version","Pfad")
                    $javaversions +=,@("Version","Pfad")
                    $dirs = @()
                    foreach ($drive in $([System.IO.DriveInfo]::getdrives() | Where-Object {$_.DriveType -eq 'Fixed'})){
                        $letter = $drive.RootDirectory.Name
                        $dirs += Get-ChildItem "$letter" -Filter "*jre*" | Select-Object -ExpandProperty FullName
                        $dirs += Get-ChildItem "$letter*" -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -ne "Windows" -and $_.Name -ne "Users"} | Select-Object -ExpandProperty FullName |
                            Foreach-Object {
                                Get-ChildItem "$_\" -Directory "*jre*" -Recurse | Select-Object -ExpandProperty FullName
                            }
                        foreach($dir in $dirs){ #Für jeden "jre" Ordner
                            # 3 Stufen tief nach "java.exe" suchen
                            $javaexes = Get-ChildItem "$dir\" -Recurse | Where-Object Name -Like "*java.exe*" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
                            #Jede Java.exe mit dem Parameter -Version ausführen
                            foreach($javaexe in $javaexes){
                                $Parms = "-version" #Parameter
                                [string]$thisversion = ((& "$javaexe" $Parms 2>&1)[0]) #Befehl ausführen
                                [string]$thisversion = $thisversion | %{$_.split('"')[1]} #Das Ergebnis ist das zwischen den ""
                                if ($null -ne $thisversion){
                                    $javaversions +=,@($thisversion,$javaexe)
                                    $thisversion = $null
                                }
                            }
                        }
                    }
                    return $javaversions
                }
            }
            if ($javaversion -ne "" -and $null-ne $javaversion -and $javaversion.Count -gt 2){
                foreach ($version in $javaversion){
                    if ($version[0] -ne "Version" -and $version[0] -ne "V" -and $version[0] -ne "P"){ #Die ersten zwei Werte ignorieren
                        if ($durchgang -eq $sync.durchgang){
                            $sync.output +=,@($thisserver,"Java",$version[0],$version[1])
                            $sync.("java$thisserver") = $version[1] #Die globale Javaexe setzen für den Tomcat
                        }
                    }
                }
            }
            else {
                $sync.logbox.AppendText("${thisserver}: Kein Java`r`n")
            }
        }
        Catch {
            $sync.logbox.AppendText("${thisserver}: Kein Java(Fehler)`r`n")
        }
        finally{
            #Wird ausgeführt egal ob es einen Fehler gibt oder nicht
            $sync.("javadone$thisserver") = $true #Die globale Javaexe setzen für den Tomcat
            $value = $sync.progressbar1.Value + $($sync.progressbarstep )
            if ($value -le 95){
                $sync.progressbar1.Value = $value
            }
        }
    }).AddArgument($thisserver).AddArgument($durchgang) # Die Parameter vom Server und dem Durchgang hinzufügen

    $process.RunspacePool = $RSP #Runspace dem Runspacepool hinzufügen
    $sync.list.Add(([pscustomobject]@{ #Der Runspace der Liste hinzufügen
        PowerShell = $Process #Die Variable PowerShell steht für den Runspace
        Handle = $Process.BeginInvoke() #Startet den Runspace, die Variable Handle steht für den Status des Runspaces
    }))        
}

############################################################
############ Funktion die PHP Version ausliest #############
function Get-PHPVersion([string]$thisserver,[int]$durchgang){
    
    $process = [PowerShell]::Create().AddScript({ ##### Erstellt einen Runspace und fügt ihm ein Scriptblock hinzu
        Param ($thisserver,$durchgang)
        while ($thisserver -ne $sync.server -and $sync.status -eq "running"){ #Wenn dieser Server am Zug ist und das Script nicht gestoppt wurde
            Start-Sleep -Milliseconds 500
            #$sync.logbox.AppendText("Warte.. auf $thisserver `r`n")
        }
        try{
            #Den folgenden Scriptblock auf dem Ziel-Server ausführen und die zurückgegebene Variable in $phpversion speichern
            $phpversion = Invoke-Command -ComputerName $thisserver -ScriptBlock {
                & { 
                    $process = Get-Process -Id $pid
                    $process.PriorityClass = 'BelowNormal' #Priorität tief, sodass andere wichtige Prozesse davor ausgefürt werden 
                    # 2 Zeilen werden dem Array hinzugefügt, da sonst Fehler auftreten
                    $phpversions = @()
                    $phpversions +=,@("Version","Pfad")
                    $phpversions +=,@("Version","Pfad")
                    $dirs = @()
                    foreach ($drive in $([System.IO.DriveInfo]::getdrives() | Where-Object {$_.DriveType -eq 'Fixed'})){
                        $letter = $drive.RootDirectory.Name
                        $dirs += Get-ChildItem "$letter" -Filter "*php*" | Select-Object -ExpandProperty FullName
                        $dirs += Get-ChildItem "$letter*" -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -ne "Windows" -and $_.Name -ne "Users"} | Select-Object -ExpandProperty FullName |
                            Foreach-Object {
                                Get-ChildItem "$_\" -Directory "*php*" -Recurse | Select-Object -ExpandProperty FullName
                            }
                        foreach($dir in $dirs) { #Für jeden "php" Ordner
                            # 3 Stufen tief nach "php.exe" suchen
                            $phpexes = Get-ChildItem "$dir\" -Recurse  | Where-Object Name -Like "*php.exe*" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
                            
                            foreach($phpexe in $phpexes){
                                if (Test-Path -Path $phpexe){
                                    #Alte Variante --> #Jede php.exe mit dem Parameter -Version ausführen
                                    <#$Parms = "-version"
                                    $thisversion = (($(& "$phpexe" $Parms)[0].SubString(3,6)).trim())
                                    #>
                                    $phpproperty = Get-ItemProperty -Path $phpexe 
                                    $thisversion = $phpproperty.VersionInfo.ProductVersion 
                                    if ($null-ne $thisversion){
                                        $phpversions +=,@($thisversion,$phpexe)
                                        $thisversion = $null
                                    }
                                }
                            }
                        }
                    }
                    return $phpversions
                }
            }
            if ($phpversion -ne "" -and $null -ne $phpversion -and $phpversion.Count -gt 2){
                foreach ($version in $phpversion){
                    if ($version[0] -ne "Version" -and $version[0] -ne "V" -and $version[0] -ne "P"){ #Die ersten 2 Werte ignorieren
                        if ($durchgang -eq $sync.durchgang){
                            $sync.output +=,@($thisserver,"PHP",$version[0],$version[1])
                        }
                        
                    }
                }
            }
            else {
                $sync.logbox.AppendText("${thisserver}: Kein PHP `r`n")
            }
        }
        Catch {
            $sync.logbox.AppendText("${thisserver}: Kein PHP(Fehler)`r`n")
        }
        finally{
            #Wird ausgeführt egal ob es einen Fehler gibt oder nicht
            $value = $sync.progressbar1.Value + $($sync.progressbarstep )
            if ($value -le 95){
                $sync.progressbar1.Value = $value
            }
        }
    }).AddArgument($thisserver).AddArgument($durchgang) # Die Parameter vom Server und dem Durchgang hinzufügen
        
    $process.RunspacePool = $RSP #Runspace dem Runspacepool hinzufügen
    $sync.list.Add(([pscustomobject]@{ #Der Runspace der Liste hinzufügen
        PowerShell = $Process #Die Variable PowerShell steht für den Runspace
        Handle = $Process.BeginInvoke() #Startet den Runspace, die Variable Handle steht für den Status des Runspaces
    }))
}

############################################################
############ Funktion die JDK Version ausliest #############
function Get-JDKVersion([string]$thisserver,[int]$durchgang){
    
    $process = [PowerShell]::Create().AddScript({ ##### Erstellt einen Runspace und fügt ihm ein Scriptblock hinzu
        Param ($thisserver,$durchgang)
        while ($thisserver -ne $sync.server -and $sync.status -eq "running"){ #Wenn dieser Server am Zug ist und das Script nicht gestoppt wurde
            Start-Sleep -Milliseconds 500
            #$sync.logbox.AppendText("Warte.. auf $thisserver `r`n")
        }
        try{
            #Den folgenden Scriptblock auf dem Ziel-Server ausführen und die zurückgegebene Variable in $jdkversion speichern
            $jdkversion = Invoke-Command -ComputerName $thisserver -ScriptBlock {
                & { 
                    $process = Get-Process -Id $pid
                    $process.PriorityClass = 'BelowNormal' #Priorität tief, sodass andere wichtige Prozesse davor ausgefürt werden 
                    # 2 Zeilen werden dem Array hinzugefügt, da sonst Fehler auftreten
                    $jdkversions = @()
                    $jdkversions +=,@("Version","Pfad")
                    $jdkversions +=,@("Version","Pfad")
                    $dirs = @()
                    foreach ($drive in $([System.IO.DriveInfo]::getdrives() | Where-Object {$_.DriveType -eq 'Fixed'})){
                        $letter = $drive.RootDirectory.Name
                        $dirs += Get-ChildItem "$letter" -Filter "*jdk*" | Select-Object -ExpandProperty FullName
                        $dirs += Get-ChildItem "$letter*" -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -ne "Windows" -and $_.Name -ne "Users"} | Select-Object -ExpandProperty FullName |
                            Foreach-Object {
                                Get-ChildItem "$_\" -Directory "*jdk*" -Recurse | Select-Object -ExpandProperty FullName
                            }
                        foreach($dir in $dirs) { #Für jeden "jdk" Ordner
                            # 3 Stufen tief nach "java.exe" suchen
                            $jdkexes = Get-ChildItem "$dir\" -Recurse | Where-Object Name -Like "*java.exe*" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
                            #Jede Java.exe mit dem Parameter -Version ausführen
                            foreach($jdkexe in $jdkexes){
                                $Parms = "-version" #Parameter
                                [string]$thisversion = ((& "$jdkexe" $Parms 2>&1)[0]) #Befehl ausführen
                                [string]$thisversion = $thisversion | %{$_.split('"')[1]} #Das Ergebnis ist das zwischen den ""
                                if ($null -ne $thisversion){
                                    $jdkversions +=,@($thisversion,$jdkexe)
                                    $thisversion = $null
                                }
                            }
                        }
                    }
                    return $jdkversions
                }
            }

            
            if ($jdkversion -ne "" -and $null -ne $jdkversion -and $jdkversion.Count -gt 2){
                foreach ($version in $jdkversion){
                    if ($version[0] -ne "Version" -and $version[0] -ne "V" -and $version[0] -ne "P"){
                        if ($durchgang -eq $sync.durchgang){
                            $sync.output +=,@($thisserver,"OpenJDK",$version[0],$version[1])
                        }
                        
                    }
                }
            }
            else {
                $sync.logbox.AppendText("${thisserver}: Kein OpenJDK`r`n")
            }    
        }
        Catch {
            $sync.logbox.AppendText("${thisserver}: Kein OpenJDK (Fehler)`r`n")
        }
        finally{
            #Wird ausgeführt egal ob es einen Fehler gibt oder nicht
            $value = $sync.progressbar1.Value + $($sync.progressbarstep )
            if ($value -le 95){
                $sync.progressbar1.Value = $value
            }
        }
    }).AddArgument($thisserver).AddArgument($durchgang) # Die Parameter vom Server und dem Durchgang hinzufügen
  
    $process.RunspacePool = $RSP #Runspace dem Runspacepool hinzufügen
    $sync.list.Add(([pscustomobject]@{ #Der Runspace der Liste hinzufügen
        PowerShell = $Process #Die Variable PowerShell steht für den Runspace
        Handle = $Process.BeginInvoke() #Startet den Runspace, die Variable Handle steht für den Status des Runspaces
    }))
}

############################################################
############ Funktion die IIS Version ausliest #############
function Get-IISVersion([string]$thisserver,[int]$durchgang){
    
    $process = [PowerShell]::Create().AddScript({ ##### Erstellt einen Runspace und fügt ihm ein Scriptblock hinzu
        Param ($thisserver,$durchgang)
        while ($thisserver -ne $sync.server -and $sync.status -eq "running"){ #Wenn dieser Server am Zug ist und das Script nicht gestoppt wurde
            Start-Sleep -Milliseconds 500
            #$sync.logbox.AppendText("Warte.. auf $thisserver `r`n")
        }
        try{
            #Den folgenden Scriptblock auf dem Ziel-Server ausführen und die zurückgegebene Variable in $IISVersion speichern
            $IISVersion = Invoke-Command -ComputerName $thisserver -ScriptBlock {
                & { 
                    $process = Get-Process -Id $pid
                    $process.PriorityClass = 'BelowNormal' #Priorität tief, sodass andere wichtige Prozesse davor ausgefürt werden 
                    $w3wpPath = $Env:WinDir + "\System32\inetsrv\w3wp.exe" 
                    if(Test-Path $w3wpPath) { 
                        $productProperty = Get-ItemProperty -Path $w3wpPath 
                        $IISVersion1 = $productProperty.VersionInfo.ProductVersion 
                        return $IISVersion1 #Die Version zurückgeben
                    }
                }
            }
            if ($IISVersion -ne "" -and $null -ne $IISVersion){
                if ($durchgang -eq $sync.durchgang){
                    $sync.output +=,@($thisserver,"IIS",$IISVersion,"-")
                }
            }
            else {
                $sync.logbox.AppendText("${thisserver}: Kein IIS `r`n")
            }
        }
        Catch {
            $sync.logbox.AppendText("${thisserver}: Kein IIS(Fehler)`r`n")
        }
        finally{
            #Wird ausgeführt egal ob es einen Fehler gibt oder nicht
            $value = $sync.progressbar1.Value + $($sync.progressbarstep )
            if ($value -le 95){
                $sync.progressbar1.Value = $value
            }
        }
    }).AddArgument($thisserver).AddArgument($durchgang) # Die Parameter vom Server und dem Durchgang hinzufügen

    $process.RunspacePool = $RSP #Runspace dem Runspacepool hinzufügen
    $sync.list.Add(([pscustomobject]@{ #Der Runspace der Liste hinzufügen
        PowerShell = $Process #Die Variable PowerShell steht für den Runspace
        Handle = $Process.BeginInvoke() #Startet den Runspace, die Variable Handle steht für den Status des Runspaces
    }))
}

############################################################
############ Funktion die .Net Version ausliest #############

function Get-DotNetVersion([string]$thisserver,[int]$durchgang){
    
    $process = [PowerShell]::Create().AddScript({ ##### Erstellt einen Runspace und fügt ihm ein Scriptblock hinzu
        Param ($thisserver,$durchgang)
        while ($thisserver -ne $sync.server -and $sync.status -eq "running"){ #Wenn dieser Server am Zug ist und das Script nicht gestoppt wurde
            Start-Sleep -Milliseconds 500
            #$sync.logbox.AppendText("Warte.. auf $thisserver `r`n")
        }
        try{
            #Den folgenden Scriptblock auf dem Ziel-Server ausführen und die zurückgegebene Variable in $DotNetVersion speichern
            $DotNetVersion = Invoke-Command -ComputerName $thisserver -ScriptBlock {
                & { 
                    $process = Get-Process -Id $pid
                    $process.PriorityClass = 'BelowNormal' #Priorität tief, sodass andere wichtige Prozesse davor ausgefürt werden 
                    $DotNetVersions = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse |
                    Get-ItemProperty -name Version, Release -EA 0 | 
                    Where-Object { $_.PSChildName -eq "Full"} | Select-Object -ExpandProperty Release
                    return $DotNetVersions
                }
            }
            #CSV mit den releases importieren und in ein Hashtable schreiben
            $sync.dotnetcsv = (Import-Csv -Path $sync.dotnetcsvpath -Delimiter ";")
            $Lookup=@{}
            foreach($row in $sync.dotnetcsv)
            {
                $Lookup[($row.Release).Trim()]=($row.Version).Trim()
            }
            #Alle Versionen in den Output schreiben
            if ($null -ne $DotNetVersion -and $DotNetVersion -ne ""){
                foreach ($version in $DotNetVersion){
                    if ($null -ne $($Lookup.Get_Item("$version"))){
                        if ($durchgang -eq $sync.durchgang){
                            $sync.output +=,@($thisserver,".NET",$($Lookup.Get_Item("$version")),"-")

                        }
                    }
                    else {
                        if ($durchgang -eq $sync.durchgang){
                            $sync.output +=,@($thisserver,".NET","Release $version","-")  
                        }
                        $sync.logbox.AppendText(".Net Release ${version}: Version unbekannt`r`n")
                    }
                }
            }
            else {
                $sync.logbox.AppendText("${thisserver}: Kein .Net`r`n")
            }
        }
        Catch {
            $sync.logbox.AppendText("${thisserver}: Kein .Net (Fehler)`r`n")
        }
        finally{
            #Wird ausgeführt egal ob es einen Fehler gibt oder nicht
            $value = $sync.progressbar1.Value + $($sync.progressbarstep )
            if ($value -le 95){
                $sync.progressbar1.Value = $value
            }
        }
    }).AddArgument($thisserver).AddArgument($durchgang) # Die Parameter vom Server und dem Durchgang hinzufügen

    $process.RunspacePool = $RSP #Runspace dem Runspacepool hinzufügen
    $sync.list.Add(([pscustomobject]@{ #Der Runspace der Liste hinzufügen
        PowerShell = $Process #Die Variable PowerShell steht für den Runspace
        Handle = $Process.BeginInvoke() #Startet den Runspace, die Variable Handle steht für den Status des Runspaces
    }))
}

############################################################
############ Funktion die TomCat Version ausliest ############
function Get-TomCatVersion([string]$thisserver,[int]$durchgang){
    
    $process = [PowerShell]::Create().AddScript({ ##### Erstellt einen Runspace und fügt ihm ein Scriptblock hinzu
        Param ($thisserver,$durchgang)
        while ($thisserver -ne $sync.server -and $sync.status -eq "running"){ #Wenn dieser Server am Zug ist und das Script nicht gestoppt wurde
            Start-Sleep -Milliseconds 500
            #$sync.logbox.AppendText("Warte.. auf $thisserver `r`n")
        }
        try{
            while ($sync.("javadone$thisserver") -ne $true){
                Start-Sleep -Milliseconds 500
                #$sync.logbox.AppendText("Warte $thisserver" + $sync.("javadone$thisserver") +"`r`n")
            }
            $javaexe = $sync.("java${thisserver}")
            if ($null -eq $javaexe -or $javaexe -eq ""){
                return 1
            }
            #Den folgenden Scriptblock auf dem Ziel-Server ausführen und die zurückgegebene Variable in $javaversion speichern
            $tomcatversion = Invoke-Command -ComputerName $thisserver -ScriptBlock {
                    # 2 Zeilen werden dem Array hinzugefügt, da sonst Fehler auftreten
                    $tomcatversions = @()
                    $tomcatversions +=,@("Version","Pfad")
                    $tomcatversions +=,@("Version","Pfad")
                    $dirs = @()
                    foreach ($drive in $([System.IO.DriveInfo]::getdrives() | Where-Object {$_.DriveType -eq 'Fixed'})){
                        $letter = $drive.RootDirectory.Name
                        $dirs += Get-ChildItem "$letter" -Filter "*tomcat*" | Select-Object -ExpandProperty FullName
                        $dirs += Get-ChildItem "$letter*" -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -ne "Windows" -and $_.Name -ne "Users"} | Select-Object -ExpandProperty FullName
                        foreach($dir in $dirs)
                        {
                            # 3 Stufen tief nach "catalina.jar" suchen
                            $tomcatexes = Get-ChildItem "$dir\" "*catalina.jar*" -Recurse | Select-Object -ExpandProperty FullName
                            foreach($tomcatexe in $tomcatexes){
                                $javaexe = $args[0]
                                if ($null -ne $javaexe -and $javaexe -ne ""){
                                    $thisversion = (& $javaexe -cp "$tomcatexe" org.apache.catalina.util.ServerInfo) #Befehl ausführen
                                    $thisversion = $thisversion[2].Split(':')[1].Trim() #Die Werte auslesen und so kürzen, dass nur noch die Version übrig bleibt (alle Werte nach dem ':')
                                    if ($null -ne $thisversion){
                                        $tomcatversions +=,@($thisversion,$tomcatexe)
                                        $thisversion = $null
                                    }
                                }
                            }
                        }
                    }
                    return $tomcatversions
                
            } -ArgumentList $javaexe #Parameter mitgeben --> Javaexe um die Version von Tomcat herauszufinden

            $sync.($thisserver) = $null
            if ($tomcatversion -ne "" -and $null-ne $tomcatversion -and $tomcatversion.Count -gt 2){
                foreach ($version in $tomcatversion){
                    if ($version[0] -ne "Version" -and $version[0] -ne "V" -and $version[0] -ne "P"){ #Die ersten zwei Werte ignorieren
                        if ($durchgang -eq $sync.durchgang){
                            $sync.output +=,@($thisserver,"Tomcat",$version[0],$version[1])
                        }  
                    }
                }
            }
            else {
                $sync.logbox.AppendText("${thisserver}: Kein Tomcat `r`n")
            }
        }
        Catch{
            $sync.logbox.AppendText("${thisserver}: Kein Tomcat(Fehler)`r`n")
        }
        finally{
            #Wird ausgeführt egal ob es einen Fehler gibt oder nicht
            $value = $sync.progressbar1.Value + $($sync.progressbarstep )
            if ($value -le 95){
                $sync.progressbar1.Value = $value
            }
        }

    }).AddArgument($thisserver).AddArgument($durchgang) # Die Parameter vom Server und dem Durchgang hinzufügen
    $process.RunspacePool = $RSP #Runspace dem Runspacepool hinzufügen
    $sync.list.Add(([pscustomobject]@{ #Der Runspace der Liste hinzufügen
        PowerShell = $Process #Die Variable PowerShell steht für den Runspace
        Handle = $Process.BeginInvoke() #Startet den Runspace, die Variable Handle steht für den Status des Runspaces
    }))      
}

################################################################################
############ Funktion die ein Array in ein CSV exportiert ausliest #############
function Out-ToCsv([string]$dateipfad=$global:defaultoutputfile,$values=$sync.output){
    $sync.logbox.AppendText("Export CSV nach: `r`n$dateipfad `r`n")
    #Bestehende CSV löschen falls sie existiert
    if (Test-Path -Path $dateipfad){
        Remove-Item $dateipfad
    }

    #Daten in CSV ausgeben
    $values = $values | Get-Unique
    $values = $values | Select-Object -Unique #Nur eindeutige Werte behalten (Duplikate verwerfen)
    foreach($item1 in $values) { 
        $csv_string = "";
        foreach($item in $item1){
            $csv_string = $csv_string + $item +([char]9) + ";"; #Char9 ist ein Tab --> Dass Nummern nicht als Datum formatiert werden
        }
        Add-Content $dateipfad $csv_string;
    }
}
###########################################################
############ Funktion, die den Output anzeigt #############
function show-output {
    $dateipfad = "$($sync.path)\tmpoutput.csv" #Eine temporäre Datei fürs Anzeigen im DataGridView
    if (Test-Path -Path $dateipfad){ 
        Remove-Item $dateipfad #Lösche die Datei falls sie existiert
    }
    $thisoutput = $sync.output | Get-Unique
    $thisoutput = $thisoutput | Select-Object -Unique
    $thisoutput +=,@("-","-","-","-") #Temporäre Zeile hinzufügen

    #Output in das CSV schreiben
    foreach($item1 in $thisoutput) { 
        $csv_string = "";
        if ($item1[0] -ne "-"){ #Oben erstellte Zeile rausfiltern
            foreach($item in $item1){
                $csv_string = $csv_string + $item + ",";
            }
        }
        Add-Content $dateipfad $csv_string;
    }

    #Daten aus dem CSV einlesen und im DataGridView anzeigen
    $outputdatagridvalues = $null
    $outputdatagridvalues = New-Object System.Collections.ArrayList
    if (Test-Path $dateipfad){
        $outputdatagridvalues.AddRange(@(Import-Csv $dateipfad))  
        $sync.dataGridoutput.DataSource = $outputdatagridvalues;
        $sync.Form.refresh()
    }

    if (Test-Path -Path $dateipfad){
        Remove-Item $dateipfad #CSV Datei wieder löschen
    }
}
#########################################################################
############ Funktion, die die Software der Server ausliest #############
function Get-Software(){
    foreach ($servername in $global:servers){
        $global:server = ($servername).trim()

        if ($server -ne ""){
            if (!(Test-Connection -Computername $server -Count 1 -Quiet)) { #Wenn der Ping fehlschlägt dann 
                $sync.logbox.AppendText("Verbindung zu $server fehlgeschlagen`r`n")
                Move-Progressbar -anzahl $global:checkboxcounter
                $sync.failedconnection++
            }
            #Wenn der Server Online ist wird die Software abgerufen
            Else {
                Try {
                    #Java Version auslesen
                    if ($checkboxes[0].Checked -or $checkboxes[5].Checked){
                        Get-JavaVersion -thisserver $server -durchgang $sync.durchgang
                    }
                    #PHP Version auslesen
                    if ($checkboxes[1].Checked){
                        Get-PHPVersion -thisserver $server -durchgang $sync.durchgang
                    }
                    #OpenJDK Version auslesen
                    if ($checkboxes[2].Checked){
                        Get-JDKVersion -thisserver $server -durchgang $sync.durchgang
                    }
                    #IIS Version auslesen
                    if ($checkboxes[3].Checked){
                        Get-IISVersion -thisserver $server -durchgang $sync.durchgang
                    }
                    #.Net Version auslesen
                    if ($checkboxes[4].Checked){
                        Get-DotNetVersion -thisserver $server -durchgang $sync.durchgang
                    }
                    #TomCat Version auslesen
                    if ($checkboxes[5].Checked){
                        Get-TomCatVersion -thisserver $server -durchgang $sync.durchgang
                    }
                }
                Catch {
                    #Beim Fehler mach das
                    $sync.logbox.AppendText("########################`r`n Es gab einen schwerwiegenden Fehler!`r`n########################`r`n")
                    [System.Windows.Forms.MessageBox]::Show("Es gab einen Fehler!.","Fehler!",0,[System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }        
        }

        #Wenn der Server = "" ist drei Schritte weiter --> Der Server existiert nicht
        else {
            Move-Progressbar -anzahl $global:checkboxcounter
        }
    }
}
function Set-ProgressbarStep{
    $i = 0
    foreach($srv in $global:servers.name){
        if ($srv -ne ""){
            $i++
        }
    }
    if ($i -eq 0){
        $i = 5
    }
    if ($global:checkboxcounter -ne 0){
        $sync.progressbarstep = (100/$i/$global:checkboxcounter) #Der Schritt der Progressbar ist so, dass es am Ende auf 100% ist. $i = Anzahl Server; $Checkboxcounter = Anzahl Software --> Nach jeder Software 1 step
    }
}
function Move-Progressbar([int]$anzahl=1){
    $value = $sync.progressbar1.Value + $($sync.progressbarstep * $anzahl)
    if ($value -le 95){
        $sync.progressbar1.Value = $value
    }
}
#################################################################################
############ Funktion die beim Click vom Startknopf ausgeführt wird #############

$sync.startbutton.Add_Click({
    
    if ($sync.status -eq "standby"){
        if ($ServerCsvExists -or $script:csvloaded){
            $sync.startbutton.Text = "Stop"
            $sync.status = "running"
            $sync.durchgang++
            $sync.failedconnection = 0

            Save-ToCsv ###### DataGridView Input ins CSV speichern 

            ######### Definition der Variablen #########
            $sync.output = $null #Output zurücksetzen
            $sync.output = @() #Output ist ein Array
            $sync.output +=,@("Servername","Software","Version","Pfad") #Die Header zum Output hinzufügen
            $sync.logbox.Text="" #Die Logbox zurücksetzen
            $sync.progressbar1.Value = 0 #Progressbar auf 0 setzen
            $global:checkboxcounter = 0 #Die Anzahl Checkboxes
            $global:servers = Import-Csv -Path $script:csvpath #Die Server einlesen
            $global:servers = $global:servers.name | Select-Object -unique
            $sync.servers = $global:servers
            $sync.dotnetcsvpath = "$($sync.path)\dotNetVersions.csv" #Daten fürs CSV: https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed
            $sync.stopwatch =  [system.diagnostics.stopwatch]::StartNew() #Stopuhr starten
            ############################################

            ############################################ RunSpacePool erstellen ############################################

            #Den Sessionstate erstellen und die Variable $sync hinzufügen
            $Variable = New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'sync',$sync,$Null
            $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

            #Die Variable zum Sessionstate hinzufügen
            $InitialSessionState.Variables.Add($Variable)
            
            #Den Runspacepool mit der SessionStateVariable erstellen
            $RSP = [runspacefactory]::CreateRunspacePool(1,10,$InitialSessionState,$Host)
            $RSP.Open()

            #################################################################################################################
            #Zählen wie viel Checkboxen angehakt wurden
            foreach ($checkbox in $checkboxes){
                if ($checkbox.Checked){
                    $global:checkboxcounter++
                }
            }
            #Wenn nach DotNet gesucht werden soll wird zuerst geprüft ob die VerweisDatei "dotNetVersions.csv" existiert
            if ($checkboxes[4]){
                if (!(Test-Path -Path $sync.dotnetcsvpath)){
                    [System.Windows.Forms.MessageBox]::Show("Kein dotNetVersions.csv gefunden. Leere Datei wird erstellt","Kein dotNetVersions.csv",0,[System.Windows.Forms.MessageBoxIcon]::Error)
                    "{0}" -f '"Release";"Version"' | add-content -path $sync.dotnetcsvpath #Die Datei wird erstellt
                }
            }

            Set-ProgressbarStep #Den Step für die Progressbar setzen
            #Wenn mindestens eine Checkbox ausgewählt wurde
            if ($global:checkboxcounter -gt 0){
                $sync.checkboxcounter = $global:checkboxcounter
                $process = [PowerShell]::Create().AddScript({ ##### Erstellt einen Runspace und fügt ihm ein Scriptblock hinzu
                    $sync.server = $sync.servers[0].Trim()
                    $i = 1
                    if ($sync.servers.Count -le 1){
                        $sync.server = $sync.servers.Trim()
                        $sync.logbox.Appendtext("Software von $($sync.server) wird ausgelesen`r`n")
                    }
                    else {
                        $sync.logbox.Appendtext("Software von $($sync.server) wird ausgelesen`r`n")
                        while ($i -lt $sync.servers.Count - $sync.failedconnection -and $sync.servers.Count -gt 1){
                            if (($sync.list.handle.iscompleted | Where-Object {$_ -eq $true}).Count + $sync.failedconnection * $sync.checkboxcounter -ge $i * $sync.checkboxcounter){
                                $sync.server = $sync.servers[$i].Trim()
                                if ($sync.server -ne "" -and $null -ne $sync.server){
                                    $sync.logbox.Appendtext("Software von $($sync.server) wird ausgelesen`r`n")
                                }
                                $i++
                            }
                            Start-Sleep -Milliseconds 1000
                        }
                    }
                })
                $process.RunspacePool = $RSP #Runspace dem Runspacepool hinzufügen
                $sync.list.Add(([pscustomobject]@{ #Der Runspace der Liste hinzufügen
                    PowerShell = $Process #Die Variable PowerShell steht für den Runspace
                    Handle = $Process.BeginInvoke() #Startet den Runspace, die Variable Handle steht für den Status des Runspaces
                }))

                ########################################################
                Get-Software ####### Funktion zur Auslesung der Software
                ########################################################
                
                ###########################################################################
                # Prozess welcher wartet, bis das Script fertig ist und es dann abschliesst
                ###########################################################################

                $sync.processid = (Get-Process -id $pid).Path
                $process = [PowerShell]::Create().AddScript({ ##### Erstellt einen Runspace und fügt ihm ein Scriptblock hinzu
                        
                        #Warte bis alle Prozesse abgeschlossen sind, ausser einer (dieser Prozess) | Oder wenn der Status nicht mehr running ist, also das Script gestopt wurde
                        while (($sync.list.handle.iscompleted | Where-Object {$_ -eq $true}).Count -ne ( $sync.list.handle.iscompleted.count -1) -and $sync.status -eq "running"){
                            Start-Sleep -Milliseconds 500
                        }
                        $sync.progressbar1.Value = 100 #Die Progressbar auf 100 setzen, da die Aktion beendet ist.
                        $sync.stopwatch.Stop() #Stopuhr stoppen

                        #Wenn mehr als 30 Sekunden gebraucht wurden wird eine Benachrichtigung gesendet, dass die Abfrage beendete ist
                        if ($sync.stopwatch.Elapsed.TotalSeconds -gt 30){
                            $balmsg = New-Object System.Windows.Forms.NotifyIcon
                            $balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($sync.processid)
                            $balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
                            $balmsg.BalloonTipText = 'Der Software-Check wurde abgeschlossen'
                            $balmsg.BalloonTipTitle = "Software-Script beendet"
                            $balmsg.Visible = $true
                            $balmsg.ShowBalloonTip(1000)
                        }
                        #Den Benutzer informieren, dass das Script fertig ist
                        $sync.logbox.AppendText("Fertig!`r`n")
                        $sync.logbox.AppendText("Der Vorgang dauerte: `r`n `t$($sync.stopwatch.Elapsed.Minutes) Minuten`r`n `t$([math]::Round($sync.stopwatch.Elapsed.Seconds,0)) Sekunden`r`n")
                        $sync.logbox.AppendText("#####################################`r`n")
                        $sync.startbutton.Text = "Start"
                        $sync.status = "standby"
                        $RSP.Close() #Den Runspacepool schliessen
                        $RSP.Dispose() #Den Runspacepool auflösen
                        $list = $sync.list #Die List der Runspaces temporär speichern
                        $sync.list = New-Object System.Collections.ArrayList #Die Liste der Runspaces zurücksetzen
                        ################ Alle Runspaces stoppen und löschen ################
                        foreach($runspace in $list){
                            $runspace.powershell.EndInvoke()
                            $runspace.powershell.Dispose()
                        }
                    })
                    $process.RunspacePool = $RSP #Runspace dem Runspacepool hinzufügen
                    $sync.list.Add(([pscustomobject]@{ #Der Runspace der Liste hinzufügen
                        PowerShell = $Process #Die Variable PowerShell steht für den Runspace
                        Handle = $Process.BeginInvoke() #Startet den Runspace, die Variable Handle steht für den Status des Runspaces
                    }))
            }
            #Wenn keine Checkbox ausgewählt wurde
            else {
                [System.Windows.Forms.MessageBox]::Show("Es wurde keine Software ausgewählt, also kann auch nicht danach gesucht werden.","Keine Software ausgewählt",0,[System.Windows.Forms.MessageBoxIcon]::Error)
                $sync.startbutton.Text = "Start"
                $sync.status = "standby"
            }   
        }
        else{
            [System.Windows.Forms.MessageBox]::Show("Es wurde keine Server eingetragen, versuchen Sie das CSV nochmals zu laden.","Keine Server",0,[System.Windows.Forms.MessageBoxIcon]::Error)
            $sync.startbutton.Text = "Start"
            $sync.status = "standby"
        }
    }
    #Stoppt das Abfragen der Server, falls der Benutzer den Stop-Knopf drückt 
    elseif($sync.status -eq "running"){
        $sync.startbutton.Text = "Start"
        $sync.status = "standby"
        $sync.durchgang++
    }
})
#####################################################################################
#################################### Form zeigen ####################################

$sync.form.ShowDialog() #Stoppt hier solange, bis das Fenster geschlossen wird
Save-ToCsv #Wenn das Form geschlossen wurde wird nochmals gespeichert

#####################################################################################