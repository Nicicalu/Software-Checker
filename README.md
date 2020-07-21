# Software-Checker V2.1
*Die deutsche Version ist weiter unten!*

## English
*Note: All comments in the file are german.*

This readme describes how the script works and what the requirements are.
[Github Repository](https://github.com/Nicicalu/Datenbank-Script-M122)

### Requirements

To be able to read the .Net version, a CSV file named "dotnetversions.csv" must be placed in the same folder as the script. In this document the first column must contain the releases with the column header "Release" and the second column must contain the version numbers with the column header "Version". The data for the file can be found at [Link](https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed#find-newer-net-framework-versions-45-and-later). This list is from Microsoft.
***

The data of the specified computers are automatically saved in the current CSV file. By default it uses the file "server.csv" in the same folder as the script. If the file does not exist, the table is empty and you can import a CSV file yourself (button "Import CSV").

All changes are then saved in this file. At the next start of the script it checks again if the file "server.csv" exists. If it exists, it will be imported. Otherwise the user can import a file by himself again. Only one column with the heading "Name" may exist in the CSV. The servers must then be listed below it. This is how the file must be structured:  
"Name"  
"Servername01"  
"Servername02"

**Attention:** If the CSV is incorrect, i.e. wrong headers and/or wrong syntax, the CSV is reset to the default. That means it contains only the column header and no values.

### How the the script works

Once you have listed all the servers you need, you can still choose which software to check. If you set a check mark the software will be checked. Otherwise not.

How the check works:

**Java and OpenJDK**  
For Java, a search is made on the target server for a "Java.exe" which is somewhere in a folder where "jre" occurs. This would be the case with this path, for example:  
C:\Program Files\Java\jre1.8.0_201\bin\java.exe  
For OpenJDK it is also the "Java.exe" but the one that is somewhere in a folder containing the string "JDK".

_*Note:*_ It searches on all drives that are installed in the PC (no USB sticks or network drives). To improve performance, the Windows and Users folders are not included in the search.

**IIS and PHP:**  
For IIS the property "ProductVersion" is read from the file "w3wp.exe". This is located in the folder "%windir%\System32\inetsrv\". For PHP it is simply the PHP.exe which is searched in all drives.

**.Net:**  
For the .Net Framework the registry key "Full" is read in the following path:  
"HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP"

This value is the release number, which is then converted to the version in the file "dotnetversions.csv". Info about "dotnetversions.csv" see "Requirements".

**Tomcat:**
The Java version that was read from Java is then executed with the parameter "-cp path/the/catalina.jar org.apache.catalina.util.ServerInfo".Since Tomcat requires Java for this process, the Java check mark is automatically set if you set the check mark for Tomcat.

### Further links
 - Github repository: [https://github.com/Nicicalu/Software-Checker](https://github.com/Nicicalu/Software-Checker)
 - Dotnet Versions: [https://docs.microsoft.com](https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed#find-newer-net-framework-versions-45-and-later)
 
## Deutsch
 In diesem Readme wird beschrieben, wie das Script funktioniert und was die Voraussetzungen sind.
[Github Repository](https://github.com/Nicicalu/Datenbank-Script-M122)

### Voraussetzungen

Um die .Net Version auslesen zu können, muss im gleichen Ordner wie das Script platziert ist eine CSV-Datei mit dem Namen "dotnetversions.csv" sein. In diesem Dokument müssen in der ersten Spalte die Releases mit der Spaltenüberschrift "Release" und in der zweiten Spalte die Versionsnummern mit der Spaltenüberschrift "Version" sein. Die Daten für die Datei finden Sie diesem [Link](https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed#find-newer-net-framework-versions-45-and-later). Diese Liste ist von Microsoft.
***

Die Daten, der angegebenen Computer werden automatisch in der aktuellen CSV-Datei gespeichert. Standardmässig benutzt es die Datei "server.csv" im gleichen Ordner wie das Script. Wenn die Datei nicht vorhanden ist, dann ist die Tabelle leer und man kann eine CSV-Datei selbst importieren (Knopf "Import CSV").

Alle Änderungen werden dann in dieser Datei gespeichert. Beim nächsten Start des Scripts prüft es wieder ob die Datei "server.csv" existiert. Wenn sie existiert, wird sie importiert. Ansonsten kann der User wieder eine Datei selber importieren. Im CSV darf nur eine Spalte vorhanden sein mit der Überschrift "Name". Darunter müssen dann die Server aufgelistet werden. So muss die Datei aufgebaut sein:  
"Name"  
"Server01"  
"Server02"

**Achtung:** Wenn das CSV Fehlerhaft ist, das heisst falsche Überschriften und oder falsche Syntax, dann wird das CSV auf den Standard zurückgesetzt. Das heisst, sie beinhaltet nur noch die Spaltenüberschrift und keine Werte.

### Funktionsweisen der Prüfung

Wenn man alle benötigten Server aufgelistet hat, kann man noch auswählen, welche Software geprüft werden soll. Wenn man einen Haken setzt wird die Software geprüft. Ansonsten nicht.

Funktionsweisen der Prüfung:

**Java und OpenJDK:**  
Bei Java wird auf dem Zielserver nach einer "Java.exe" gesucht die irgendwo in einem Ordner ist, indem "jre" vorkommt. Das wäre z.B. bei diesem Pfad der Fall:  
C:\Program Files\Java\jre1.8.0_201\bin\java.exe  
Bei OpenJDK ist es auch die "Java.exe" aber die, die irgendwo in einem Ordner der "JDK" enthält ist.

_*Hinweis:*_ Es wird auf allen Laufwerken gesucht, die im PC eingebaut sind (keine USB-Sticks oder Netzlaufwerke). Um die Performance zu erhöhen werden die Ordner Windows und Users beim Suchen nicht berücksichtigt.

**IIS und PHP:**  
Beim IIS wird die Eigenschaft "ProductVersion" von der Datei "w3wp.exe" ausgelesen. Diese befindet sich im Ordner "<Windows-Directory> \System32\inetsrv\". Bei PHP ist es einfach die PHP.exe die in allen Laufwerken gesucht wird.

**.Net:**  
Beim .Net Framework wird der Registry-Key "Full" in folgendem Pfad ausgelesen:  
"HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP"

Dieser Wert ist die Release-Nummer, die dann in der Datei "dotnetversions.csv" zur Version umgewandelt wird. Info zu "dotnetversions.csv" siehe Vorraussetzungen.

**Tomcat:**
Die letzte Java-Version die bei Java ausgelesen wurde wird dann ausgeführt mit dem Parameter "-cp pfad/der/catalina.jar} org.apache.catalina.util.ServerInfo". Bei Server-Number steht dann die Nummer. Da Tomcat für diesen Vorgang Java benötigt wird der Haken bei Java automatisch gesetzt, wenn man den Haken bei Tomcat setzt.

### Weiterführende Links
 - Github Repository: [https://github.com/Nicicalu/Software-Checker](https://github.com/Nicicalu/Software-Checker)
 - Dotnet Versionen: [https://docs.microsoft.com](https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed#find-newer-net-framework-versions-45-and-later)
***
*von Nicolas Caluori © 2020*
