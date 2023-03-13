# David Ordner-Pfadvariable
$dvArchivePath = "\\servername\david\archive\user\10004000\in";

# Zeitspanne im Format m-d-Y, wobei Monat und Tage ohne führende Nullen anzugeben sind
# Im Beispiel: 1.4.2016 - 28.2.2023
# ---------------------------------------
# WICHTIG: Die Angabe der Zeitspanne muss zwingend mit doppelten Anführungszeichen erfolgen,
# sonst funktioniert es nicht; daher hier im PS Beispiel einfache Anführungszeichen zur Definition
# des Strings und darin nochmal doppelte Anführungszeichen für die Angabe des DvItemFilterBits Strings
# ---------------------------------------
$dvTimeRange = '"4-1-2016 00:00:00 - 2-28-2023 0:00:00"';

# Ausgabe der Pfadvariable zur Visualisierung bei der Skriptausführung
Write-Host "Der verwendete Pfad lautet: $dvArchivePath.";
Read-Host "Fortfahren mit Enter";

# API initialisieren
$dvApi = new-object -comobject DVOBJAPILib.DvISEApi;

# Account Objekt erzeugen
$oAcc = $dvApi.Logon('','','','','','AUTH');

# Archive Objekt erzeugen (gewünschten Pfad eintragen)
$oArchive = $oAcc.ArchiveFromID($dvArchivePath);

# Alle Einträge des Ordners als MessageItem2 einlesen
$entries = $oArchive.GetArchiveEntries("Default");

# Schleife über alle Eintrag und Ausgabe des Betreffs mit DocumentType
Write-Host "--- Ausgabe des Betreffs und DokumentTyp aller MessageItem-Typen (DvItemFilterBits Default) ---";
Write-Host;

foreach ($item in $entries) { 
	$output = $item.Fields('Subject').Value + " (Type: " + $item.Fields('DocumentType').Value + ")";
	Write-Host $output;
}

Write-Host;

# Nur Faxe des Ordners als MessageItem2 einlesen ()
# DvItemFilterBits = DvFilterOnlyFax | String = OnlyFax
# Da die DvFilterBit-Typen in Powershell nicht funktionieren, wird die Stringvariante verwendet
$entries = $oArchive.GetArchiveEntries('OnlyFax');

# Schleife über alle eingelesenen Einträge und Ausgabe des Betreffs mit DocumentType
Write-Host "--- Ausgabe des Betreffs und DokumentTyp von Fax-Typen (DvItemFilterBits OnlyFax) ---";
Write-Host;

foreach ($item in $entries) { 
	$output = $item.Fields('Subject').Value + " (Type: " + $item.Fields('DocumentType').Value + ")";
	Write-Host $output;
}

Write-Host;

# Alle Einträge des Ordners innerhalb einer Zeitspanne als MessageItem2 einlesen
$entries = $oArchive.GetArchiveEntries('StatusTime=' + $dvTimeRange);

# Schleife über alle Eintrag und Ausgabe des Betreffs mit DocumentType
Write-Host "--- Ausgabe des Betreffs und DokumentTyp aller MessageItem-Typen innerhalb der Zeitspanne $dvTimeRange ---";
Write-Host;

foreach ($item in $entries) { 
	$output = $item.Fields('Subject').Value + " (Type: " + $item.Fields('DocumentType').Value + ")";
	Write-Host $output;
}

Write-Host;
