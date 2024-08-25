```markdown
# HexBreaker Tech - The Tome of System Insights

## Description

Welcome to HexBreaker Tech's mystical library, where we delve into the arcane arts of VBA to extract, decipher, and reveal hidden system data! This repository contains a powerful Excel Workbook, enchanted with a VBA script, nestled within the ThisWorkbook Sheet Property. This spell automatically gathers key system insights, such as user details, operating system information, and screen resolution upon opening the workbook.

## What's Included

- Enchanted Workbook: A pre-configured Excel Workbook that invokes the VBA script to extract and display system information automatically when the workbook is opened.
- VBA Spell (Code): The VBA script embedded within the ThisWorkbook Sheet Property, written to gather user-specific data and populate it into an Excel sheet for easy analysis.

## How to Use the Enchanted Workbook

### Step 1: Download the Artefact

- Secure the Excel Workbook from this repository.

### Step 2: Open the Grimoire

- Unseal the Workbook in Excel. As you open it, the VBA spell will be automatically invoked, and key system details will manifest within the active sheet.

### Step 3: Witness the Revelation

- The script will extract and populate information such as:
  - Excel Username
  - Windows Username and Domain
  - Computer Name and Environment
  - User Email (via Outlook)
  - Operating System Version
  - Office Version and Language
  - Regional Settings
  - Screen Resolution

### Step 4: Inspect the Results

- The data will be clearly laid out in columns A and B of the active sheet, with labels and values aligned for easy reading.

## The Spell (VBA Code)

This mystical VBA code is embedded in the ThisWorkbook Sheet Property within the VBA editor. Upon opening the workbook, the script is executed automatically. Below is the code for reference:

```vba
Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim wshNetwork As Object
    Dim outlookApp As Object
    Dim userEmail As String
    Dim osVersion As String
    Dim officeVersion As String
    Dim officeLanguage As Long
    Dim regionalSettings As String
    Dim screenResolution As String
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Clear existing content on the worksheet
    ws.Cells.Clear

    ' Create an instance of WScript.Network
    Set wshNetwork = CreateObject("WScript.Network")

    ' Set labels in column A
    ws.Range("A1").Value = "Username (Excel)"
    ws.Range("A2").Value = "Windows Username"
    ws.Range("A3").Value = "User Domain"
    ws.Range("A4").Value = "Computer Name"
    ws.Range("A5").Value = "User Profile Path"
    ws.Range("A6").Value = "Home Drive"
    ws.Range("A7").Value = "Home Path"
    ws.Range("A8").Value = "Computer Name (Environment)"
    ws.Range("A9").Value = "User Email Address (Outlook)"
    ws.Range("A10").Value = "Operating System Version"
    ws.Range("A11").Value = "Office Version"
    ws.Range("A12").Value = "Office Language"
    ws.Range("A13").Value = "Regional Settings"
    ws.Range("A14").Value = "Screen Resolution"

    ' Set values in column B
    ws.Range("B1").Value = Application.UserName ' Excel Username
    ws.Range("B2").Value = Environ("USERNAME") ' Windows Username
    ws.Range("B3").Value = Environ("USERDOMAIN") ' User Domain
    ws.Range("B4").Value = Environ("COMPUTERNAME") ' Computer Name from Environ
    ws.Range("B5").Value = Environ("USERPROFILE") ' User Profile Path
    ws.Range("B6").Value = Environ("HOMEDRIVE") ' Home Drive
    ws.Range("B7").Value = Environ("HOMEPATH") ' Home Path
    ws.Range("B8").Value = wshNetwork.ComputerName ' Computer Name from WScript.Network

    ' Try to get the user's email address via Outlook
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If Not outlookApp Is Nothing Then
        userEmail = outlookApp.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    Else
        userEmail = "Email not available"
    End If
    On Error GoTo 0
    ws.Range("B9").Value = userEmail ' User Email Address

    ' Get the Operating System version
    osVersion = CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
    ws.Range("B10").Value = osVersion ' OS Version

    ' Get the Office version
    officeVersion = Application.Version
    ws.Range("B11").Value = officeVersion ' Office Version

    ' Get the Office language using the LCID (Locale Identifier)
    officeLanguage = CLng(Application.LanguageSettings.LanguageID(msoLanguageIDUI))
    ws.Range("B12").Value = GetLanguageName(officeLanguage) ' Office Language

    ' Get the Regional Settings
    regionalSettings = Environ("LANG")
    ws.Range("B13").Value = regionalSettings ' Regional Settings

    ' Get the Screen Resolution
    screenResolution = Application.UsableWidth & "x" & Application.UsableHeight
    ws.Range("B14").Value = screenResolution ' Screen Resolution

    ' Adjust column width for better readability
    ws.Columns("A:B").AutoFit

    ' Clean up
    Set wshNetwork = Nothing
    Set outlookApp = Nothing
End Sub

Function GetLanguageName(lcid As Long) As String
    Dim locale As Object
    On Error Resume Next
    Set locale = CreateObject("MSForms.Globalization")
    GetLanguageName = locale.Language(lcid)
    If Err.Number <> 0 Then
        GetLanguageName = "Unknown Language"
    End If
    On Error GoTo 0
End Function
```

## Notes

- Environment Compatibility: This VBA script is designed for systems running Microsoft Excel with access to Outlook and WScript.
- Limitations: The user email extraction relies on Outlook being available on the system. If Outlook is not available, the script will display "Email not available."

## Contribute

Want to improve this script or add your own enhancements? Fork this repository and submit a pull request. Let's continue mastering digital wizardry together!

---

HexBreaker Tech - Unleashing the Power of Digital Sorcery
```
