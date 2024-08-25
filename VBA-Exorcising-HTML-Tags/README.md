# VBA Grimoire - Exorcising HTML Tags

**HexBreaker Tech** - **Mastering Digital Exorcisms**

## Description

Welcome to the **VBA Grimoire**, where we conjure powerful scripts to banish the digital curses haunting your data! This repository provides an enchanted Excel template equipped with a pre-configured button that invokes a VBA spell to cleanse your cells of pesky HTML tags. This spell is ideal for purging data polluted by web extractions and restoring it to its purest form.

## What’s Included

- **Enchanted Excel Template**: A blank Excel workbook, mystically imbued with a button that unleashes the VBA spell to remove HTML tags.
- **VBA Spell (Code)**: A `.txt` scroll containing the incantation (VBA code) for those who wish to wield the spell independently.

## How to Use the Enchanted Template

1. **Download the Artefacts**:
   - Secure the Excel template and the spell scroll (`.txt` file) from this repository.

2. **Open the Template**:
   - Unseal the `HTML_Tag_Exorcism_Template.xlsx` file in Excel.

3. **Summon Your Data**:
   - Paste your data (inflicted by HTML tags) into the worksheet.

4. **Invoke the Spell**:
   - Press the "Remove HTML Tags" button to cleanse the data.

5. **Witness the Purification**:
   - The spell will purge all HTML tags from the entire used range of the worksheet, leaving only pure, hex-free data.

## How to Use the VBA Spell Independently

1. **Open Your Grimoire (Excel)**:
   - Launch Excel and press `ALT + F11` to access the VBA editor.

2. **Transcribe the Spell**:
   - Go to "File" > "Import File..." and select the `RemoveHTMLTags.txt` scroll from your repository.

3. **Cast the Spell**:
   - Run the macro from the VBA editor or bind it to a button or shortcut to invoke it whenever needed.

## Artefacts

- **`HTML_Tag_Exorcism_Template.xlsx`**: The enchanted Excel template with a pre-configured button for HTML tag removal.
- **`RemoveHTMLTags.txt`**: The ancient scroll containing the VBA spell for the macro.

## VBA Spell (from `RemoveHTMLTags.txt`)

```vba
Sub RemoveHTMLTags()

    ' Declare basic set of variables to perform the HTML Tag Cleanup in Scanned range
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim regEx As Object
    Dim strInput As String
    Dim strOutput As String

    ' Set the worksheet variable to the active sheet
    Set ws = ActiveSheet
    
    ' Establish the data range with the used range object
    Set rng = ws.UsedRange
    
    ' Create the regular expression object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "<.*?>"
    End With

    ' Loop through each cell in the declared data range variable
    For Each cell In rng
        strInput = cell.Value
        If regEx.test(strInput) Then
            strOutput = regEx.Replace(strInput, "")
            cell.Value = strOutput
        End If
    Next cell

    ' Clean up the established values
    Set regEx = Nothing
    Set rng = Nothing
    Set ws = Nothing

End Sub
