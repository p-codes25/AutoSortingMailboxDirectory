# Auto Sorting Mailbox Directory

This is an Excel template and macro that sorts and formats printable mailbox directory listing sheets for an apartment complex, condominium building, etc.

This project came about because our old mailbox directory printout was done in Excel with residents' names entered and updated by hand -- every time a person or family would move in or out, someone had to manually move the entries around in order to make things line up and eliminate the gaps.

I put together this Excel document and macro to simplify all that.

This spreadsheet lets you enter and update residents' names by unit number (which is how changes occur, a unit at a time) and produces a directory listing sorted alphabetically by last name with one resident's name per line (which is how the postal carriers need to look them up when a piece of mail doesn't have a unit number).

Below are the steps to set it up and use it.

## Configure the Number of Units in the Directory

Go to the first sheet in the Excel spreadsheet, labelled "Sorted By Unit Number."  This sheet is where you'll enter the names for each unit.  First, add or remove rows and columns in order to obtain the total number of units that will appear in your directory.

The number of rows and columns, and the grouping or ordering of unit numbers within this sheet, do not particularly matter, unless you will also want to print this sheet for a directory listing by unit number.

There are 100 units in the example sheet, but you can add more or remove entries as needed.  Be sure to follow these rules:

* Leave no more than one blank row between groups (two consecutive blank rows stops the macro from searching further for units/names).

* The column headings must start with the word "Unit" (in cells A1, E1, etc. for as many columns as you need).

## Configure the Printed Page Layout

At the top of the macro code are various settings that control the appearance and layout of the printed directory.

On the Excel ribbon bar, go to Developer -> Macros and click Edit to edit the UpdateAlphabeticList macro.  At the top of the code is a section with the settings listed below.

Changing the settings below is optional; you can run the macro as a test and see how the output looks, and only make changes to the settings necessary in order to have the printed directory look the way you want.

The printed directory shows one first/last name per entry, so you will typically have more entries in the output than there are unit numbers, to the extent that there are multiple names for each unit.

### Configurable Settings

Set the number of units you want to have per printed page in the output directory sheets.  You can put all the entries on one page if they will fit; or, you might want two or more sheets if you have different mailbox areas, for example.

    ' Number of units per printed page
    Const nUnitsPerPage As Integer = 50

Define the settings below to control how many rows and columns will be on each page of the printed directory:

    ' Number of names per row and column on each printed page
    Const nRowsPerPage As Integer = 25
    Const nColumnsPerPage As Integer = 3

Adjust these font settings so the type size looks good based on your number of rows and columns per page:

    ' Font settings
    Const nFontHeight As Integer = 14
    Const strFontFace As String = "Calibri"
    Const nHeadingFontHeight As Integer = 16

Set the column widths based on your choice of font and font size:

    ' Column widths
    Const nNameColumnWidth As Integer = 23
    Const nUnitColumnWidth As Integer = 4
    Const nSeparatorColumnWidth As Integer = 1

Finally, adjust the colors if you want:

    ' Colors
    Dim rgbSeparator As Long, rgbHeading As Long
    rgbSeparator = RGB(204, 255, 204)
    rgbHeading = RGB(102, 179, 102)

## Enter the Names

On the Sorted By Unit Number sheet, enter the names for each unit.  Follow these rules:

* If there are different family names (last names) within one unit, enter those names in one cell, separated by a forward slash character.

* If there are different given names (first names) for one family name, enter those first names separated by a comma.  To enter first names for different family names, separate them using a forward slash character.

* A last name may contain commas (for example: Jones, Jr.) but a first name may not.

For example:

	Marta Solis, Rene Solis, Lucy Mendez

would be entered as:

	Last Name          First Name
	---------          ----------
	Solis/Mendez       Marta,Rene/Lucy

## Process the Directory Listing

Once you have entered the names and settings (or just want to do a test), click the blue "Click Here to Update" button near the top-right corner of the first sheet.  This runs the macro, which breaks each unit's entry into individual names, sorts the names within groups and outputs printable sheets.

The macro will show you the output sheet, titled "Alpha For Printing" (there is also a second, intermediate sheet called "Alpha By Last Name", which you can ignore).

Tweak any settings or names as needed, and re-run the macro to update the third sheet.

Once you're done, you can export the Alpha For Printing sheet to a PDF file or print it directly from Excel!


