# PSWikiTable

PSWikiTable is a PowerShell module for converting an Excel table to a MediaWiki [wikitext table](https://meta.wikimedia.org/wiki/Help:Table). PSWikiTable supports basic formatting, templates and links (internal and external).

More complex tables and tables that fetch data from different sources can sometimes be easier to maintain as an Excel table than as a wiki table. This module lets you convert from the former to the latter.

## Formatting

The following formatting elements are supported:

* Font size, bold, italic, strike through and color.
* Background color.
* Text adjustment (left, right, center).

## Links

Internal and external links are supported. For internal links to work you have to supply the base wiki link (i.g. ``https://mywiki.com/wiki`` or ``https://mywiki.com/wiki/index.php``) using the ``-WikiBaseUri`` parameter.

## Templates

MediaWiki templates are supported through the ``-Templates`` parameter.

``-Templates @{Yes = 'Yes'}`` will replace all instances where a cell contains only the value "Yes" (case does not matter) with the MediaWiki template "{{Yes}}".

## Example

Convert the first worksheet in the workbook to a wiki table.

```powershell
ConvertTo-WikiTable -Path .\MyTable.xlsx
```

Select a specific worksheet by name

```powershell
ConvertTo-WikiTable -Path .\MyTable.xlsx -Worksheet 'Server list'
```

Correctly convert internal wiki links

```powershell
ConvertTo-WikiTable -Path .\MyTable.xlsx -WikiBaseUri 'https://mywiki.com/wiki/index.php'
```

Use templates

```powershell
ConvertTo-WikiTable -Path .\MyTable.xlsx -Templates @{yes = 'Yes'; no = 'No'; warn = 'Warning'}
```

## Build

The project is multitarget and builds .NET Framework 4.72 and NET 5.0 assemblies. The former is for use with PowerShell 5 and the later for PowerShell 7.

```powershell
dotnet build -c=Release
```

## Install

Create a new folder named "PSWikiTable" where you keep your modules (i.g. %USERPROFILE%\Documents\WindowsPowerShell\Modules). Copy PSWikiTable.dll and EPPlus.dll to the folder from the correct target build depending on PowerShell version.

Import the module manually if it's not in one of the module auto-load locations ``Import-Module 'C:\MyModules\PSWikiTable'``.

## Todo

Things that may or may not happen in the future:

* Add comment based help.
* Support more formatting.
* Make default text and background color configurable.
* Make table header configurable.
* Add a module manifest.
* Clean up the code.
