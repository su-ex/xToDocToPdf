# xToDocToPdf
This is a PowerShell script suite which provides through three callable scripts various functions that can be used together or on its own.

The main purpose is Word documents and PDF documents being concatenated to a single Word document which then can be converted to a PDF document.

Additionally xToDoc has a few neat features like replacing variables inside Word documents. Details are explained below.

![Übersicht](https://user-images.githubusercontent.com/9221456/90759422-ad85cc00-e2e0-11ea-821f-61f7890de7cf.png)

## How to "install"
You'll need PowerShell Core Version 7 or up installed. You can get it [here](https://github.com/powershell/powershell#get-powershell).

This project itself is only a collection of scripts, so just clone this repo to any directory or extract the zip archive if you don't have git installed.

**For PDF file handling** you'll need [qpdf](https://github.com/qpdf/qpdf). Download the 32bit version with the Windows executable called `qpdf-<version>-bin-mingw32.zip`, extract it and put the whole extracted folder inside the parent folder of this repo's folder.

To grab variables out of Excel and replace them inside Word documents, you'll need the ImportExcel PowerShell module inside the parent folder of this repo's folder. To do that run the following in a PowerShell Console (adjust the path):  
`Find-Module -Name 'ImportExcel' -Repository 'PSGallery' | Save-Module -Path 'this\repo's\parent\folder'`

After all this you should end up with a directory containing
- xToDocToPdf _(this repo)_
- qpdf\__\<version\>_
- ImportExcel

## How to call the scripts (work around PowerShell execution policy)
`pwsh.exe -ExecutionPolicy Bypass -Command "& \"path\\to\\script\" -string-parameter \"blabla\""`

Of course adjust path and name according to this repo's path on your filesystem and put the name in of the script you want. And you can add multiple parameters of course. Just be sure to escape `"` and `\` with a `\` and always put paths between (escaped) quotation marks.

## xToDoc

### SYNOPSIS

**xToDoc** \[**-working-directory** _path_] \[**-target-file** _path_] \[**-selected-description-file** _path_] **-template-description-file** _path_ \[**-lang** _language\_identifier_] [**-get-variables-from-excel** **-excel-variables-workbook-file** _path_ **-excel-variables-worksheet-name** _worksheet\_name_ **-excel-variables-table-name** _table\_name_] [**-get-translations-from-excel** **-excel-translations-workbook-file** _path_ **-excel-translations-worksheet-name** _worksheet\_name_ **-excel-translations-table-name** _table\_name_] \[**-custom-template-pdf-page** _path_] \[**-custom-base-path** **,**_path_ ...]

### DESCRIPTION

`xToDoc` does these main steps by
the following sequence:

 1. Read in a template-description file (this file's structure is explained below)
 2. Show a tree dialog in which elements of the read in description can be selected and store this selection back in a selected-description file
 3. Concatenate each word file and one word placeholder page for each page of a pdf file, the order is given in the template-description file
 4. Replace variables (how to use that is explained below)
 5. Update headings so they have right numbering and formatting
 6. Update fields and table of contents
 7. Apply formatting tags inside normal word text
     - `{{b}}some text that should be bold{{/b}}`
     - `{{u}}some text that should be underlined{{/u}}`
     - `{{i}}some text that should be italic{{/i}}`
 8. Save the generated target word document

#### Options

`-working-directory`  
:   Sets a directory to which paths of other options will be relative to if there hasn't been given an absolute path. Defaults to current directory.

`-target-file`  
:   Specifies the target Word document's path and filename.

`-selected-description-file`  
:   Specifies the path and filename of the description file where the selection in the tree dialog is stored.

`-template-description-file`  
:   Specifies the path and filename of the description file which is read in if there was no tree selection yet or if it should explicitly be used instead of the already existing selection.

`-lang`  
:   Specifiy a language identifier. Template documents will be taken from a subfolder with this language identifier as folder name if not left empty.

`-get-variables-from-excel`  
:   This switch enables grabbing variables from an Excel table.

`-excel-variables-workbook-file`  
:   Specifies the Excel workbook file's path to grab the variables from. This option can only be used together with `-get-variables-from-excel`.

`-excel-variables-worksheet-name`  
:   Specifies the Excel worksheet's name of the given workbook to grab the variables from. This option can only be used together with `-get-variables-from-excel`.

`-excel-variables-table-name`  
:   Specifies the Excel table's name on the given  worksheet of the given workbook to grab the variables from. This option can only be used together with `-get-variables-from-excel`.

`-get-translations-from-excel`  
:   This switch enables grabbing translations for pdf headings from an Excel table.

`-excel-translations-workbook-file`  
:   Specifies the Excel workbook file's path to grab the translations for pdf headings from. This option can only be used together with `-get-translations-from-excel`.

`-excel-translations-worksheet-name`  
:   Specifies the Excel worksheet's name of the given workbook to grab the translations for pdf headings from. This option can only be used together with `-get-translations-from-excel`.

`-excel-translations-table-name`  
:   Specifies the Excel table's name on the given  worksheet of the given workbook to grab the translations for pdf headings from. This option can only be used together with `-get-translations-from-excel`.

`-custom-template-pdf-page`  
:   Specifies the path to a template page that should be used for each page of a pdf document. See assets/PDF.docx for the default one.

`-custom-base-path`  
:   Here you can give a comma separated list of custom base paths you can make use of in a description file. They'll be numerated starting by 1.

### Syntax of a description file

The file is read linewise.

Each line starting with any number (also meaning 0) of whitespace characters (tab or space and such) followed by a `#` is treated as a comment line and thus ignored. You can't put comments in lines that should be processed. Lines containing only any number of whitespace characters are also ignored.

A normal line should look like this (every placeholder like this `⟨..⟩` needs to be replaced, phrases enclosed by these `[..]` are optional):  
`[;]⟨indent⟩⟨description⟩: [⟨flags⟩>]⟨path⟩`

- A line starting with a `;` means it's disabled.
- The indent **must be** through tabs. With this you can build a tree.
- The description is shown in the tree selection dialogue.
- Flags can alter the handling of the line and are described in detail in the table below.
- The path needs to be to a word or pdf document or to a folder.

All paths in a description file are relative to the **template** description file (this means relative paths in the selected description file are still relative to the template description file) if custom base path flag `c` is not set for a line (see below)!

#### Flags

flag&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | meaning
:--- | :---
`a[r][k]` | Extract each word and pdf document alphabetically sorted out of the given path which needs to be a folder when this flag is set. A folder alphabetically included will start one indentation higher than the description line instantiating it. If you add an `r` this means to recurse (also go through subfolders, indentation is increased each step). If you add a `k` this means keep indentation (documents in the topmost directory have the indentation of the description line instantiating it).
`s` | This flag means skip the language folder. This has the same effect as if `-lang` was not set, but just for the current description line. If `-lang` is not set at all, this flag has no effect of course.
`c[⟨n⟩][r]` | This flag lets you choose a custom base path. Normally each template file is searched for relative to the template description file. With this flag it's searched for relative to the custom base path. If you don't add a number n, it will be 1, but you can add a number starting by 1 which should match the number of the custom base path you want and fed into the script via the `-custom-base-path` option. If you add an `r` the custom base path will apply recursively to all description lines immediately below with a higher indentation too.
`h[⟨n⟩\|r]` | This option applies only to pdf documents (also useful when searching a folder for documents alphabetically with the `a` flag and there are pdf documents in it). It is useful for the target word's table of contents and jump marks. This option will add a word heading for a pdf document with the content of the description. You can add a number n from 1 to 9 which will match word's headings from 1 to 9. If you don't add a number, it will match the current indentation. Instead of a number you can also append an `r` which will again apply this option recursively to lower indentations too (the heading tier always matching the indentation).
`p(n\|t)` | This option let's you force **t**his concatenated word document to start with a new page if you append a `t`. If you append a `n` there will be a page break in between each of the word documents on the **n**ext higher indentation level. You have to append either `n` or `t`.

All flags can be freely mixed.

In a description line they must be ended with a `>` right before the path starts.

### Special syntax for alphabetically searched folders

The problem with alphabetically searching a folder for documents is that there's no easy way adding information like a description and heading tier level for pdf documents. Thus there is a special file-/foldername syntax:
- everything between two twice underscores will become the description: \_\_⟨description⟩\_\_
- everything between two twice hashtags becomes the pdf heading tier: \#\#⟨pdf heading tier from 1 to 9 or r⟩\#\#

If there's no description through the underscore syntax, the whole filename will be the description.

### Variable replacement in Word files

To use this feature, you'll need an Excel table with two columns and a header line (what you use as headers doesn't matter). Very important is that you assign a name to the table. The left column should contain the variable names, the right column should contain the replacements. This table can be fed into the script via the `-excel-` options.

The syntax for variables inside Word is `{{$⟨variable_name⟩}}`.

So if you have a variable like `{{$asupervariable}}` inside a word document and in the Excel table you fed into the script a row with `asupervariable` in the left column and `a super text` in the right column, `{{$asupervariable}}` will be replaced with `a super text` in the target word document.

Variables may contain formatting tags like explained above for bold, underline and italic and any special characters word understands in its replace dialogue starting with those carets `^` like `^p` for a paragraph break.

## docToPdf

### SYNOPSIS

**docToPdf** \[**-working-directory** _path_] \[**-source-word-file** _path_] \[**-target-pdf-file** _path_] 

### DESCRIPTION

`docToPdf` does these main steps by
the following sequence:

 1. Identify the pdf placeholder pages in the source word file and extract their information (the paths must be absolute, see assets/PDF.docx for an example template page)
 2. Determine the orientation of the word document
 3. Hide the placeholders and export the Word document as PDF document
 4. Extract the orientation of each page in each PDF document and rotate pages in the exported PDF document whose orientation doesn't match
 5. Overlay the pages in the exported PDF document with those pages of the PDF documents they should be overlaid with

#### Options

`-working-directory`  
:   Sets a directory to which paths of other options will be relative to if there hasn't been given an absolute path. Defaults to current directory.

`-source-word-file`  
:   Specifies the source Word document's path. This should probably have the same value like the xToDoc `-target-file` option.

`-target-pdf-file`  
:   This is optional. It specifies the path and filename of the target PDF document. If omitted, the source Word document's file extension will be swapped to the pdf file extension.

## BUGS

See GitHub Issues: <https://github.com/su-ex/xToDocToPdf/issues>

## LICENSE

GPLv3, see [LICENSE](LICENSE)

## AUTHOR

su-ex \<codeworks@supercable.onl\>
