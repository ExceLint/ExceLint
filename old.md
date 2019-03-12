[](#glossary)
## Glossary

|term|meaning|
|----|-------|
|ExceLint|ExceLint is a static analysis that finds formula errors in spreadsheets.|
|ExceLint UI|The ExceLint UI is an implementation of the ExceLint analysis, written as a plugin for Microsoft Excel on Windows.|
|Workbook|An Excel file is called a _workbook_.  Workbooks usually end in `.xls` or `.xlsx`|
|Worksheet|An Excel workbook usually contains many spreadsheets; each spreadsheet is called a `worksheet`. You can navigate worksheets in Excel by using the tabs on the bottom left of the workbook.|
|Ribbon|The ribbon is a user interface component that groups buttons together.  Buttons are grouped by function and organized by function, with that function's name appearing in the tabs at the top of the ribbon.  The ribbon is usually found at the top of a workbook, just under the Excel window's title bar.|
|Formula|A _formula_ is an Excel expression.  All Excel formulas are purely functional.  Every formula is prefixed by a `=` character.|
|Reference|A reference in Excel is a syntactic construct that indicates where another cell's value should be substitued into a formula during evaluation.  For example, the formula `=A1+A2` means that the values stored in cells `A1` and `A2` should be substituted into the expression where `A1` and `A2` occur, respectively, when the formula is evaluated.
|Reference shape|Two formulas are _reference equivalent_ if they refer to the same cell offsets, _relative to the position of the formula itself_.  Such formulas are said to have the same _reference shape_. Refer to the definition of _reference equivalence_ on page 4 of our paper for further elaboration.|
|Vector fingerprint|Each reference in a formula induces a reference vector, which is a vector encoding of the reference relative to the location of the formula itself.  Since a formula may have multiple refernces, it induces a set of vectors.  For performance reasons, ExceLint "compresses" this set of vectors into a single vector, called the _vector fingerprint_.  See section 4.1.1. on page 10 of the paper for further elaboration.|
|Formula error|A _formula error_ is a formula that deviates from the intended reference shape by either including an extra reference, omitting a reference, or misreferencing data. We also include manifestly wrong calculations in this category, such as choosing the wrong operation.|

We tested ExceLint using Windows 10/Windows Server 2016 and Excel 2016. While ExceLint works in principle with other versions of Windows and Excel (e.g., Excel 2010/2013), we have not tested these alternative configurations and do not recommend using them.

If you want to build the ExceLint add-in from source code, 
1. (optionally) ExceLint source code
1. (optionally) Visual Studio 2017: (we use the Professional edition; Community may also work) [https://visualstudio.microsoft.com/vs/whatsnew/](https://visualstudio.microsoft.com/vs/whatsnew/)

