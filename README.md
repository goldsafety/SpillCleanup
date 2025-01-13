SpillCleanup (C)Copyright Stephen Goldsmith 2024-2025. All rights reserved.
Distributed at https://github.com/goldsafety/ and https://aircraftsystemsafety.com/code/

Excel VBA script to cleanup spilling from dynamic arrays by resolving #SPILL! errors and removing blank rows.

Eclipse Public License - v 2.0
THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
https://www.eclipse.org/legal/epl-2.0/

As described by Microsoft on their support page, "Excel formulas that return a set of values, also known
as an array, return these values to neighboring cells. This behavior is called spilling. Formulas that can
return arrays of variable size are called dynamic array formulas. Formulas that are currently returning
arrays that are successfully spilling can be referred to as spilled array formulas". For more information:
https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531

An example of a dynamic array formula in Excel is the FILTER function, which allows you to filter a range
of data based on criteria you define. The data returned by a dynamic array formula will change as the
source data range is updated, so the size of the spilled array can also change its shape. However, Excel
will only successfully spill if there is no other values (and no merged cells) which would overlap the
spilled range. As Excel makes no attempt to move other values as a spilled array changes its shape, you
will frequently find that either a dynamic array formula will return a #SPILL! error or large blank areas
appear where data was previously being spilled into. When using these formulas to create a dynamic report
that you might want to print or export, this leaves you with potentially a lot of manual inserting or
deleting of rows to accommodate the changed data.

This script has been written to automate the process of resolving #SPILL! errors and removing blank rows
below spilled ranges where data used to reside. It assumes that spilled ranges will only change the number
of rows being returned, and inserts or deletes entire rows below selected dynamic array formulas until
each #SPILL! error has been resolved and only a single blank row exists underneath it. A blank row is
defined as a row without any data, even if it contains formatting. The script further handles disabling
worksheet protection if set, ensuring formatting of the first row is reflected in inserted rows, and
wrapping text where it has been set.

Dynamic array formulas have some significant limitations, as they can be slow when you have many formulas
and they cannot spill into merged cells, which can make vertical layout of a report challenging. When you
need more control to produce a dynamic report, try ProtoSheet. This Excel VBA script requires you to
layout a prototype worksheet which it then uses to construct a completed version in a new worksheet on
demand. To find out more and to try this option, visit the following site:
https://github.com/goldsafety/ProtoSheet/

Known limitations of SpillCleanup are that it currently only resolves #SPILL! errors caused by data in
rows below the formula. If the #SPILL! error is caused by data or merged cells to the right of the dynamic
array formula, it will either raise an error or attempt to insert many rows until it figures something has
gone wrong (at which point a hundred or more rows will already have been inserted). In addition, only
#SPILL! errors caused by the FILTER or SORT dynamic array functions will be resolved, though it is
relatively simple to add support for others (please let me know).

#Installation

SpillCleanup is a sub procedure which should be called from either a command button inside your workbook
or from a user command added to the quick access toolbar or to the ribbon. To install, import the
sgSpillCleanup.bas file into the macro-enabled workbook in which you want to use the procedure, or to a
macro-enabled workbook which you can then save as an Excel Add-In.

For more information about getting started with VBA in Microsoft Office, please visit
https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office

#Acknowledgements

Microsoft and Excel are registered trademarks of Microsoft Corporation. All third party trademarks belong
to their respective owners, and the code and the author of this VBA script are in no way affiliated with
or endorsed by Microsoft.