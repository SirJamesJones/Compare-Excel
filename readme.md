This was absolutly horrendous...

it was a fast job just import pandas and use openpyxl to read and modify the excel files, some quick code and we are good to go...

thats what you thought idiot .xls is so old it is not supported by pandas anymore and since I need pandas to actually overlap the two excel files, I converted the .xls files into the newer .xlsx format

we use goddamn excel 2016 which doesnt really support .xlsx

this is so fucked up,I dont even wanna look at this code anymore.

The code kinda does what it has to do by generating a new .xlsx file and marking all the numbers, in the first given excel file, that are overlapping in the first column of both excel files.

it kinda fucks with the structure so I guess its better that a new file got generated?

also the second given file is also converted to a .xlsx file, whick creates a new file aswell, so there is a little bit of unneeded bloat. but what the hell man, at this pont I don't care.

TODO

- give a new .xls file as output, to not destroy the original file because of conversion.
- create an executable that works on windows.