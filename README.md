# Finland IE344 - Excel merger

Finland IE344 is a small program used to merge data from two chosen excel files.
First file (ELL) has all the data needed for manifestation and second file has
Movement Reference Number (MRN) from Finnish customs tied to containers.
The data is merged and then new files are created:

- One excel file with all information merged together.
- Each Main Line Operator (MLO) will have their own file.

If script is run directly then the two files in folder 'example_files' will be used
and the result will be saved on the desktop.

## Packages in use
'Tkinter'
'Pandas'
'Xlwings'
'Numpy'

## License
MIT