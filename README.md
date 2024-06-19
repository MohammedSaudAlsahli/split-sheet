## How to install this tool


```bash
pip install split_sheet
```

## How to use this tool

#### simply write the name of the tool then the input file, output file and the trigger to split the columns with

```bash 
split-sheet "input.csv" "output.xlsx" -c "column name" # this is to split by column
or
split-sheet "input.csv" "output.xlsx" -n max_number_of_cells # this is to split in ordered number of cells
```
