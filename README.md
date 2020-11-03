# Convert .xlsx spreadsheet to .py Python script

If you've reached the limits of what can be done in an Excel
spreadsheet, use this tool to convert the .xlsx spreadsheet
file to a Python .py script.

```shell
./xlsx2py example.xlsx > example.py
```

To run the pytest unit test:

```shell
python3 -m pytest -sv
```

Another example:

```shell
$ ./xlsx2py --print 'Sheet_Sheet2.A2()' example.xlsx | python3
13.5
$
```
