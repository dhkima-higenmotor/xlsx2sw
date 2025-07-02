# xlsx2sw

_Automatic Solidworks Parts Generation from Excel Parameter Table_

## Dependency & Patch

* Solidworks
* uv (python package manager)


## Test

*  `example.SLDPRT` and `example.xlsx` should be exist in same directory.
* Base name `example` should be same on `.SLDPRT` and `.xlsx`
* `xlsx2sw.py` read from 3rd raw in `example.xlsx`
* Check `ACTIVATION` column in `xlsx2sw.py`...

```bash
cd D:/github/xlsx2sw
xlsx2sw.bat
# Select xlsx file on popup
```

## How to use

* Kill every `SLDWORKS.exe`
* Make `A.xlsx` and `A.SLDPRT` (`A` is user defined base name)
* Make Global Variables in `A.SLDPRT`
* Make Parameter Tables in `A.xlsx` with Global Variables
* Command

```
cd D:/github/xlsx2sw
uv run xlsx2sw.py
```

* Wait to Finish

