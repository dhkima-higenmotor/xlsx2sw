# xlsx2sw

_Automatic Solidworks Parts Generation from Excel Parameter Table_

## Dependency & Patch

* in Miniconda3

* Packages
```
conda install pandas
conda install openpyxl
pip install pySW
```

* Patch
```
copy /Y .\patch\commSW.py %LOCALAPPDATA%\miniconda3\Lib\site-packages\pySW\commSW.py
```

## Test

*  `example.SLDPRT` and `example.xlsx` should be exist in same directory.
* Base name `example` should be same on `.SLDPRT` and `.xlsx`
* `xlsx2sw.py` read from 3rd raw in `example.xlsx`

```bash
cd D:/github/xlsx2sw
python xlsx2sw.py D:/github/xlsx2sw/example/example.xlsx
```

## How to use

* Kill every `SLDWORKS.exe`
* Make `A.xlsx` and `A.SLDPRT` (`A` is user defined base name)
* Make Global Variables in `A.SLDPRT`
* Make Parameter Tables in `A.xlsx` with Global Variables
* Command

```
cd D:/github/xlsx2sw
python xlsx2sw.py <path>/A.xlsx
```

* Wait to Finish
