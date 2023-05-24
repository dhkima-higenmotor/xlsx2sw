# xlsx2sw

_Automatic Solidworks Parts Generation from Excel Parameter Table_

## Dependency & Patch

* Solidworks 2023 SP2.1
* in Miniconda3

* Packages
```
conda install pandas
conda install openpyxl
conda install psutil
pip install pySW
```

* Patch
```
where python
copy /Y .\patch\commSW.py %userprofile%\scoop\apps\miniconda3\current\Lib\site-packages\pySW\commSW.py
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
