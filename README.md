## install requirement project's packages

```commandline
pip install -r requirements.txt
```

## UPS_zipcode

This is the solution to a test task.

Description of customer requirements in the *.rtf file in the folder [Inbox date]

The program downloads zip code ranges from the UPS website and saves them in the file.
    

То start run `main.py`

Input data in the `[Inbox Data]` folder

Downloaded data in the `[Output Data]` folder

The file with the input ranges can be opened by the user, so the file with the corrected ranges is saved in the `[Output Date]` folder.

The file with the corrected ranges is called `NEW Carriers zone ranges.xlsx`


Range uniqueness was not checked - there was no such condition in the input task.

Downloads` *. xls` files have the wrong extension, so instead of converting, rename them to `*.xlsx`


To limit the number of downloaded files, set `COUNT_FILES = 20`, 
`COUNT_FILES = None` - downloads all files

To simplify debugging/opening in IDE, an additional text file with ranges is created in the `[Output Date]` folder
