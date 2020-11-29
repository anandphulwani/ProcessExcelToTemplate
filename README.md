# PyExcelToMustache

Python Script that converts a read records from first sheet of `Excel(xlsx) File`, process into `Mustache Template` and generate `Rendered Text` file from processing it.

## Usage

This is a script for personal projects. Created to run in any environment.

It uses the records in  ```Excel``` and sets the values in the ```Mustache Template``` and copies the output to the ```Rendered Text```.

## Driving environment

-Any OS running `Python 3.6` or higher

### Dependency tools

Dependency packages are recorded in `Pipfile`, the `pipenv` configuration file.

 - [`pandas`]
 - [`chevron (mustache)`]

## install

After installing `Python 3.6` or higher, execute the following statement.

> ```$ pip install -r requirements.txt```

## Excel Sheet Rules

1. Only first Sheet of the excel workbook will be considered.
2. Lines 1 in the sheet should always be observed.

- Line 1: Description of the column

## How to use

1. Prepare an Excel (xlsx) file in the same format as ```sample.xlsx```. (In the Repository)

2. Enter the appropriate argument values ​​and execute ```ProcessExcelToTemplate.py```. It does not matter the order in which you put the arguments.

- ```-x```, ```--excel```: Excel(xlsx) path and file name. **Must be entered during execution.**
- ```-m```, ```--mustachetemplate```: Path to Template/Mustache file. **Must be entered when running.**
- ```-o```, ```--output```: Specify the path of the `output` folder, if filename specified with path it will use it, if just folder then it will generate a "Output.txt" inside it.
- ```-p```, ```--escape```: Characters to escape inside the data imported from excel.

3. Check the output in the `output` path.

## Sample Usage Commands

Without escaping anything inside the data imported from excel.

```$ ProcessExcelToTemplate.py -x sample.xlsx -m template.mustache -o .\\Output.txt```

Escaping single quote inside the data imported from excel.

```$ ProcessExcelToTemplate.py -x sample.xlsx -m template.mustache -o .\\Output.txt -p \'```

Escaping double quote inside the data imported from excel.

```$ ProcessExcelToTemplate.py -x sample.xlsx -m template.mustache -o .\\Output.txt -p \"```

## Thing to do

Nothing as of now
                        
