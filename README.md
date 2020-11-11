# CLEWs_regionalization

This allows regionalization **regionalization.py**

**In this readme file**:
- [Description](#description)
- [Files](#files)
- [Input file](#input-file)
- [User inputs](#user-inputs)
  * [Input excel file name](#input-excel-file-name)
  * [Area name](#area-name)
  * [Regions](#regions)
  * [Output file names](#output-file-names)
    + [CSV file names](#csv-file-names)
    + [Excel workbook file name (if applicable)](#excel-workbook-file-name--if-applicable-)
- [Output files](#output-files)


## Description
The script **regionalization.py** takes in an excel workbook
with information about the whole area and each individual 
region. It then produces a regionalized map. 

It saves the map, a legend, and a list of any overlaps 
between regions. A formatted map can also be saved if specified
by the user.

## Files


## Input file
Input file: an excel workbook with the following files:
* One sheet with entire area
	* '[area_name]': Copy pasted .asc file of entire area
* Multiple sheets with each individual region
	* '[region_name]': Copy pasted .asc file of region
* One sheet with the list of regions and their number
	* 'list': Region numbers in column 1, region names in column 2 
	* Region names in column 2 should match the names of their 
	corresponding excel worksheet


## User inputs
The user will have to enter the following into the command line:
  * [Input excel file name](#input-excel-file-name)
  * [Area name](#area-name)
  * [Regions](#regions)
  * [Output file names](#output-file-names)
    + [CSV file names](#csv-file-names)
    + [Excel workbook file name (if applicable)](#excel-workbook-file-name--if-applicable-)

### Input excel file name
The default input file name is "individual_region_files.xlsx", and
the user will be asked to confirm 

### Area name
### Regions
### Output file names
#### CSV file names
#### Excel workbook file name (if applicable)
## Output files

