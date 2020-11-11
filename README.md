# CLEWs_regionalization

**Contents:**
* [Description](#description)
* [Files](#files)
* [Input File](#input-file)
* [User Inputs](#user-inputs)
  * [Input Excel File Name](#input-excel-file-name)
  * [Area Name](#area-name)
  * [Regions](#regions)
  * [Output File Names](#output-file-names)
    * [CSV File Names](#csv-file-names)
    * [Excel Workbook File Name](#excel-workbook-file-name)
* [Output Files](#output-files)


## Description
The script **regionalization.py** creates a regionalized map.
It takes in an excel workbook with information about the whole area and each individual region to produce this map. 

It saves the map, a legend, and a list of any overlaps 
between regions. A formatted map can also be saved if specified
by the user.

## Files
* **regionalization.py**
	* main script
* **individual_region_files.xlsx**
	* example of the input excel workbook for Canada

## Input File
Input file: an excel workbook with the following worksheets:
* **'[area_name]'** --> one sheet
	* One sheet with entire area
	* Copy pasted .asc file of entire area
* **'[region_name]'** --> multiple sheets
	* Multiple sheets, each with one individual region
	* Copy pasted .asc file of regions
* **'list'** --> one sheet
	* OPTIONAL: alternatively, the command line can be used for this input
	* One sheet with the list of regions and their number (numbers in column 1, names in column 2)
		* Note: Region names in column 2 should match the names of their 
	corresponding excel worksheet


## User Inputs
The user will have to enter the following into the command line:
  * [Input Excel File Name](#input-excel-file-name)
  * [Area Name](#area-name)
  * [Regions](#regions)
  * [Output File Names](#output-file-names)
    * [CSV File Names](#csv-file-names)
    * [Excel Workbook File Name](#excel-workbook-file-name)

### Input Excel File Name
The default input file name is "individual_region_files.xlsx", and
the user will be asked to confirm or change this name.

### Area Name
Input for this area name should be the same as the name of the excel 
worksheet for the entire area. If no sheet by this name can be found,
user input will be asked for again.

### Regions
As mentioned under section [Input File](#input-file), there is an optional 
worksheet 'list'. 

If this worksheet exists, then the region numbers and regions will be 
automatically taken into regions. If not, then all regions can be added
using the command line. Additionally, all regions can be 
**edited or deleted** before being confirmed.

### Output File Names
These are the names of the files to which the information will be saved.

#### CSV File Names
As mentioned under section [Description](#description), each time the 
program runs, the regionalized map, a legend, and a list of overlaps 
will be saved as .csv files. The default names are:
* **'map.csv'** for the regionalized map
* **'legend.csv'** for the list of region numbers and names
* **'overlaps.csv'** for the list of cells with overlaps

#### Excel Workbook File Name
When indicated, the user can choose whether to save a formatted regionalized 
map as an excel workbook as well. If this option is chosen, the file name
must also be specified.

The default file name is "formatted_map.xlsx", and the user will be 
asked to confirm or change this name. Note that this name must be different than the input excel file name, mentioned under section [Input Excel File Name](#input-excel-file-name).

## Output files

* .csv file for the map:
	* The same nodata value is used for this file as the area map from the 
	original input file
* .csv file for the legend:
	* Contains the region number and names that were used in the 
	regionalization
* .csv file for overlaps:
	* Contains cell numbers for any cells that had an overlap of regions
* .xlsx file for the formatted map (if applicable):
	* The same content as .csv file for the map
	* Formatted with a color scale (red, yellow, green)
	* Column width is set to be smaller



