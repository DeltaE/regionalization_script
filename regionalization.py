# mapOfRegions.py

#####################################################################
"""
This file takes in an excel workbook and produces a regionalized map
    and saves it under another workbook.

Input:
Before beginning, input excel file must have the following:
    Sheet1 = "[area]"              --> ex. CAN for canada, copy paste the .asc file
    Others = "[abbreviation/name]" --> copy pasted .asc file 
    **optional:
    list = "list" --> list of all region numbers and names

Output:
    Sheet1 = "map"          --> regionalized map
    Sheet2 = "legend"       --> region numbers and names
    Sheet3 = "overlaps"     --> cells with overlap between regions


openpyxl is used to control excel

"""
##############################################################################
import sys
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.formatting.rule import ColorScaleRule
import re
from string import digits
import csv

##############################################################################
# functions for interaction with user
##############################################################################

def y_or_n(question):
    while True:
        ans = input(question + " (y/n): ")
        if ans in ('y', 'Y'):
            return True
        elif ans in ('n', 'N'):
            return False
        else:
            print("Your answer '" + ans + "' was invalid.", 
                  "Please choose 'y' or 'n'.")


def answer(question):
    while True:
        ans = input(question + ": ")
        check = input('Please confirm your answer: "' + ans + '" (y/n): ')
        if check in ('y', 'Y'):
            return ans
        else:
            print("Answer was not confirmed. Please enter answer again.") 

def print_dict(dc):
    for key in dc:
        print(key, ':', dc[key])

def print_regions(regions):
    if regions == {}:
        return
    print("The following are your current regions:")
    print("-----------------------\n",
            "    regions:    ", 
            "\n-----------------------")
    print_dict(regions)
    return

def print_menu(regions):
    print("-----------------------\n",
            "    menu:    ", 
            "\n-----------------------")
    print("option 1: add regions")
    print("option 2: remove regions")
    print("option 3: finished editing regions")
    return


def choose_option(regions):
    print_regions(regions)
    print_menu(regions)
    
    while True:
        option = input("Please enter menu option of your choice (integer): ")
        if option in ('1', '2', '3'):
            option = int(option)
            return option
        else:
            print("Your input '" + option + "' was invalid.", 
                    "Please choose '1', '2', or '3'.")
    return

def add_regions(regions):
    print_regions(regions)
    while True:
        try:
            num = int(input("Please enter the region number (integer): "))
        except ValueError:
            print("Your input was not an integer. Try again.")
            continue
        name = input("Please enter region name (must be same as excel sheet): ")
        if num in regions:
            overwrite = y_or_n("There is already a region with that number." +  
                                " Would you like to overwrite it?")
            if not overwrite:
                continue

        regions[num] = name
        print_regions(regions)
        finished = y_or_n("Would you like to add another region?")
        if not finished:
            break
    return


def remove_regions(regions):
    print_regions(regions)
    while True:
        try:
            num = int(input("Please enter the region number (integer) to delete: "))
        except ValueError:
            print("Your input was not an integer. Try again.")
            continue
        if num in regions:
            del regions[num]
            print_regions(regions)
            finished = y_or_n("Would you like to remove another region?")
            if not finished:
                return
        else:
            print("Your input '" + str(num) + "' is not a region.",
                    "Please choose again.")
    return
 

def sort_regions(regions, ws = None):
    if ws is not None:
        for row in ws.iter_rows():
            regions[row[0].value] = row[1].value

    if regions == {}:
        print("There is no sheet 'list' in workbook.")
        print("Please either add region and region numbers manually,",
                "or create a worksheet called 'list'.")
    
    print("Use this to add or remove regions.")
    while True:
        option = choose_option(regions)
        if option == 1:
           add_regions(regions)
           continue
        elif option == 2:
           remove_regions(regions)
           continue
        # check if finished, then return
        elif option == 3:
            while True:
                finished = y_or_n("Please confirm. Are all regions correct?")
                if finished:
                    return
                break

def change_csv_names(save_csv_names):
    while True:
        change = input("Which file name would you like to change?" +
                       " ('map', 'legend', or 'overlaps'): ")
        
        q_pt_1 = "What would you like to change the saved file name for '"
        q_pt_2 = "' to?"

        if change in ('map', "'map'", '"map"'):
            save_csv_names['map'] = answer(q_pt1 + change + q_pt2)
        elif change in ('legend', "'legend'", '"legend"'):
            save_csv_names['legend'] = answer(q_pt1 + change + q_pt2)
        elif change in ('overlaps', "'overlaps'", '"overlaps"'):
            save_csv_names['overlaps'] = answer(q_pt1 + change + q_pt2)
        else:
            print("Your answer was not understood.")

        print_dict(save_csv_names)
        another = y_or_n("Would you like to change another filename?")
        if not another:
            return

def print_explain_csv(save_csv_names):
    print("These are the csv file names that the corresponding information will be saved to:") 
    print_dict(save_csv_names)
    print("***Please note that this program will overwrite existing files.***")


##############################################################################
# defining variables
##############################################################################


def define_variable(variable_name, ws=None):
    line_end = "===================="
    line_begin = "\n" + line_end

    if variable_name == "xlFilename":
        print(line_begin, "input excel file", line_end)
        xlFilename = "individual_region_files.xlsx"
        print('The current input excel file name is "' + xlFilename +'".')
        change = y_or_n("Would you like to change the input file name")
        if change:
            xlFilename = answer("Please enter input excel filename")
        return xlFilename
    
    elif variable_name == "area_name":
        print(line_begin, "area name", line_end)
        area_name = answer("Please enter the name of the entire area" + 
                           " (must match the excel sheet name)")
        return area_name
    
    elif variable_name == "regions":
        regions = {}
        print(line_begin, "regions", line_end)
        print("Define region names", 
              "(region names should be the same as excel sheet names)")
        if ws is not None:
            sort_regions(regions,ws)
        else:
            sort_regions(regions)
        return regions
    
    elif variable_name == "save_csv_names":
        print(line_begin, "output CSV file names", line_end)
        save_csv_names = {
            'map'       :'map.csv',
            'legend'    :'legend.csv',
            'overlaps'  :'overlaps.csv'
        }
        print_explain_csv(save_csv_names)
        change = y_or_n("Would you like to change csv names")
        if change:
            change_csv_names(save_csv_names)
        return save_csv_names

    elif variable_name == "save_wb_name":
        print(line_begin, "output excel workbook name", line_end)
        save_wb_name = "formatted_map.xlsx"
        print('The current input excel file name is "' + save_wb_name +'".')
        change = y_or_n("***Please note that this program will overwrite the existing file***" + 
                        "\nWould you like to change the output workbook file name?")
        if change:
            save_wb_name = answer("Please enter output workbook filename" +  
                                    " (must be different than input name)")
        return save_wb_name
    
    else:
        print("Problem: variable name does not exist")
        return('')



def define_all_variables():

    # define variables
    xlFilename = define_variable("xlFilename")
    wb = load_input_workbook(xlFilename)

    save_csv_names = define_variable("save_csv_names")
    format_map = y_or_n("\nShould the regionalized map also be formatted" + 
                        " and saved as an excel workbook?")
    save_wb_name = ''
    if format_map:
        while True:
            save_wb_name = define_variable("save_wb_name")
            if xlFilename == save_wb_name:
                print("ERROR: filename to save to cannot be the same as original workbook")
            else:
                break

    while True:
        area_name = define_variable("area_name")
        if area_name not in wb.sheetnames:
            print("ERROR: No sheet with this name was found")
        else:
            break

    if 'list' not in wb.sheetnames:
        regions = define_variable("regions")
    else:
        regions = define_variable("regions", wb["list"])

    return xlFilename, wb, save_csv_names, format_map, save_wb_name, area_name, regions



##############################################################################
# functions for regionalization
##############################################################################

def load_input_workbook(xlFilename):
    print("Now loading workbook: " + xlFilename)
    wb = load_workbook(xlFilename)
    print("Finished loading workbook: " + xlFilename)

    return wb

# returns a dictionary with info from header of each file 
# (first 2 columns and first 6 rows of a worksheet)
def get_file_header(ws):
    ncols = ws['B1'].value
    nrows = ws['B2'].value
    xllcorner = ws['B3'].value # bottom left corner
    yllcorner = ws['B4'].value # bottom left corner (column)
    cellsize = ws['B5'].value
    nodata_value = ws['B6'].value
    file_header = {
        "ncols"         :   ncols,
        "nrows"         :   nrows,
        "xllcorner"     :   xllcorner,
        "yllcorner"     :   yllcorner,
        "cellsize"      :   cellsize,
        "nodata_value"  :   nodata_value
    }
    return file_header


# returns a dictionary with basic information about top left and bottom left corners
def get_area_cell_info(area_header):

    # bottom left col coord (combined region)   --> used later for individual regions
    bottom_left_coord_to_cell_col = int(round(area_header["xllcorner"] /
                                        area_header["cellsize"])) 
    # bottom left row coord (combined region)   --> used later for individual regions
    bottom_left_coord_to_cell_row = int(round(area_header["yllcorner"] /
                                        area_header["cellsize"]))

    bottom_left_col = 1 + num_extra_left_cols
    bottom_left_row = area_header["nrows"] + num_extra_top_rows
    top_left_col = bottom_left_col
    top_left_row = bottom_left_row - area_header["nrows"] + 1 # looking for first row

    area_cell_info = {
        "bottom_left_coord_to_cell_col" :   bottom_left_coord_to_cell_col,
        "bottom_left_coord_to_cell_row" :   bottom_left_coord_to_cell_row,
        "bottom_left_col"               :   bottom_left_col,
        "bottom_left_row"               :   bottom_left_row,
        "top_left_col"                  :   top_left_col,
        "top_left_row"                  :   top_left_row
    } 
    return area_cell_info


# returns a dictionary with basic information about top left and bottom left corners
# note: rows and columns are for placement in area sheet (not individual region sheets)
def get_region_cell_info(region_header, area_header, area_cell_info):

    bottom_left_coord_to_cell_col = int(round(region_header["xllcorner"] /
                                        region_header["cellsize"])) 
    bottom_left_coord_to_cell_row = int(round(region_header["yllcorner"] /
                                        region_header["cellsize"]))

    # calculate the top left cell  
    bottom_left_col = (area_cell_info["bottom_left_col"] + 
                        (bottom_left_coord_to_cell_col - 
                        area_cell_info["bottom_left_coord_to_cell_col"]))
    
    bottom_left_row = (area_cell_info["bottom_left_row"] - 
                        (bottom_left_coord_to_cell_row - 
                        area_cell_info["bottom_left_coord_to_cell_row"]))
    
    top_left_col = bottom_left_col
    top_left_row = bottom_left_row - region_header["nrows"] + 1 # +1 because looking for first row

    region_cell_info = {
        "bottom_left_coord_to_cell_col" :   bottom_left_coord_to_cell_col,
        "bottom_left_coord_to_cell_row" :   bottom_left_coord_to_cell_row,
        "bottom_left_col"               :   bottom_left_col,
        "bottom_left_row"               :   bottom_left_row,
        "top_left_col"                  :   top_left_col,
        "top_left_row"                  :   top_left_row
    } 
    
    return region_cell_info


# this function gets the address of a cell when receiving row and col as numbers
def get_cell_address(col, row):
    col_letter = get_column_letter(col)
    address = col_letter + str(row)
    return address

# this function adds cell with overlap
def add_overlap(overlaps, row, col):
    overlaps.append(get_cell_address(col, row))
    return

# this function sets values in map_ws to the value of change_to_num, according to region_ws
# does not include blanks or nodatavalues 
def set_cells_to_num_except_blanks(map_ws, region_ws, region_header, region_cell_info, change_to_num, overlaps):
    ncols = region_header["ncols"]
    nrows = region_header["nrows"]
    nodata_value = region_header["nodata_value"]
    region_top_left_row = region_cell_info["top_left_row"]
    region_top_left_col = region_cell_info["top_left_col"]
    
    for row in range(nrows):
            for col in range(ncols):
                cell_val = region_ws.cell(row=row + num_extra_top_rows + 1, 
                                            column=col + num_extra_left_cols + 1).value
                
                # check to see if there will be overlap
                map_ws_current_row = region_top_left_row + row
                map_ws_current_col = region_top_left_col + col
                map_ws_current_val = map_ws.cell(row=map_ws_current_row, 
                                    column=map_ws_current_col).value

                if cell_val not in (nodata_value, None, ""):
                    # check to see if there will be overlap
                    if map_ws_current_val not in (None, ""):
                        add_overlap(overlaps, map_ws_current_row, map_ws_current_col)

                    # change value on map_ws
                    map_ws.cell(row=map_ws_current_row, 
                                    column=map_ws_current_col).value = change_to_num

# this function takes in information from each region
# and adds that information to map_ws
def each_region(wb, regions, map_ws, area_header, area_cell_info, overlaps, xlFilename):
    for num in regions:
        region = regions[num]

        # if given region does not exist in workbook
        if not region in wb.sheetnames:
            print("ERROR: A worksheet with region name '" + 
                region + "' does not exist in workbook '" + xlFilename +"'")
            continue
        
        region_ws = wb[region]
        region_header = get_file_header(region_ws)
        region_cell_info = get_region_cell_info(region_header, area_header, area_cell_info)
        
        set_cells_to_num_except_blanks(map_ws, region_ws, region_header, region_cell_info, num, overlaps)

        print("Now finished region: " + regions[num])
    
    return


# this function sets range of 'A1':'B6' to be equal 
# between sheets map_ws & area_ws
def set_headers_equal(map_ws, area_ws):
    # set header in map as the same as header in area
    cellsToCopy = []
    for letter in ['A','B']:
        for num in range(1,7):
            cellsToCopy.append(letter + str(num))

    for val in cellsToCopy:
        map_ws[val].value = area_ws[val].value

# this function sets blanks in map_ws that are within borders to the no data value
# information from the area sheet is used 
def set_blanks_to_nodata(map_ws, area_header, area_cell_info):
    ncols = area_header["ncols"]
    nrows = area_header["nrows"]
    nodata_value = area_header["nodata_value"]
    area_top_left_row = area_cell_info["top_left_row"]
    area_top_left_col = area_cell_info["top_left_col"]
    
    for row in range(nrows):
            for col in range(ncols):
                cell_val = map_ws.cell(row=row + num_extra_top_rows + 1, 
                                            column=col + num_extra_left_cols + 1).value
                if (cell_val == "") or (cell_val == None):
                    map_ws.cell(row=area_top_left_row + row, 
                                    column=area_top_left_col + col).value = nodata_value

# this function sets column_width to user-defined variable
# taken from https://stackoverflow.com/a/60801712
def set_column_width(map_ws):
    col_width = 2
    dim_holder = DimensionHolder(worksheet=map_ws)

    for col in range(map_ws.min_column, map_ws.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(map_ws, min=col, max=col, width=col_width)

    map_ws.column_dimensions = dim_holder

# this function sets a color scale for range within borders of map_ws
# colours and scales set by user above
def set_color_scale(map_ws, area_header, area_cell_info):
    # the following is for the color scale
    color_start_value = 00 # percentage (between 0-100)
    color_start_value_color = 'ED5F49' # light red
    color_mid_value = 60 # percentage (between 0-100)
    color_mid_value_color = 'CEE740' # yellow
    color_end_value = 100 # percentage (between 0-100)
    color_end_value_color = '22910C' # green


    colorscale_rule = ColorScaleRule(start_type='percentile', start_value=color_start_value, start_color=color_start_value_color,
                          mid_type='percentile', mid_value=color_mid_value, mid_color=color_mid_value_color,
                          end_type='percentile', end_value=color_end_value, end_color=color_end_value_color)
    
    top_left_col = area_cell_info['top_left_col']
    top_left_row = area_cell_info['top_left_row']
    bottom_right_col = area_cell_info['bottom_left_col'] + (area_header['ncols'] - 1)
    bottom_right_row = area_cell_info['bottom_left_row']

    top_left_cell_in_range = get_cell_address(top_left_col, top_left_row)
    bottom_right_cell_in_range = get_cell_address(bottom_right_col, bottom_right_row)
    range_to_format = top_left_cell_in_range + ":" + bottom_right_cell_in_range
    
    map_ws.conditional_formatting.add(range_to_format, colorscale_rule)


# this function formats map_ws
def format_map_ws(map_ws, area_ws, area_header, area_cell_info, format_map):
    # set header of map_ws to be the same as area_ws
    set_headers_equal(map_ws, area_ws)
    # set to nodata value (of area) if blank
    set_blanks_to_nodata(map_ws, area_header, area_cell_info)

    if format_map:
        # set column width
        set_column_width(map_ws)
        # conditional format cells
        set_color_scale(map_ws, area_header, area_cell_info)

# this function prints the region names and numbers into legend_ws
def create_legend(legend_ws, regions):
    row = 1
    col1 = 1
    col2 = 2

    legend_ws.cell(row=row, column=col1).value = "region number" 
    legend_ws.cell(row=row, column=col2).value = "region abbreviation"

    for num in regions:
        row += 1
        legend_ws.cell(row=row, column=col1).value = num 
        legend_ws.cell(row=row, column=col2).value = regions[num]

# this function prints the region names and numbers into legend_ws
def print_overlaps(overlaps_ws, overlaps):
    row = 1
    col = 1

    overlaps_ws.cell(row=row, column=col).value = "list of cells with overlaps" 

    for cell in overlaps:
        row += 1
        overlaps_ws.cell(row=row, column=col).value = cell


# this function saves the sheets 'map' and 'legend' and 'overlaps' into new workbook
def save_files(wb, save_wb_name, save_csv_names, format_map):
    
    # save sheets into csv file
    # using code inspired by https://stackoverflow.com/a/10803229
    save_sheets = ['map', 'legend', 'overlaps']
    for ws_name in save_sheets:
        print("Now saving " + ws_name + " to: " + save_csv_names[ws_name])

        with open(save_csv_names[ws_name],'w', newline='') as file:
            writer = csv.writer(file)
            for row in wb[ws_name].rows:
                writer.writerow([cell.value for cell in row])
    
    # save only sheet 'map' by removing other sheets
    # code copied from https://stackoverflow.com/a/46237894
    if format_map:
        sheets = wb.sheetnames
        for ws_name in sheets:
            if (ws_name != 'map'):
                wb.remove(wb[ws_name])

        print("Now saving formatted map to workbook: " + save_wb_name)
        wb.save(save_wb_name)
        print("Everything has been saved")
    





##############################################################################
# global variables
##############################################################################
num_extra_top_rows = 6 # the top of every asc file has 6 extra rows
num_extra_left_cols = 1 # the left of every asc file has 1 extra column
   


##############################################################################
# main script
##############################################################################



def main():
    line_end = "===================="
    line_begin = "\n" + line_end

    xlFilename, wb, save_csv_names, format_map, save_wb_name, area_name, regions = define_all_variables()
   
    # create sheets in workbook for regionalized map, legend, and overlaps
    map_ws = wb.create_sheet("map")
    legend_ws = wb.create_sheet("legend")
    overlaps_ws = wb.create_sheet("overlaps")
    
    # area worksheet
    area_ws = wb[area_name]
    area_header = get_file_header(area_ws)
    area_cell_info = get_area_cell_info(area_header)

    # list to keep track of cells with overlaps
    overlaps = []

    # go through each region and add to map_ws
    print(line_begin, "Making the regionalization map", line_end)
    each_region(wb, regions, map_ws, area_header, area_cell_info, overlaps, xlFilename)   

    # format map worksheet to have no data value, color scale, and smaller column width
    format_map_ws(map_ws, area_ws, area_header, area_cell_info, format_map)
    
    # create legend
    create_legend(legend_ws, regions)

    #print overlaps
    print_overlaps(overlaps_ws, overlaps)

    print(line_begin, "Saving files", line_end)
    # save workbook
    save_files(wb, save_wb_name, save_csv_names, format_map)
    
    return


main()


