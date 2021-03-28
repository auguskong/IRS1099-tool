import pyautogui as pgui
import pyperclip as pclip
import time

x_coordinates_dict = {
    "date_acquired": 600,
    "date_sold": 650,
    "proceeds": 700,
    "cost_basis": 780,
    "code_from_instructions": 840,
    "amount_of_adjustment": 900,
    "gain_or_loss": 980
}

def fill_8949_form():
    first_row_y_coordinate = 420
    row_interval = 26
    num_rows = 14 # form 8949
    
    curr_data_index = 28
    curr_row_index = 0

    x_coordinates = [500, 600, 650, 700, 780, 840, 900, 980]
    
    pgui.click(x_coordinates[0], first_row_y_coordinate, interval=0.5)
    wash_sale_data_keys = list(wash_sale_data[0])

    print("Start filling the form")
    while curr_row_index < num_rows:
        curr_row_y_coordinate = first_row_y_coordinate + row_interval * curr_row_index
        print("filling form row", curr_row_index + 1, " data index is: ", curr_data_index)
        for col_idx, col_x_coordinate in enumerate(x_coordinates):
            pgui.click(col_x_coordinate, curr_row_y_coordinate, interval=0.2)
            time.sleep(1)
            pclip.copy(wash_sale_data[curr_data_index][wash_sale_data_keys[col_idx]])
            pgui.hotkey('command', 'v')
        curr_row_index += 1
        curr_data_index += 1
    print("complete filling the form")
