from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import xlrd
import xlsxwriter
import time
import re


CREDENTIALS_EXCEL_FILE = r'C:\Tars2\Credentials.xlsx'
ADDRESS = "127.0.0.1:4000"
NAME_REGEX = r"\</b\>([a-zA-Z0-9]+)"
TIME_REGEX = r"\</td\>\<td>([0-9]+\:[0-9]+)"
COORDINATES_REGEX = r"\<\/td\>\<td\>(X \: [0-9]+\.[0-9]+ \, Y \: [0-9]+\.[0-9]+)"
UPDATED_EXCEL_FILE_PATH = r"C:\Tars2\Updated.xlsx"
GOOGLE_MAPS_WEB = r"https://www.google.com/maps"


def read_chart():

    """
    The function loads an excel file from known destination reads it's content.
    Finally the function returns a dictionary with the row as index and the rest of the content of that row as value.
    """
    my_dictionary = {}  # Creating new empty dictionary.

    # Opening the excel file from it's location

    wb = xlrd.open_workbook(CREDENTIALS_EXCEL_FILE)
    sheet = wb.sheet_by_index(0)

    # Running in a loop from 1 to the number of rows.
    for index in range(1, sheet.nrows):

        # Adds the index as a key and all the row content as value.
        my_dictionary.update({index: sheet.row_values(index)})

    return my_dictionary


def get_credentials(index, dictionary):
    """
    The function gets index as a key and returns the values of that key from the given dictionary.

    :param index:
    :param dictionary:
    :return: Username and password from the dictionary.
    """

    # Running on every key and value in the given dictionary.
    for key, value in dictionary.items():
        if key == index:  # If one of the keys equals to the given index.
            username, password = value[0], value[1]   # Insert the values to the parameters.
            break  # Quits the loop.
    else:  # If the break command was not executed.
        raise IndexError("Index was not found in the given dictionary.")  # Raise a new error.

    return username, password


def login(username, password):

    browser = webdriver.Chrome()
    print("http:\\{}".format(ADDRESS))
    try:
        browser.get("http:\\{}".format(ADDRESS))
    except Exception:
        print("Web is unreachable...")
    username_element = browser.find_element_by_name('username')
    password_element = browser.find_element_by_name('password')

    username_element.send_keys(username)
    password_element.send_keys(password)

    browser.find_element_by_name('Log in').click()
    time.sleep(3)

    html_code = browser.page_source
    return html_code


def get_info(user, html_code, my_dictionary):

    """
    Search with regexes the media username, the coordinates and the last logon time.
    :param user:
    :param html_code:
    :param my_dictionary:
    :return: dictionary with all parameters as value and the user as key.
    """
    media_name = re.search(NAME_REGEX, html_code).group(1)
    coordinates = re.search(COORDINATES_REGEX, html_code).group(1)
    login_time = re.search(TIME_REGEX, html_code).group(1)

    my_dictionary.update({user: [media_name, login_time, coordinates]})
    return my_dictionary


def commit(dictionary):
    """
    Writing all parameters from the dictionary to the excel file.
    :param dictionary:
    :return: None
    """
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(UPDATED_EXCEL_FILE_PATH)
    sheet1 = workbook.add_worksheet()

    # Adding titles.
    sheet1.write(0, 0, "User")
    sheet1.write(0, 1, "Social Media Name")
    sheet1.write(0, 2, "Last Logon time")
    sheet1.write(0, 3, "Coordinates")

    row_number = 1

    # Running on the keys and values of the dictionary.
    for key, value in dictionary.items():

        # Writing the username
        sheet1.write(row_number, 0, key)

        # Transferring the values to list so we can take the index of each value in the list.
        items = list(value)
        for item in items:

            # Getting the index and by that creating new position to write to the chart.
            position = items.index(item) + 1
            sheet1.write(row_number, position, item)

        row_number += 1

    # Closing the workbook and prints that the data was inserted.
    workbook.close()
    print("\n\nWrote the data to the excel file in the path: {}".format(UPDATED_EXCEL_FILE_PATH))


def get_location(username):
    """
    Takes the coordinates of the given user (From the excel file) and gets his location by google maps.
    :param username:
    :return: None
    """
    wb = xlrd.open_workbook(UPDATED_EXCEL_FILE_PATH)
    sheet = wb.sheet_by_index(0)

    # Running in a loop from 1 to the number of rows.
    for index in range(1, sheet.nrows):

            # If the username equals to one of the user cells.
            if username == sheet.row_values(index)[0]:
                # Takes the coordinates value from the user.
                coordinates_string = sheet.row_values(index)[3]
                # Splits it and take only the numeric values.
                items = coordinates_string.split(" ")
                print(items)
                final_coordinates = "{} , {}".format(items[2], items[-1])
                break
    else:  # If we didn't get to the break command
        final_coordinates = ""

    if final_coordinates:
        browser = webdriver.Chrome()
        # Trying to get to google maps web.
        try:
            browser.get(GOOGLE_MAPS_WEB)
        except Exception:
            print("Web is unreachable...")
        try:
            # Let the website load himself.
            time.sleep(5)

            # Search for the search box input field and send it the coordinates we found earlier.
            search_field = browser.find_element_by_id(id_="searchboxinput")
            search_field.send_keys(final_coordinates)
            # Clicking the search button.

            browser.find_element_by_id(id_="searchbox-searchbutton").click()

            # Loads and takes screenshots.
            time.sleep(2)
            browser.save_screenshot(r"C:\Tars2\{}.png".format(username))

        except Exception:
            print("ERROR")

    else:
        print("\n\nUsername was not found...")


def main():

    my_dictionary = {}
    print(read_chart())
    for i in range(1, 4):
        user, password = get_credentials(i, read_chart())
        html = login(user, password)
        my_dictionary = get_info(user, html, my_dictionary)
        commit(my_dictionary)
    user = input("Please enter the username you are looking for: ")
    while user != "exit":
        get_location(user)
        user = input("Please enter the username you are looking for: ")
    print("\n\n^^^^^^^   PROGRAM IS DONE   ^^^^^^^")


if __name__ == '__main__':
    main()


