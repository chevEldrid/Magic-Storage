# imports the io library
import io
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook


# INPUT: txt file of magic cards
# OUTPUT: a list with each card's name repeated the number of times that card is in the deck
def file_to_list(filename):
    f = open(filename, "r")
    # list to hold temporary list of cards
    temp = []
    for line in f:
        # I think there was an empty line at the end of the file so this exception clause helps that
        try:
            num = int(line[:1])
        except ValueError:
            num = 0
        for i in range(num):
            temp.append(line[2:(len(line) - 1)])
    # file_contents = file.read()
    # print file_contents
    file.close(f)
    return temp


# INPUT: The list of cards, and the name of the deck
# SIDE EFFECTS: Adds a new sheet for a deck to the Spreadsheet
def add_deck(card_list, deck_name):
    # creates a Card Library excel doc, and if one deoesn't exist. Makes a new one
    try:
        wb = load_workbook("Card Library.xlsx")
    # catches specific openpyxl exception to cannot find excel book
    except openpyxl.shared.exc.InvalidFileException:
        wb = Workbook()
    # checks to make sure deck doesn't already exist
    try:
        ws = wb[deck_name]
    # catches error for that worksheet not existing
    except KeyError:
        ws = wb.create_sheet(0, deck_name)
    # Deck header...
    ws['A1'] = "Cards"
    # sets parameters to then iterate through card list to put all in excel doc
    row = 2
    col = 'A'
    for card in card_list:
        pos = col + str(row)
        ws[pos] = card
        row += 1
    wb.save('Card Library.xlsx')


def update_card_library():
    deck_file_name = raw_input("What's the file name of the deck?: ")
    deck_name = raw_input("And what do you want to call your deck?: ")
    try:
        add_deck(file_to_list(deck_file_name), deck_name)
        print "deck added successfully!"
    except Exception:
        print "Oops! That deck was invalid! Please check the supposed filename of this deck!"


def load_library():
    # creates a Card Library excel doc, and if one deoesn't exist. Makes a new one
    try:
        wb = load_workbook("Card Library.xlsx")
        for sheet in wb.worksheets:

    # catches specific openpyxl exception to cannot find excel book
    except openpyxl.shared.exc.InvalidFileException:
        print "You don't have a card library yet! Try adding some decks!"


# actual program
while True:
    response = raw_input('What would you like to do?: ').lower()
    if response == "add a deck":
        update_card_library()
    if response == "modify a deck":
        print "If this deck cannot be found in the Card Library, it will be added automatically!"
        update_card_library()
    if response == "exit":
        break
    if response == "help":
        print "Commands available in this program are: add a deck, modify a deck, exit"
    if response == "load library":
        load_library()
    else:
        print "Could not recognize that command! Type 'help' if stuck!"

# insert other if conditions here...


print "test complete!"
