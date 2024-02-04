#!/usr/bin/env python
# -*-coding:utf-8 -*-
'''
@File    :   Password_Generator
@Time    :   2023/04/01 10:09:52
@Author  :   Matt Lester 
@Version :   1.0
'''

import string
import random as rand
import xlwings as xw

def removeSpecial(characters_to_remove):                                                                            # Accepts a list of characters as an argument and compares each list item to the special character list
                                                                                                                    # if value is in special character list than that value is removed from special character list.
    for value in characters_to_remove:                                                                              # No value is returned as special character list is modified in place
        
        if value in specialChar:
            specialChar.remove(value)


# file = "SomeFileName.xlsx"                                                                                        # File name if in same directory as this script. ex) someFileName.xlsx
file_location = 'C:\\SomeFileLocation\\SomeFileName.xlsx'                                                           # File path, allows the script to access file no matter the location. ex) C:\\someFilePath\\someFileName.xlsx
                                                                                                                    # Use double \\ or raw strings. Using single \ with out raw string will cause python to recognize it as an indicator.

# wb = xw.Book(file, password= "YourPassword")                                                                      # Connects to file that is in current working directory, password only required if file is password protected
wb = xw.Book((file_location), password = "YourPassword")                                                            # Connects to file using raw string and direct file path. This allows the file to be accessed without the file open in the same directory
sheet = wb.sheets[0]                                                                                                # Instantiates sheet object

numPasswords = sheet['A1'].current_region.last_cell.row                                                             # gets current number of passwords in the excel file by looking at a specific column and last row of that column

specialChar = ["!", "@", "#", "$", "%", "&", "_", "-"]                                                              # List of special characters, used for creating password
number = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]                                                         # list of numbers used for creating password
shuffleNum = rand.randint(3, 25)                                                                                    # Randomly selects the number of times to shuffle the characters used in the password. Inclusive of both arguments.
                                                                                                                    # Least amount of time it will shuffle is 3, most is 25
                                                                                            


key = input("What do you need a password for? ")                                                                    # Takes in user input for what the password will be for

print('Are there any special characters listed below that cannot be included?\n')                                   # Asks user if there are any special characters that cannot be included and then for loop lists out
                                                                                                                    # current list of special characters.
for character in specialChar:
    print(character)

deleteSpecialChar = input("\nPlease type Y for yes and N for no and hit Enter: ")                                   # Assigns variable of y or n depending on user input based on pervious question

if deleteSpecialChar.lower() == "y":                                                                                # If user answers yes, a prompt asking for the excluded special characters will be given and the desired format for entering those characters.
    
    charToRemove = input("Please enter the characters that cannot be included in this format: 1,2,3\n")
    remove_list = charToRemove.split(",")                                                                           # Creates a list of special characters to be removed    

    removeSpecial(remove_list)                                                                                      # Calls removalSpecial function to remove excluded characters from specialChar list                                


check = True                                                                                                        # Check variable for while loop
  
                                                                                                                    # While loop creates password and checks that there are uppercase characters. If there aren't then it will re-create a password.
while check == True:                                                                                                # Once it has a password with capital letters included, it will set the while loop check to False and the loop will end.
    
    passwordList = []                                                                                               # Creating empty list to store password characters
    password = ""                                                                                                   # Creating empty string for password assignment
    
    for i in range (19):                                                                                            # Randomly selects characters for the password and appends them to password list. 
                                                                                                                    # Selects two special characters and two numbers and the rest are letters
        if i < 2 :                                                                                                  # Currently there is a potential that all letters are lower case..... need to fix
                                                                                                                    # Values are stored in the order they are selected. The list will always be [special, number, special, number, letters...]
            passwordList.append(specialChar[rand.randint(0, len(specialChar))])
            passwordList.append(number[rand.randint(0, 9)])
        else:
            passwordList.append(string.ascii_letters[rand.randint(0, 51)])
        

    for j in range (shuffleNum):                                                                                    # Shuffling the password list so that the characters are not always in the format stated above.
                                                                                                                    # Number of shuffle is determined by shuffleNum previously established.
        rand.shuffle(passwordList)

    for k in range(len(passwordList)):                                                                              # Writes characters in password list to string in preparation for output
        
        password = password + passwordList[k]
    
    
    if password == password.lower():
        check = True
    else:
        check = False        

print(f"\nHere is your password for {key}:  {password} \nIt has been saved in {file_location}\n")

sheet[f"A{numPasswords + 1}"].value = key                                                                           # Writes "key" value determined by user input to the next available cell in column A
sheet[f"B{numPasswords + 1}"].value = password                                                                      # Writes password value to next available cell in column B

wb.save()
wb.close()

