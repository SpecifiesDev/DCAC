import openpyxl



def parseRoster(fileName):
    #Load the roster file
    file = openpyxl.load_workbook(fileName)

    sheet = file.active

    """
    Here, we create a cells variable.
    The first value being the first row to read, the last value being the last row to read.
    For any member using this to iterate in the future, change the x/y to the amount of rows existing on the sheet (past The title // data headers)
    """
    #the starting number, due to headers
    x = 3
    #last row
    y = sheet.max_row
    print("There are currently " + str(y) + " members displayed on the roster.")

    #cell structure
    startCellString = "A" + str(x)
    endCellString = "E" + str(y)
    cells = sheet[startCellString: endCellString]
    
    #lists
    names = []
    number = []
    email = []
    section = []
    
    for c1, c2, c3, c4, c5 in cells:
        #null check because people like to add blank rows
        if c1.value is None:
            pass
        else:
            #String together their full name
            fullName = c1.value + " " + c2.value
            names.append(fullName)

            #Add the member's number
            number.append(c3.value)

            #Add the member's email
            email.append(c4.value)

            #Add what they do in the group
            section.append(c5.value)
    toWriteFile = open("formattedData.txt", "w")
    toWriteFile.write("DCAC Roster Information \n")
    toWriteFile.write("\n")
    toWriteFile.write("Current Members: " + str(y))

    drawing = 0
    music = 0
    film = 0
    
    #loop sections
    for x in range(0, len(section)):
        dataString = section[x]
        if dataString == "Film":
            film = film + 1
        if dataString == "Music":
            music = music + 1
        if dataString == "Drawing":
            drawing = drawing + 1
    toWriteFile.write("\n")
    toWriteFile.write("Section information: ")
    toWriteFile.write("\n")
    toWriteFile.write("Drawing: " + str(drawing))
    toWriteFile.write("\n")
    toWriteFile.write("Music: " + str(music))
    toWriteFile.write("\n")
    toWriteFile.write("Film: " + str(film))
    toWriteFile.write("\n")
    
    #loop all of it
    for x in range(0, len(names)):
        nameString = names[x]
        numberString = number[x]
        emailString = email[x]
        sectionString = section[x]

        outputString = "Displaying data for the member: " + nameString + ". Number: " + numberString + " Email: " + emailString + " They are in the section: " + sectionString
        print(outputString)
        #Now we output data to a formatted file.
        toWriteFile.write(" \n")
        toWriteFile.write("Member #" + str(x + 1) + "'s Data:")
        toWriteFile.write("\n")
        toWriteFile.write("Full Name: ")
        toWriteFile.write(nameString)
        toWriteFile.write("\n")
        toWriteFile.write("Number: ")
        toWriteFile.write(numberString)
        toWriteFile.write("\n")
        toWriteFile.write("Email: ")
        toWriteFile.write(emailString)
        toWriteFile.write("\n")
        toWriteFile.write("Section: ")
        toWriteFile.write(sectionString + "\n")
        toWriteFile.write("\n")
        
parseRoster('roster.xlsx')
