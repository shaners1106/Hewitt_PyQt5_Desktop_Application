# Shane Snediker
# Hewitt Research Foundation Desktop Application
# File last updated: 6-22-2021

# This Python library that works in conjunction with a Windows32 API. The
# library provides functionality of sending print jobs directly to the default
# printer of the computer running the program.
# 
# I have discovered that for many Python libraries, documentation is lacking,  
# particularly all libraries involved in printing.  In general, finding anything
# regarding highly specialized printing is not easy.  However, the following link
# is where we found the win32ui library:
# 
# https://www.oreilly.com/library/view/python-programming-on/1565926218/ch10s03.html

# The flow of this script is as follows:

#       1. Configure database connection
#       2. Extract necessary scantron information for each student test
#       3. Use Python's Windows library win32ui to instantiate a print document 
#       4. Use a loop to configure each individual student scantron exam
#       7. Update the date_printed attribute for each of these student tuples to today's date 

# Import libraries
from PyQt5.QtCore import QObject # PyQt5 functionality
import sys                # Needed for configuring the database connection
import pyodbc             # Connect to MS SQL database
from configparser import ConfigParser # Separate configuration settings into separate file
import win32ui            # Pull in the Python Windows32 libraries
import win32con           # Font modifications
from datetime import date # Import functionality for accessing today's date

# This class comprises the full scantron test printing script
# Define this script as a class to allow the desktop application to properly utilize it
class print_tests(QObject):

    # Scantron print class MEMBER VARIABLES

    # Flag to communicate with the main GUI in the event that an error occurs during the
    # execution of this script
    exception_thrown = False
    # Flag to signify whether or not Hewitt needs to reprint student exams
    need_reprints = False
    # In the event that Hewitt does need to reprint scantrons, we will ask Hewitt to provide 
    # the first and last name of the last student whose test printed accurately
    first_name = ''
    last_name = ''
    # Instantiate a variable to hold the location within the student list where printing will
    # begin (for a new print this will be 0, or in other words we'll begin at the beginning
    # of the student list, but for a reprint situation, the reprint index will be 1 past the
    # index of the last student whose test successfully printed)
    reprint_index = 0

    def __init__(self, parent=None): 
        super(print_tests, self).__init__(parent)
        
    # The mother function: pulls all student scantron info from Hewitt's database view and 
    # instantiates win32ui print documents for each one and prints them out
    def print_scantrons(self):   
        try:
            # Let's access today's date so that every time Hewitt runs this print script to generate printing
            # templates, the pdf that is rendered can be saved with the corresponding date on which it is generated
            # The date.today() function call captures today's date in YYYY-MM-DD format while the call to strftime
            # alters the format and saves it in a string.  The following is an example of the format that we have
            # decided to set up to be saved in the filename of each print template : 05-10-2021
            today = date.today().strftime("%m-%d-%Y")        

            #################### MAKE A DATABASE CONNECTION ##########################################
            
            config = ConfigParser()
            # This will determine if its the .exe or just the script
            if 'dist' in sys.path[0]: config.read('config.ini')
            else: config.read(sys.path[0] + '\\config.ini')
            # Transforms the config infor into a dictionary for the program to use
            config_dict = dict(config.items())

            # Connects to the database
            hewitt_db = pyodbc.connect(
            # NOTE: confidential database credentials omitted for security purposes
            )
            ##########################################################################################
            
            ####### CAPTURE ALL OF THE STUDENT INFO FOR THE TESTS THAT NEED TO BE PRINTED ############

            # The cursor function is returns a control structure that enables the 
            # execution of queries against the DB
            my_cursor = hewitt_db.cursor()
            # Pull in all current tests that still need to be printed
            # NOTE: the print_tests VIEW from the Hewitt DB always contains tests that need 
            #       to be printed, so selecting everything from this VIEW is what we want

            my_cursor.execute('SELECT * FROM print_tester') 
            # Save the tuple information in a variable that we can iterate over.  The fetchall()
            # function returns a tuple of tuples (the internal tuples consisting of individual
            # test order attributes)
            current_tests = my_cursor.fetchall()

            # Now that we have a data structure holding all of our student data, we can begin
            # to work with the data. First, how many tests does this print job contain?
            num_tests = len(current_tests)
            
            # How many attributes does each test tuple have?
            num_attributes = len(current_tests[0])
            
            ########### CONVERT NULL VALUES TO EMPTY STRINGS ###################################

            # Python hates null values (or what they term 'None types')
            # We begin interacting with the data by converting all nulls to empty strings, which
            # is something that Python can deal with
            # We need to use an array in order to rebuild the tuple
            no_null_attributes = []
            # For each individual tuple:
            for ind_tup in range(num_tests):
                # check it for nulls and if found, convert to '':
                # However modifying an existing tuple is not permitted in Python, so we'll have to
                # do some extra work to create new tuples.  This is necessary because if this script
                # ever assigns an attribute that has a null value, it will crash the script
                # Track the attributes within this tuple that hold null values
                null_indices =[]
                for attribute in range(num_attributes):
                    # Does this index hold a null value?
                    if current_tests[ind_tup][attribute] is None:
                        # Then save this index number
                        null_indices.append(attribute)
                # Instantiate an inner array to hold this tuple's attributes
                inner_tuple = []
                # For each index, if it is an index number that held a null value, append an empty
                # string into the new array, otherwise append the original value
                for index in range(num_attributes):
                    if index in null_indices:
                        inner_tuple.append("")
                    else:
                        inner_tuple.append(current_tests[ind_tup][index])
                # Finally, add this inner array as a newly created null-less tuple into the new outer array    
                no_null_attributes.append(inner_tuple)
                
                # Now we have a Python array corresponding to the data tuple with null values replaced by empty strings
            ###############################################################################################
            
            # NOTE: The win32ui TextOut function by which we send textual elements to the page only
            #       accepts string values within the function call.  Therefore, even the attributes
            #       that are ints must be stored as string values

            # Now that we know how many tests to print, let's review the tuple info. The indices
            # of each tuple are as follows:

            #   [0] : group_id      NOTE: this is an int data type and will need to be converted to a string in order to print to the win32ui document
            #   [1] : test_id       NOTE: this is an int data type and will need to be converted to a string in order to print to the win32ui document
            #   [2] : account_id    NOTE: this is an int data type and will need to be converted to a string in order to print to the win32ui document
            #   [3] : student_id    NOTE: this is an int data type and will need to be converted to a string in order to print to the win32ui document
            #   [4] : student_first_name
            #   [5] : student_last_name
            #   [6] : grade         NOTE: this is an int data type and will need to be converted to a string in order to print to the win32ui document
            #   [7] : customer1_first_name
            #   [8] : customer2_first_name
            #   [9] : customer1_last_name
            #   [10] : customer2_last_name
            #   [11] : address1
            #   [12] : city1 
            #   [13] : state1
            #   [14] : zipcode 1
            ###########################################################################################

            # Now we can define some const values that correspond to cartesian coordinates
            # that we will use to place our printable objects in the correct location on the page

            ################ CONSTANT VARIABLE DECLARATIONS ########################################

            # Side length of the square object that will be printed into the grade bubble
            # and onto the boxes in the "office use only" section
            SQUARE_LEN = 50

            # Vertical coordinates (Y location) of the student info
            GROUP_NUM = 270
            STUDENT_NUM = 410
            STUDENT = 550
            STU_GRADE = 690
            PRINT_DATE = 830
            PARENTS1 = 1110
            PARENTS2 = 1245
            PARENTS3 = 1380 

            # Horizontal coordinates (X location) of the specific student tuple attributes
            # NOTE: these attributes are left justified at the same x coordinate value :
            THIS_STUDENT_DATA = 770

            # Non transferrable statement coordinates
            NOT_TRANSFER_X = 400
            NOT_TRANSFER_Y = 1670

            # Grade dot coordinates
            # Vertical alignment coordinates for grade level dots
            GRADE3 = 2035 
            GRADE4 = 2135
            GRADE5 = 2235
            GRADE6 = 2335
            GRADE7 = 2435
            GRADE8 = 2535
            # Horizontal alignment coordinate value (all grade dots are vertically aligned
            # within 1 horizontal column).  This value is the left side of the dot
            GRADE_COLUMN = 1220

            # Office use only coordinates
            # NOTE: These coordinates have been dialed in and confirmed on Hewitt's Ricoh printer
            #       as of May 15th, 2021
            # There are 14 rows and 10 columns within the "office use only" section on the scantron
            # The following are the horizontal coordinates (x values) for the columns and
            # the vertical coordinates (y values) for the rows:
            # The vertical (y) row coordinate values
            ROW0_V = 4944   
            ROW1_V = 5044   
            ROW2_V = 5144   
            ROW3_V = 5244   
            ROW4_V = 5344   
            ROW5_V = 5444   
            ROW6_V = 5544   
            ROW7_V = 5644   
            ROW8_V = 5744   
            ROW9_V = 5844   
            ROW10_V = 5944  
            ROW11_V = 6044  
            ROW12_V = 6144  
            ROW13_V = 6244  
            # The horizontal (x) column coordinate values
            COLUMN0_H = 309     
            COLUMN1_H = 409     
            COLUMN2_H = 499     
            COLUMN3_H = 599     
            COLUMN4_H = 699     
            COLUMN5_H = 805     
            COLUMN6_H = 905     
            COLUMN7_H = 995     
            COLUMN8_H = 1100
            COLUMN9_H = 1205    

            ########################################################################################
            
            # Next we define the pieces of text that will need to be printed within the top left
            # corner of the print template. The following is the layout of the top corner student
            # information:
            #
            #   Group ID:       specific student group id (if student is a part of a group)
            #   Student ID:     specific student's id #
            #   Student:        specific student's first and last name
            #   Grade:          specific student's grade level
            #   Date Printed:   date that this specific test was printed on
            #   Parents:        specific student's parents' first and last name and address

            ############### TEXT OBJECT VARIABLE DEFINITIONS #######################################
            
            # Top Corner Student info header variables
            group_id_header = "Group ID:"
            student_id_header = "Student ID:"
            student_header = "Student:"
            grade_header = "Grade:"
            date_header = "Date Printed:"
            parents_header = "Parents:"

            # Not Transferrable variable
            not_transferrable = "Not Transferrable to Any Other Parent or Student"
            ###############################################################################

            # Now we can define the functions that will be used in this file

            ############## FUNCTION DEFINITIONS ###########################################

            # This function uses a mathematical calculation to convert the desired standard
            # font size to the scale of the way that win32ui prints to the page.
            # args: dc : the win32ui object that the font will be used on
            #       PointSize : the desired standard font size
            # We found this function at the following URL:
            # https://stackoverflow.com/questions/48549555/how-to-set-font-type-and-size-for-printing-using-windows-gdi
            def getfontsize(dc, PointSize):
                inch_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)
                return int(-(PointSize * inch_y) / 72)
            ###################################################################################
            
            # Is this a reprint iteration?
            if(self.need_reprints):

                ############ FIND RESTART INDEX ################################################
                # Start a loop to find which tuple Hewitt's print job ended on (they will have 
                # just entered the first and last name of the student whose test was the last
                # one to successfully print)
                for iter in range(num_tests):
                    # If the first and last name entered by Hewitt matches this tuple, save the index
                    # NOTE: Python's casefold() function changes a string to all lowercase characters. 
                    #       This is a good way of avoiding errors involving case inconsistencies which
                    #       are a common thing when working with fallable user input
                    if(no_null_attributes[iter][4].casefold() == self.first_name.casefold() and no_null_attributes[iter][5].casefold() == self.last_name.casefold()):
                        # And increment from this index 1 person
                        self.reprint_index = iter + 1
                        break
                ################################################################################
            
            # It is now time to instantiate print documents.  Each student exam will need to 
            # have its own Python win32ui document object because we looked for a long time 
            # searching for the functionality of multi-page documents to no avail.  It may
            # be possible, but the documentation for the win32ui library is so scarce that
            # we are still not sure.  So we chose to create a for loop that will run 1 time
            # for each scantron
            
            for test in range(self.reprint_index, num_tests):
                ################ CAPTURE DATABASE TUPLE ATTRIBUTES ###############################
                
                # Populate this student's top left info corner with corresponding data
                # The group_id is each tuple's 0th index
                group_id = str(no_null_attributes[test][0])  # Cast int to string
                # The student id line on the scantron is composed of the student's account id
                # + individual student id.  The account_id and student_id attributes are each
                # tuple's 2nd and 3rd indices
                student_id = str(no_null_attributes[test][2]) + '-' + str(no_null_attributes[test][3]) # Cast ints to strings
                # Student first and last name attributes are stored in the 4th and 5th indices
                student_name = no_null_attributes[test][4] + ' ' + no_null_attributes[test][5]
                # The student grade is held in the 6th index
                grade = str(no_null_attributes[test][6])
                # We use Python's built in date function to print today's date
                date_printed = today
                
                # This is where we have to get a little cute with how we parse the parent information
                # Living in a day and age where children often have parents with different last names,
                # we will build in a filter that will check to see if the 2 parent/guardians have 
                # different last names.  If so it will print the full names of both parents, else it will
                # print the first names of each parent followed by their last names
                # Example of different last names:  Joe Hewitt & Sally Hewittson
                # Example of parents with same last name: Joe & Sally Hewitt
                # NOTE : the layout for the parent info lines is as follows:
                #   line 1: parent names
                #   line 2: street address
                #   line 3: city, state and zipcode
                
                # If there's only 1 parent or guardian in the system
                if ((no_null_attributes[test][8] == "") and (no_null_attributes[test][10] == "")):
                    parents_line1 = no_null_attributes[test][7] + ' ' + no_null_attributes[test][9]
                    parents_line2 = no_null_attributes[test][11]  
                    parents_line3 = no_null_attributes[test][12] + ' ' + no_null_attributes[test][13] + ', ' + no_null_attributes[test][14]
                
                # If customer1's last name matches customer2 or customer2 has a null or none value for last name
                elif no_null_attributes[test][9] == no_null_attributes[test][10] or no_null_attributes[test][10] == "":
                    if no_null_attributes[test][8] != "":
                        parents_line1 = no_null_attributes[test][7] + ' & ' + no_null_attributes[test][8] + ' ' + no_null_attributes[test][9]
                        parents_line2 = no_null_attributes[test][11]  
                        parents_line3 = no_null_attributes[test][12] + ' ' + no_null_attributes[test][13] + ', ' + no_null_attributes[test][14]
                    else:
                        parents_line1 = no_null_attributes[test][7] + ' ' + no_null_attributes[test][9]
                        parents_line2 = no_null_attributes[test][11]  
                        parents_line3 = no_null_attributes[test][12] + ' ' + no_null_attributes[test][13] + ', ' + no_null_attributes[test][14]

                # The parents/guardians have different last names, print them accordingly
                else:
                    parents_line1 = no_null_attributes[test][7] + ' ' + no_null_attributes[test][9] + ' & ' + no_null_attributes[test][8] + ' ' + no_null_attributes[test][10]
                    parents_line2 = no_null_attributes[test][11]  
                    parents_line3 = no_null_attributes[test][12] + ' ' + no_null_attributes[test][13] + ', ' + no_null_attributes[test][14]
                ##########################################################################################
                
                #################### BEGIN TO BUILD THE WIN32UI DOCUMENT #################################

                # Now that we've captured the data and populated the appropriate variables, we can now
                # populate the document with that information

                # Instantiate a win32ui object that we can use to generate printable objects
                dc = win32ui.CreateDC()
                # Connect the win32ui object to the printer
                dc.CreatePrinterDC()
                # Begin a new document and initialize the page
                dc.StartDoc('Scantron Print Template')
                dc.StartPage()

                # Now we define the necessary fonts that will be used within the template

                ################ DEFINE FONTS #################################################

                # Make the text size 11
                fontsize = getfontsize(dc,10)
                # Create a bold font for the student info headers (left column of the student info section) as well as the non-transferrable statement
                fontdata_headers = { 'name':'Arial', 'height':fontsize, 'italic':False, 'weight':win32con.FW_BOLD}
                # Create a win32ui bold font object and save it in a variable that we can utilize throughout this document
                bold_font = win32ui.CreateFont(fontdata_headers)
                # Create a normal-weight font for the specific student information (right column of the student info section)
                fontdata_stud_info = { 'name':'Arial', 'height':fontsize, 'italic':False, 'weight':win32con.FW_NORMAL}
                # Create a win32ui normal font object and save it in a variable that we can utilize throughout this document
                normal_font = win32ui.CreateFont(fontdata_stud_info)

                ###############################################################################

                # The 3 core functions of this program (TextOut, MoveTo, and LineTo)
                # take as arguments cartesian coordinates that tell the function where
                # on the page to implement the specified functionality.  The Cartesian 
                # coordinate function arguments follow this patter: (x location, y location),
                # where x covers the horizontal space of the page and y covers the vertical space
                # NOTE: within the win32ui application, (0,0) is located in the upper left 
                #       corner of the page

                # NOTE: we have chosen to pursue the structure of building this document from the 
                #       top down.  This is how a printer prints onto a page, so in the absence of
                #       adequate documentation, it is our perception that this will be the best 
                #       way to construct this template.  This will, however, create more code
                #       because we will have to switch fonts twice per line in the student info
                #       section because the headers are bold and the specific student info is not

                ################### STUDENT INFORMATION SECTION #################################

                # NOTE: We hard code the x coordinates of the student headers because they are
                #       individually specific locations (whereas the y location is the same for
                #       both the header and the specific tuple attributes, so having a variable
                #       to apply to the whole row works well).  The win32ui printing TextOut 
                #       function that we use to place textual elements on the page has the 
                #       convention of starting the first character of the text at the specific
                #       x coordinate provided in the function call.  So, because of the way that
                #       the scantron student info section is layed out, this works well for left
                #       justifying the specific student information (the right column), however,
                #       we had to find very specific x locations for the student headers (left
                #       column) and it was different for each string because each string is a 
                #       different length.  That's why we felt hard coding the x coordinates of 
                #       the student headers made more sense than creating variables for them.

                # For the group_id section, we only print this line if the student is actually
                # affiliated with a group.  If the student isn't affiliated with a group he/she
                # will have a group_id value of an empty string or a string value of 0
                if group_id != "":
                    # Print the first line:  GROUP ID : specific student group if applicable 
                    # Select BOLD font for the header and then print it to the page
                    dc.SelectObject(bold_font)
                    dc.TextOut(345, GROUP_NUM, group_id_header)
                    # Now select normal font for the corresponding student group id and print it to the page
                    dc.SelectObject(normal_font)
                    dc.TextOut(THIS_STUDENT_DATA, GROUP_NUM, group_id)
                
                # Print the second line:  STUDENT ID : specific student id num 
                # Select BOLD font for the header and then print it to the page
                dc.SelectObject(bold_font)
                dc.TextOut(280, STUDENT_NUM, student_id_header)
                # Now select normal font for the corresponding student id and print it to the page
                dc.SelectObject(normal_font)
                dc.TextOut(THIS_STUDENT_DATA, STUDENT_NUM, student_id)
                
                # Print the third line:  STUDENT : specific student name 
                # Select BOLD font for the header and then print it to the page
                dc.SelectObject(bold_font)
                dc.TextOut(385, STUDENT, student_header)
                # Now select normal font for the corresponding student name and print it to the page
                dc.SelectObject(normal_font)
                dc.TextOut(THIS_STUDENT_DATA, STUDENT, student_name)
                
                # Print the fourth line:  GRADE : specific student grade 
                # Select BOLD font for the header and then print it to the page
                dc.SelectObject(bold_font)
                dc.TextOut(450, STU_GRADE, grade_header)
                # Now select normal font for the corresponding student grade and print it to the page
                dc.SelectObject(normal_font)
                dc.TextOut(THIS_STUDENT_DATA, STU_GRADE, grade)
                
                # Print the fifth line:  DATE PRINTED : today's date 
                # Select BOLD font for the header and then print it to the page
                dc.SelectObject(bold_font)
                dc.TextOut(205, PRINT_DATE, date_header)
                # Now select normal font for today's date and print it to the page
                dc.SelectObject(normal_font)
                dc.TextOut(THIS_STUDENT_DATA, PRINT_DATE, date_printed)
                
                # Print the sixth line:  PARENTS : specific student parent info 
                # Select BOLD font for the header and then print it to the page
                dc.SelectObject(bold_font)
                dc.TextOut(370, PARENTS1, parents_header)
                
                # Now select normal font for the corresponding parent info and print
                # the 3 lines of parent info to the page
                dc.SelectObject(normal_font)
                dc.TextOut(THIS_STUDENT_DATA, PARENTS1, parents_line1)
                dc.TextOut(THIS_STUDENT_DATA, PARENTS2, parents_line2)
                dc.TextOut(THIS_STUDENT_DATA, PARENTS3, parents_line3)
                
                #################################################################################

                ################## NOT TRANSFERRABLE SECTION ####################################

                # Switch back to BOLD font for the not transferrable statement and print it
                dc.SelectObject(bold_font)
                dc.TextOut(NOT_TRANSFER_X, NOT_TRANSFER_Y, not_transferrable)

                #################################################################################
                
                ################## GRADE DOT SECTION ############################################
                
                # What grade is this student in? We use a switch statement to find it:
                def get_grade(grade):
                    if grade == "3":
                        this_grade = GRADE3
                    elif grade == "4":
                        this_grade = GRADE4
                    elif grade == "5":
                        this_grade = GRADE5
                    elif grade == "6":
                        this_grade = GRADE6
                    elif grade == "7":
                        this_grade = GRADE7
                    elif grade == "8":
                        this_grade = GRADE8
                    # I'm giving the grade attribute a default value of 3
                    # This will only ever happen if the database grade attribute
                    # is null or ''.  I'm including this to prevent the desktop
                    # application from crashing in the rare event of this happening
                    # So take NOTE: if the student associated with this test has an 
                    #               incomplete value for his/her grade it will print
                    #               grade 3 on the scantron
                    else:
                        this_grade = GRADE3
                    return this_grade
                # Now that we've captured this student's grade, assign it a corresponding horizontal coordinate
                this_grade = get_grade(grade)        

                # Create a loop that will repeatedly print horizontal lines 1 horizontal
                # coordinate value lower upon each iteration.  This effectively creates
                # a rectangle object that we can print inside of the "grade" circle.
                # Through trial and error, we determined that 18 consecutive horizontal 
                # lines that are 50 vertical coordinate values long generate an appropriately 
                # sized square to fit inside the grade dot circles.  This loop starts at
                # the very top of the circle and upon each iteration draws a 50 unit
                # long line before decrementing to the next lower horizontal unit. 
                # NOTE: I have yet to determine what the cartesian coordinates in this
                #       program represent.  1 unit in the x or y direction may correspond
                #       with pixels, however, I have yet to confirm that.
                # range args: (where iter begins, where iter ends, increment/decrement amt)  
                for iter in range((SQUARE_LEN + this_grade), this_grade, -1):
                    dc.MoveTo(GRADE_COLUMN, iter)
                    dc.LineTo((GRADE_COLUMN + SQUARE_LEN), iter)

                #################################################################################

                ################## OFFICE USE ONLY SECTION ######################################

                # Here we follow the same procedure as we conducted with the grade dot. The 
                # following for loops create 14 little black boxes that will be printed to the
                # "office use only" section of the scantron sheets.  Each for loop corresponds
                # with one of the rows in that section (0 - 13)

                # However, we need to connect each for loop to its corresponding piece of data.
                # The 14 rows have the following configuration:
                #
                #   ROW 0 : 1st digit of test_id
                #   ROW 1 : 2nd digit of test_id
                #   ROW 2 : 3rd digit of test_id
                #   ROW 3 : 4th digit of test_id
                #   ROW 4 : 5th digit of test_id
                #   ROW 5 : 6th digit of test_id
                #   ROW 6 : 1st digit of account_id
                #   ROW 7 : 2nd digit of account_id
                #   ROW 8 : 3rd digit of account_id
                #   ROW 9 : 4th digit of account_id
                #   ROW 10 : 5th digit of account_id
                #   ROW 11 : 6th digit of account_id
                #   ROW 12 : 1st digit of student_id
                #   ROW 13 : 2nd digit of student_id

                # Let's capture these 3 attributes in variables which will allow us to more
                # readily access their individual digits
                this_test_id = str(no_null_attributes[test][1])
                this_account_id = str(no_null_attributes[test][2])
                this_stu_id = str(no_null_attributes[test][3])        

                # Just as we did for the grade dot x coordinate, we will use a switch statement
                # to calculate each row's x coordinate (which column each column needs to print to)
                def get_digit(one_digit):
                    if one_digit == "0":
                        this_digit = COLUMN0_H
                    elif one_digit == "1":
                        this_digit = COLUMN1_H
                    elif one_digit == "2":
                        this_digit = COLUMN2_H
                    elif one_digit == "3":
                        this_digit = COLUMN3_H
                    elif one_digit == "4":
                        this_digit = COLUMN4_H
                    elif one_digit == "5":
                        this_digit = COLUMN5_H
                    elif one_digit == "6":
                        this_digit = COLUMN6_H
                    elif one_digit == "7":
                        this_digit = COLUMN7_H
                    elif one_digit == "8":
                        this_digit = COLUMN8_H
                    elif one_digit == "9":
                        this_digit = COLUMN9_H
                    # For default we're going to assign this digit a value of zero.  This
                    # would only ever happen if there's an erroneous entry within the DB
                    # But take NOTE: if this happens, this digit will get assigned a value
                    #                of 0.  This will prevent the desktop app from crashing
                    #                if this ever happens. 
                    else:
                        this_digit = COLUMN0_H 
                    return this_digit
                
                # To determine the digit of each row, we call the switch statement function get_digit
                
                # TEST ID
                
                # Row 0
                row_zero_x = get_digit(this_test_id[0])
                for line in range((SQUARE_LEN + ROW0_V), ROW0_V, -1): # Row coordinates
                    dc.MoveTo(row_zero_x, line)                       # Column coordinates
                    dc.LineTo((row_zero_x + SQUARE_LEN), line)
                # Row 1
                row_one_x = get_digit(this_test_id[1])
                for line in range((SQUARE_LEN + ROW1_V), ROW1_V, -1): 
                    dc.MoveTo(row_one_x, line)
                    dc.LineTo((row_one_x + SQUARE_LEN), line)
                # Row 2
                row_two_x = get_digit(this_test_id[2])
                for line in range((SQUARE_LEN + ROW2_V), ROW2_V, -1): 
                    dc.MoveTo(row_two_x, line)
                    dc.LineTo((row_two_x + SQUARE_LEN), line)
                # Row 3
                row_three_x = get_digit(this_test_id[3])
                for line in range((SQUARE_LEN + ROW3_V), ROW3_V, -1):
                    dc.MoveTo(row_three_x, line)
                    dc.LineTo((row_three_x + SQUARE_LEN), line)
                # Row 4
                row_four_x = get_digit(this_test_id[4])
                for line in range((SQUARE_LEN + ROW4_V), ROW4_V, -1): 
                    dc.MoveTo(row_four_x, line)
                    dc.LineTo((row_four_x + SQUARE_LEN), line)
                # Row 5
                row_five_x = get_digit(this_test_id[5])
                for line in range((SQUARE_LEN + ROW5_V), ROW5_V, -1):
                    dc.MoveTo(row_five_x, line)
                    dc.LineTo((row_five_x + SQUARE_LEN), line)

                # ACCOUNT ID

                # Within the account_id section, we need to account for the fact
                # that there are still some active account_id's within Hewitt's 
                # system that are 4 and 5 digit accounts.  Per Hewitt's request,
                # we prepend these accounts with 0's so that every box is still
                # marked upon printing.  So, let's determine how many digits are
                # in this account_id number

                # Most account_id's will be 6 digits:
                if (len(this_account_id) == 6):
                    # Row 6
                    row_six_x = get_digit(this_account_id[0])
                    for line in range((SQUARE_LEN + ROW6_V), ROW6_V, -1):
                        dc.MoveTo(row_six_x, line)
                        dc.LineTo((row_six_x + SQUARE_LEN), line)
                    # Row 7
                    row_seven_x = get_digit(this_account_id[1])
                    for line in range((SQUARE_LEN + ROW7_V), ROW7_V, -1):
                        dc.MoveTo(row_seven_x, line)
                        dc.LineTo((row_seven_x + SQUARE_LEN), line)
                    # Row 8
                    row_eight_x = get_digit(this_account_id[2])
                    for line in range((SQUARE_LEN + ROW8_V), ROW8_V, -1):
                        dc.MoveTo(row_eight_x, line)
                        dc.LineTo((row_eight_x + SQUARE_LEN), line)
                    # Row 9
                    row_nine_x = get_digit(this_account_id[3])
                    for line in range((SQUARE_LEN + ROW9_V), ROW9_V, -1):
                        dc.MoveTo(row_nine_x, line)
                        dc.LineTo((row_nine_x + SQUARE_LEN), line)
                    # Row 10
                    row_ten_x = get_digit(this_account_id[4])
                    for line in range((SQUARE_LEN + ROW10_V), ROW10_V, -1):
                        dc.MoveTo(row_ten_x, line)
                        dc.LineTo((row_ten_x + SQUARE_LEN), line)
                    # Row 11
                    row_eleven_x = get_digit(this_account_id[5])
                    for line in range((SQUARE_LEN + ROW11_V), ROW11_V, -1):
                        dc.MoveTo(row_eleven_x, line)
                        dc.LineTo((row_eleven_x + SQUARE_LEN), line)

                # Account for the account_id's that are 5 digits
                elif (len(this_account_id) == 5):
                    # Row 6
                    # Prepend the first column of this account_id with a forced zero value
                    for line in range((SQUARE_LEN + ROW6_V), ROW6_V, -1):
                        dc.MoveTo(COLUMN0_H, line)
                        dc.LineTo((COLUMN0_H + SQUARE_LEN), line)
                    # Row 7
                    row_seven_x = get_digit(this_account_id[0])
                    for line in range((SQUARE_LEN + ROW7_V), ROW7_V, -1):
                        dc.MoveTo(row_seven_x, line)
                        dc.LineTo((row_seven_x + SQUARE_LEN), line)
                    # Row 8
                    row_eight_x = get_digit(this_account_id[1])
                    for line in range((SQUARE_LEN + ROW8_V), ROW8_V, -1):
                        dc.MoveTo(row_eight_x, line)
                        dc.LineTo((row_eight_x + SQUARE_LEN), line)
                    # Row 9
                    row_nine_x = get_digit(this_account_id[2])
                    for line in range((SQUARE_LEN + ROW9_V), ROW9_V, -1):
                        dc.MoveTo(row_nine_x, line)
                        dc.LineTo((row_nine_x + SQUARE_LEN), line)
                    # Row 10
                    row_ten_x = get_digit(this_account_id[3])
                    for line in range((SQUARE_LEN + ROW10_V), ROW10_V, -1):
                        dc.MoveTo(row_ten_x, line)
                        dc.LineTo((row_ten_x + SQUARE_LEN), line)
                    # Row 11
                    row_eleven_x = get_digit(this_account_id[4])
                    for line in range((SQUARE_LEN + ROW11_V), ROW11_V, -1):
                        dc.MoveTo(row_eleven_x, line)
                        dc.LineTo((row_eleven_x + SQUARE_LEN), line)
                
                # Account for the account_id's that are 4 digits
                else:
                    # We will prepend both the 1st and 2nd rows of 
                    # this account_id with forced zero values
                    # Row 6
                    for line in range((SQUARE_LEN + ROW6_V), ROW6_V, -1):
                        dc.MoveTo(COLUMN0_H, line)
                        dc.LineTo((COLUMN0_H + SQUARE_LEN), line)
                    # Row 7
                    for line in range((SQUARE_LEN + ROW7_V), ROW7_V, -1):
                        dc.MoveTo(COLUMN0_H, line)
                        dc.LineTo((COLUMN0_H + SQUARE_LEN), line)
                    # Row 8
                    row_eight_x = get_digit(this_account_id[0])
                    for line in range((SQUARE_LEN + ROW8_V), ROW8_V, -1):
                        dc.MoveTo(row_eight_x, line)
                        dc.LineTo((row_eight_x + SQUARE_LEN), line)
                    # Row 9
                    row_nine_x = get_digit(this_account_id[1])
                    for line in range((SQUARE_LEN + ROW9_V), ROW9_V, -1):
                        dc.MoveTo(row_nine_x, line)
                        dc.LineTo((row_nine_x + SQUARE_LEN), line)
                    # Row 10
                    row_ten_x = get_digit(this_account_id[2])
                    for line in range((SQUARE_LEN + ROW10_V), ROW10_V, -1):
                        dc.MoveTo(row_ten_x, line)
                        dc.LineTo((row_ten_x + SQUARE_LEN), line)
                    # Row 11
                    row_eleven_x = get_digit(this_account_id[3])
                    for line in range((SQUARE_LEN + ROW11_V), ROW11_V, -1):
                        dc.MoveTo(row_eleven_x, line)
                        dc.LineTo((row_eleven_x + SQUARE_LEN), line)

                # STUDENT ID

                # Most students have a 1 digit student_id (because most families have less than 9 kids)
                # So, for all of the 1 digit student numbers, we need to print a zero to the row that
                # corresponds with the first digit of the student id number

                if (len(this_stu_id) < 2):
                    # ROW 12
                    for line in range((SQUARE_LEN + ROW12_V), ROW12_V, -1):
                        dc.MoveTo(COLUMN0_H, line)
                        dc.LineTo((COLUMN0_H + SQUARE_LEN), line)
                    # Row 13
                    row_thirteen_x = get_digit(this_stu_id[0])
                    for line in range((SQUARE_LEN + ROW13_V), ROW13_V, -1):
                        dc.MoveTo(row_thirteen_x, line)
                        dc.LineTo((row_thirteen_x + SQUARE_LEN), line)
                else:
                    # ROW 12
                    row_twelve_x = get_digit(this_stu_id[0])
                    for line in range((SQUARE_LEN + ROW12_V), ROW12_V, -1):
                        dc.MoveTo(row_twelve_x, line)
                        dc.LineTo((row_twelve_x + SQUARE_LEN), line)
                    # ROW 13
                    row_thirteen_x = get_digit(this_stu_id[1])
                    for line in range((SQUARE_LEN + ROW13_V), ROW13_V, -1):
                        dc.MoveTo(row_thirteen_x, line)
                        dc.LineTo((row_thirteen_x + SQUARE_LEN), line)

                #################################################################################

                ##################### PRINT AND END THIS DOCUMENT ###############################
                dc.EndPage()
                dc.EndDoc()
                #################################################################################

            ####### UPDATE THE DATE_PRINTED ATTRIBUTE IN THE TEST_ORDER TABLE TO TODAY ################

            # Build an array to hold the tuples that will be used in the WHERE clause of the
            # SQL query that will update the date_printed attribute
            update_data =[]
            for iter in range(num_tests):
                # account_id - index 2, student_id - index 3, and test_id - index 1
                update_data.append((no_null_attributes[iter][2], no_null_attributes[iter][3], no_null_attributes[iter][1]))
            # Define the SQL query that will update the date_printed attribute for these tests to today

            update_query = """UPDATE test_order_tester 
                            SET date_printed = getdate() 
                            WHERE account_id = ? AND student_id = ? AND test_id = ?;"""

            # The executemany function will execute multiple UPDATE commands within 1 function call
            my_cursor.executemany(update_query, update_data) 
            # HAVE TO COMMIT YOUR DATABASE UPDATES
            hewitt_db.commit()

        ############### ERROR/EXCEPTION HANDLING ########################################################
        
        # It's important that we implement some exception handling for Hewitt's scripts.  When you're
        # working with large scale databases, the chances of errors and discrepancies are very high.
        # Python uses try/exception blocks (as opposed to C's try/catch blocks). Here, we've placed 
        # the entire script within a try block, and if anything goes wrong throughout the execution 
        # of the script, we will send an error message to the desktop application provoking Hewitt to 
        # look into the error.  For this file, since we are reading from a CSV file, parsing the 
        # column data, and importing new test orders into the database, the following is a small 
        # list of potential errors that might arise:
        #       -Internet connectivity issue
        #       -Error connecting to MS SQL database
        #       -Error connecting to the Ricoh printer
        #       -Miscellaneous database inconsistency with specific student tuple/s

        # If something went wrong in the script, the compiler will enter the following sequence to
    # throw a notification to the desktop app to let Hewitt know of the error.
        except:          
            self.exception_thrown = True
        ###################################################################################################

# NOTE: the following is a class instantiation and function call used only for testing 
# my_example = print_tests()
# my_example.print_scantrons()
