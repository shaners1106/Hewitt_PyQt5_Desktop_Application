# Shane Snediker
# Hewitt Research Foundation Desktop Application
# File last updated: 6-22-2021
# This file contains the Python script for Hewitt's small label printing. 

# The following links were very helpful in the construction of this script:
# win32print library documentation
# http://timgolden.me.uk/pywin32-docs/win32print.html
# update default printer:
# https://codereview.stackexchange.com/questions/193488/changing-the-default-printer-in-windows
# Python and printers:
# https://www.blog.pythonlibrary.org/2010/02/14/python-windows-and-printers/

# The flow of this script is as follows:

#       1. Configure database connection
#       2. Extract all pertinent information for each student who has a test order.  We need
#          each student's account_id, test_id and student_id in order to update their outstanding_order
#          attribute from a 0 to 1 at the end of this script
#       3. We then begin a process of separating students who are affiliated with a group from students
#          who are not.  This is important because ALL ORDERS WITHIN A GROUP ARE SENT TO THE SAME 
#          MAILING ADDRESS.  Once we've separated the group students from non-group students, we can 
#          extract mailing information for these tuples.  NOTE: we only need to print 1 little label
#          for each group and 1 little label for each individualized account_id (thus if one family
#          is ordering multiple tests for their children, we will only print 1 little label for that family) 
#       4. Use Python's Windows library win32ui to instantiate a print document 
#       5. Use a loop to configure each label's printing information
#       5. Change the default printer from the Ricoh to the Zebra and begin printing labels
#       7. Update the outstanding_order attribute in the test_order table for these tuples
#          from a value of 0 to a value of 1

# Names of Hewitt's printers:
#   Ricoh scantron printer: 'Ricoh Aficio MP C6000 PCL5c'
#   Zebra little label printer: '\\\\192.168.1.14\\ZDesigner LP 2844-Z'
#   Zebra shipping label printer: '\\\\192.168.1.14\\ZDesigner LP 2844'


# import libraries
from PyQt5.QtCore import QObject # PyQt5 functionality
import sys               # Needed for configuring the database connection
import pyodbc            # Connect to MS SQL database
from configparser import ConfigParser # Separate configuration settings into separate file
import win32print        # Printer functions
import win32ui           # Library for creating print documents
import win32con          # Access logical pixel calculations

class little_labels(QObject):

    # Little Shipping Label Class MEMBER VARIABLES

    # Flag to communicate with the main GUI in the event that an error occurs during the
    # execution of this script
    exception_thrown = False
    # Flag to signify whether or not Hewitt needs to reprint student exams
    need_reprints = False
    # In the event that Hewitt does need to reprint some little labels, we will ask Hewitt
    # to provide the account number of the last little label that printed.  So let's create
    # a variable to hold the user input account number
    acct_num = None
    # Instantiate a variable to hold the location within the student list where printing will
    # begin (for a new print this will be 0, or in other words we'll begin at the beginning
    # of the student list, but for a reprint situation, the reprint index will be 1 past the
    # index of the last student whose test successfully printed)
    reprint_index = 0

    # The following are 2 variables that will be used in both class methods, so we will 
    # declare them as class member variables.  They will be used to interact with the database
    num_labels = None
    no_null_attributes = []
    update_data = []

    # This function is comprised of the full label printing script
    def print_labels(self):
        try:
            ########### SAVE NAMES OF HEWITT PRINTERS IN VARIABLES ###############################

            # Save the names of the printers 
            Ricoh_printer = 'Ricoh Office Printer'
            Zebra_labeler = 'Zebra Label Printer'

            ######################################################################################
            
            ############### CHANGE TO ZEBRA LABEL PRINTER #################################

            default_printer = win32print.GetDefaultPrinter()
            if default_printer != Zebra_labeler:
                win32print.SetDefaultPrinter(Zebra_labeler)
            ##############################################################################
            
            #################### MAKE A DATABASE CONNECTION ##########################################
            
            config = ConfigParser()
            # This will determine if its the .exe or just the script
            if 'dist' in sys.path[0]: config.read('config.ini')
            else: config.read(sys.path[0] + '\\config.ini')
            # Transforms the config infor into a dictionary for the program to use
            config_dict = dict(config.items())

            # Connects to the database
            hewitt_db = pyodbc.connect(
           # NOTE: confidental database credentials omitted for security purposes
            )
            ##########################################################################################

            ######### CAPTURE ALL OF THE STUDENT INFO FOR THE LABELS THAT NEED TO BE PRINTED ##########

            # The cursor function is returns a control structure that enables the 
            # execution of queries against the DB
            my_cursor = hewitt_db.cursor()
            # Pull in all current tests that still need to be printed
            # NOTE: the print_tests VIEW from the Hewitt DB always contains tests that need 
            #       to be printed, so selecting everything from this VIEW is what we want
            
            my_cursor.execute('SELECT * FROM print_labels_tester;')
            # Save the tuple information in a variable that we can iterate over.  The fetchall()
            # function returns a tuple of tuples (the internal tuples consisting of individual
            # test order attributes)
            current_labels = my_cursor.fetchall()
            
            # How many labels will we be printing on this job?
            self.num_labels = len(current_labels)
            
            # And how many attributes does each label tuple hold?
            num_attributes = len(current_labels[0])
            ##############################################################################################

            ########### CONVERT NULL VALUES TO EMPTY STRINGS #############################################

            # Python hates null values (or what they term 'None types')
            # We begin interacting with the data by converting all nulls to empty strings, which
            # is something that Python can deal with better
            
            # For each individual tuple:
            for ind_tup in range(self.num_labels):
                # check it for nulls and if found, convert to '':
                # However modifying an existing tuple is not permitted in Python, so we'll have to
                # do some extra work to create new tuples.  This is necessary because if this script
                # ever assigns an attribute that has a null value, it will crash the script
                # Track the attributes within this tuple that hold null values
                null_indices =[]
                for attribute in range(num_attributes):
                    # Does this index hold a null value?
                    if current_labels[ind_tup][attribute] is None:
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
                        inner_tuple.append(current_labels[ind_tup][index])
                # Finally, add this inner array as a newly created null-less tuple into the new outer array    
                self.no_null_attributes.append(inner_tuple)
            
                # Now we have a Python array corresponding to the data tuple with null values replaced by empty strings
            #########################################################################################################################################

            ## BUILD THE UPDATE_DATA ARRAY THAT WILL HOLD THE CREDENTIALS OF EACH STUDENT WHO WILL NEED THEIR OUTSTANDING_ORDER ATTRIBUTE UPDATED ###

            # Each student test order will need to have their outstanding_order attribute changed from a 
            # 0 to a 1 at the end of this script, so before we start separating these tuples into groups
            # vs. non-group arrays we should save the account_id's, test_id's and student_id's for when
            # we modify the database at the end.  
            for iter in range(self.num_labels):
                self.update_data.append((self.no_null_attributes[iter][1], self.no_null_attributes[iter][2], self.no_null_attributes[iter][3]))
            ########################################################################################################################################

            ############# SEPARATE GROUP STUDENT TUPLES FROM INDIVIDUAL STUDENT TUPLES #############################################################

            # Declare an array to hold the grouped student tuples
            group_tuples = []
            # Declare an index that will correspond to where the students who are affiliated
            # with a group begin within the 2 dimensional array
            idx_grp_begin = 0
            # Flag we can use to capture the index of the 1st group student (we know that the
            # SQL will order the tuples by group_id in ascending order, so the 2 dimensional
            # array will begin with NULL group_id's corresponding with students who aren't
            # affiliated with a group and will be followed by non-NULL group_id's corresponding
            # to the students who are a part of a group order.  We need that 1st index so we'll
            # set this flag to True once we've found it)
            begin = False
            # Iterate through our array of tuples and add any non-NULL group_id tuples to the 
            # newly declared group_tuples array. 
            for i in range(len(self.no_null_attributes)):
                # Remember that we've switched NULL values to empty strings to make Python happy
                if self.no_null_attributes[i][0] != "":
                    # If this is the first instance of a non-NULL group_id in this array
                    if not(begin):
                        # Then save the index and set the flag
                        idx_grp_begin = i
                        begin = True
                    # Add this tuple to the group_tuples array
                    group_tuples.append(self.no_null_attributes[i])
            # Now delete all of the group tuples from the original 2 dimensional array so that 
            # it will be left with only the non-group tuples and we will have 2 separate arrays
            # that we can work with and manipulate for printing little labels
            del self.no_null_attributes[idx_grp_begin:self.num_labels]
            ########################################################################################################################################

            ############ IDENTIFY UNIQUE GROUP_ID'S/ACCOUNT_ID'S AND GET RID OF DUPLICATES ######################################################### 

            # UNIQUE GROUP_ID'S

            # Let's initialize a list with which we can hold unique group_id's (duplicates get left out)
            grp_id_holder = []
            # Start a loop to capture every unique group_id in the group_tuples array
            for unique_grp in range(len(group_tuples)):
                if (group_tuples[unique_grp][0] not in grp_id_holder):
                    grp_id_holder.append(group_tuples[unique_grp][0])
            
            # UNIQUE INDIVIDUAL ACCOUNT_ID'S

            # Initialize a list with which we can hold unique account_id's (duplicates are left out)
            acct_id_holder = []
            # Let's also save the indices where duplicate account_id's reside so that we can delete them later
            duplicate_idxs = []
            # Start a loop to capture every unique account_id in the array that is currently 
            # holding only individual student test order tuples
            for unique_acct in range(len(self.no_null_attributes)):
                if (self.no_null_attributes[unique_acct][1] not in acct_id_holder):
                    acct_id_holder.append(self.no_null_attributes[unique_acct][1])
                else:
                    duplicate_idxs.append(unique_acct)
            ########################################################################################################################################

            ######### PREPARE UNIQUE GROUP TUPLES AND INDIVIDUAL TUPLES TO BE PRINTED ##############################################################

            # INDIVIDUAL TUPLES

            # New array to save unique tuples
            ind_saver = []
            # Trim down the individual account_id array by saving only unique tuples
            for iter in range(len(self.no_null_attributes)):
                if iter not in duplicate_idxs:
                    ind_saver.append(self.no_null_attributes[iter])
            
            # Clean up
            self.no_null_attributes.clear()

            # GROUP TUPLES

            # For the group tuples, we are going to have to query the database for each of the unique 
            # group_id's to capture the corresponding mailing_list information (NOTE: a student's group_id
            # corresponds to the account_id of the group leader.  We need the group leader's mailing info)
            # Clear out the group_tuples list so that it will be ready to hold the group leaders' info
            group_tuples.clear()
            # Start loop to capture group leaders' info
            for query in range(len(grp_id_holder)):
                my_cursor.execute('''SELECT m.account_id, m.customer1_first_name, m.customer1_last_name, m.address1, m.city1, m.state1, m.zipcode1, m.ship_first_name1, m.ship_last_name1, m.ship_address, m.ship_city, m.ship_state, m.ship_zipcode
                from mailing_list_tester m
                WHERE m.account_id = ''' + str(grp_id_holder[query]) + ''';''')
                # Now pull this query and append it to our group_tuples list
                this_tuple = my_cursor.fetchall()
                group_tuples.append(this_tuple)

            ########### CONVERT NULL VALUES TO EMPTY STRINGS #############################################

            # Python hates null values (or what they term 'None types')
            # We begin interacting with the data by converting all nulls to empty strings, which
            # is something that Python can deal with better
            
            # For each individual tuple:
            for ind_tup in range(len(group_tuples)):
                # check it for nulls and if found, convert to '':
                # However modifying an existing tuple is not permitted in Python, so we'll have to
                # do some extra work to create new tuples.  This is necessary because if this script
                # ever assigns an attribute that has a null value, it will crash the script
                # Track the attributes within this tuple that hold null values
                # NOTE: Because we didn't know the group_id numbers (and subsequently the group leaders'
                #       account_id numbers) until runtime, we had to execute the above SQL query within 
                #       a loop.  This resulted in an array of arrays of tuples.  This is different than
                #       our original SQL query which retrieved all of the tuples in 1 query and put them
                #       into 1 array of tuples.  Thus, the following code calibrates our data structure
                #       back to being 1 array holding arrays corresponding to indivdual tuples.  In other
                #       words, the "CONVERT NULL VALUES" section above after the original SQL query took
                #       the tuple attributes and placed them as arrays within an array and the following
                #       "CONVERT NULL VALUES" section will take the tuple attributes out of the tuple and
                #       store them as arrays in a single array.  The first version was working with a 2 
                #       dimensional array, the following version will be working with a 3 dimensional array.

                null_indices =[]
                for attribute in range(len(group_tuples[0][0])):
                    # Does this index hold a null value?
                    if group_tuples[ind_tup][0][attribute] is None:
                        # Then save this index number
                        null_indices.append(attribute)
                # Instantiate an inner array to hold this tuple's attributes
                inner_tuple = []
                # For each index, if it is an index number that held a null value, append an empty
                # string into the new array, otherwise append the original value
                for index in range(len(group_tuples[0][0])):
                    if index in null_indices:
                        inner_tuple.append("")
                    else:
                        inner_tuple.append(group_tuples[ind_tup][0][index])
                # Finally, add this inner array as a newly created null-less tuple into the new outer array    
                self.no_null_attributes.append(inner_tuple)
            
            # Now we have a Python array corresponding to the data tuple with null values replaced by empty strings
            #########################################################################################################################################

            ######### CALIBRATE GROUP_TUPLES ARRAY AND COMBINE THE 2 ARRAYS INTO 1 ARRAY ###########################################################

            # It's important for us to combine the group_tuples and the ind_saver tuples into 1 array.  This is because we are providing Hewitt the 
            # ability to reprint little labels at a specified location within the print queue.  This means that we really need the printing to be 
            # done within 1 loop, rather than 2 separate loops.  All that will be required to combine the arrays is to fill the group_tuples array
            # with filler indices so that indices that correspond with data that will be used to print onto the labels will be consistent for both
            # kinds of tuples.  In other words, we need the 'customer1_first_name' attribute within the group_tuples arrays to be located at the same
            # index within the array as the index within the ind_saver array holding the 'customer1_first_name' attribute.

            # Therefore, the only indices we need to fill with filler values are indices 0, 1, 2, and 3.  This is because within the ind_saver array
            # those indices correspond to attributes that are not printing data attributes.  The rest of the indices will line up if we fill the 
            # first four indices with filler values.  However, index 1 needs to be the account_id (for reprinting functionality); the other indices
            # can be dummy values.  So, I'll filly the dummy indices with empty strings:
            for iter in range(len(self.no_null_attributes)):
                self.no_null_attributes[iter].insert(0, '')
                self.no_null_attributes[iter].insert(2, '')
                self.no_null_attributes[iter].insert(2, '')
            
            # Now we have arrays whose indices are consistent and aligned.  Let's combine them into 1 array:
            for iter in range(len(ind_saver)):
                self.no_null_attributes.append(ind_saver[iter])
            
            # Clean up
            del ind_saver
            del group_tuples
            ########################################################################################################################################
            
            ################ PREPARE WIN32UI PRINT DOCUMENT ########################################################################################

            #   NOTE: The win32ui TextOut function by which we send textual elements to the page only
            #   accepts string values within the function call.  Therefore, even the attributes
            #   that are ints must be stored as string values

            # We now know how many labels we will be printing.  It will be the amount of arrays within
            # the no_null_attributes array.  
            self.num_labels = len(self.no_null_attributes)
            # Each array within no_null_attributes corresponds to a tuple
            # that represents the mailing location of testing exams.  We will instantiate a win32ui print
            # document for each label within a loop.  The following are the attributes of each tuple and
            # their corresponding indices:
            #   
            #   [0] : group_id      NOTE: this attribute is only needed for helping make sure the labels
            #                             get printed in an intuitive order; it won't be printed to the labels
            #   [1] : account_id    NOTE: this is an int data type and will need to be converted to a string 
            #                             in order to print to the win32ui document
            #   [2] : student_id   NOTE: this is only here because we use it in the query at the end of
            #                            this script that changes the outstanding_order value to a 1 value
            #   [3] : test_id      NOTE: Same as student_id (only here for updating DB data)
            #   [4] : customer1_first_name
            #   [5] : customer1_last_name
            #   [6] : address1
            #   [7] : city1
            #   [8] : state1 
            #   [9] : zipcode1
            #   [10] : ship_first_name1
            #   [11] : ship_last_name1
            #   [12] : ship_address
            #   [13] : ship_city
            #   [14] : ship_state
            #   [15] : ship_zipcode
            #   [16] : student_first_name NOTE: this attribute is only needed for helping make sure that the labels get printed in an intuitive, organized order
            #   [17] : student_last_name NOTE: this attribute is only needed for helping make sure that the labels get printed in an intuitive, organized order

            # Special function used to adjust win32ui print font size
            # This function uses a mathematical calculation to convert the desired standard
            # font size to the scale of the way that win32ui prints to the page.
            # args: dc : the win32ui object that the font will be used on
            #       PointSize : the desired standard font size
            # We found this function at the following URL:
            # https://stackoverflow.com/questions/48549555/how-to-set-font-type-and-size-for-printing-using-windows-gdi
            def getfontsize(dc, PointSize):
                inch_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)
                return int(-(PointSize * inch_y) / 72)
            #########################################################################################

            ############### PROVIDE HEWITT OPPORTUNITY TO REPRINT LABELS ############################
            
            if(self.need_reprints and self.acct_num is not None):
                ############ FIND RESTART INDEX #####################################################
                # Start a loop to find which tuple Hewitt's print job ended on (they will have just entered
                # the account_id into the desktop application)
                for iter in range(self.num_labels):
                    # Find the tuple with the same account number as the one entered by Hewitt
                    if(self.no_null_attributes[iter][1] == self.acct_num):
                        # If for some reason the account number that Hewitt inputted is the
                        # account number of the last label in the queue, set the reprint 
                        # index number to the very end of the queue and don't print anything
                        # because otherwise iterating the index value by 1 will cause an index
                        # out of range error
                        if (iter == self.num_labels - 1):
                            self.reprint_index = iter
                            # Break out of the for loop
                            break
                        # Otherwise, we don't need to reprint the label corresponding to this account_id,
                        # so go to the next one
                        else:
                            self.reprint_index = iter + 1
                            # Break out of the for loop
                            break
            #########################################################################################
            
            ############ BEGIN LOOP TO PRINT LABELS #######################################
            for label in range(self.reprint_index, self.num_labels):

                ######### DETERMINE IF THIS TUPLE HAS A UNIQUE SHIPPING NAME AND ADDRESS ###########

                # If the ship_first_name1 attribute is not null, then we need to use the shipping names
                use_ship_name = False
                if self.no_null_attributes[label][10] != "":
                    use_ship_name = True
                
                # If the ship_address attribute is not null, then we need to use the shipping info
                use_ship_info = False
                if self.no_null_attributes[label][12] != "":
                    use_ship_info = True
                ####################################################################################

                ########### VARIABLE DECLARATIONS ####################################################

                # NOTE: the win32ui TextOut function can only send string data types to the page
                # This tuple's account_id is held in index 1 of the data structure holding label info
                account_id = str(self.no_null_attributes[label][1])
                # Build this label's parent name line
                # NOTE: if there's a unique shipping name, we'll use that
                if use_ship_name:
                    parent_name = self.no_null_attributes[label][10] + ' ' + self.no_null_attributes[label][11]
                # If there isn't a specified shipping name, we use customer/guardian 1's name
                else:
                    parent_name = self.no_null_attributes[label][4] + ' ' + self.no_null_attributes[label][5]
                # Build this labels address info
                # NOTE: if there's a unique shipping address, we'll use that
                if use_ship_info:
                    street_address = self.no_null_attributes[label][12]
                    city_state = self.no_null_attributes[label][13] + ' ' + self.no_null_attributes[label][14] + ', ' + self.no_null_attributes[label][15]
                # If there isn't a specified shipping address, we use customer/guardian's
                else:
                    street_address = self.no_null_attributes[label][6]
                    city_state = self.no_null_attributes[label][7] + ' ' + self.no_null_attributes[label][8] + ', ' + self.no_null_attributes[label][9]
                #######################################################################################
                
                ################ BEGIN PRINT DOCUMENT HERE ####################################
                # Instantiate a win32ui object that we can use to generate printable objects
                dc = win32ui.CreateDC()
                # Connect the win32ui object to the printer
                dc.CreatePrinterDC()
                # Begin a new document and initialize the page
                dc.StartDoc('Print Labels Document')
                dc.StartPage()

                # Now we define the necessary fonts that will be used within the template

                ################ DEFINE FONTS #################################################

                # Make the text size 11
                fontsize = getfontsize(dc, 12)
                # Create a bold font for the student info headers (left column of the student info section) as well as the non-transferrable statement
                fontdata_headers = { 'name':'Arial', 'height':fontsize, 'italic':False, 'weight':win32con.FW_BOLD}
                # Create a win32ui bold font object and save it in a variable that we can utilize throughout this document
                bold_font = win32ui.CreateFont(fontdata_headers)
                # Give this document the font characteristics that we just defined
                dc.SelectObject(bold_font)

                ############### PLACE TEXT ELEMENTS WITHIN DOCUMENT ###########################

                # Per Hewitt's request, we left align the little label content (keep it at 0 in the x direction)
                # Place account id in the upper left corner of the label
                dc.TextOut(0, 20, account_id)
                # Leave a space in between account id and the rest of the information
                # Next piece of information is the parent name
                dc.TextOut(0, 120, parent_name)
                # Place the street address
                dc.TextOut(0, 170, street_address)
                # Finally place the city, state and zipcode
                dc.TextOut(0, 220, city_state)

                ############### PRINT LABELS ##################################################
                
                dc.EndPage()

                ############### END DOCUMENT ##################################################

                dc.EndDoc()
                ###############################################################################
            
        ############### GIVE DEFAULT PRINTER CONTROL BACK TO RICOH ##############################

        win32print.SetDefaultPrinter(Ricoh_printer)
        #########################################################################################

        ############## ERROR/EXCEPTION HANDLING #######################################################

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
        #       -Error connecting to the Zebra printer
        #       -Miscellaneous database inconsistency with specific student tuple/s

        # If something went wrong in the script, the compiler will enter the following sequence to
        # throw a notification to the desktop app to let Hewitt know of the error.
        except:
            self.exception_thrown = True

    ######## UPDATE DATABASE OUTSTANDING_ORDER ATTRIBUTE FOR EACH STUDENT TEST ORDER #################

    # We create another class method to update the outstanding_order attribute for the each tuple
    # having a test printed during this script.  This breaks up the flow of the label printing
    # so that the outstanding_order attributes aren't automatically getting updated right after the
    # labels are printed and we can communicate better with Hewitt through the GUI
    # NOTE: the database view used to change the outstanding_order attribute is the same one used
    #        
    def update_outstanding(self):
        try:

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

            # # The cursor function is returns a control structure that enables the 
            # # execution of queries against the DB
            my_cursor = hewitt_db.cursor()

            # #### UPDATE OUTSTANDING_ORDER COLUMN OF TEST_ORDER TABLE FOR THESE TEST ORDERS ##########
                
            # Now that Hewitt has confirmed that they've printed all of the labels that they
            # need, we can update the outstanding_order value for each of these tuples from
            # a 0 to 1.  NOTE: within the test_order table an outstanding_order value of 0 
            # corresponds to all tests that haven't been printed, a value of 1 corresponds to
            # all test_orders that have been printed and shipped but not yet returned. 
            # Our SQL test to make sure that we're only updating the specific test tuples of
            # the test orders that we were just printed in the print_tests script consists of 
            # querying the test_order table and updating the outstanding_order attribute to 1 
            # for all tuples with matching account_id's, student_id's and test_id's as the 
            # ones from this test order print job
           

            # Now we build the UPDATE query that we will use to update the outstanding_order attributes
            # NOTE: Here we use string formatting syntax (%s).  The %s's get populated with the
            #       corresponding tuple data provided in the executemany function call.  So, for
            #       example we've created an array of 3 value tuples (account_id, student_id, test_id)
            #       When we call the execute many function it will go through each tuple in the array
            #       and populate the WHERE clause parameters with the corresponding attribute data  
            # NOTE: The "my_cursor.execute('SELECT * FROM Hewitt_DB.print_labels;')" function call
            #       from early in this script accessed the VIEW 'print_labels'.  Database views are
            #       kind of like virtual copies of data and cannot be directly modified.  Therefore,
            #       when we are updating database tuples, we have to access the actual tables where the
            #       original data resides

            update_query = """UPDATE test_order_tester 
                    SET outstanding_order = 1 
                    WHERE account_id = ? AND student_id = ? AND test_id = ?;"""
            # The executemany function will execute multiple UPDATE commands within 1 function call
            my_cursor.executemany(update_query, self.update_data)
            # Spent a couple hours trying to figure out why I could update rows in the test_order
            # database in MySQL Workbench but not from this Python file and it turns out you 
            # HAVE TO COMMIT YOUR DATABASE UPDATES!!
            hewitt_db.commit()
            ###########################################################################################

        except:
            self.exception_thrown = True

# NOTE: the following is a class instantiation and function call used only for testing 
# my_example = little_labels()
# my_example.print_labels()