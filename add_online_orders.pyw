# Shane Snediker
# Hewitt Research Foundation Desktop Application
# File last updated: 6-22-2021

# This file contains the Python script used to import the CSV files exported from
# WooCommerce when a customer orders tests

# The flow of this script is as follows:

#       1. Configure database connection
#       2. Check database for any outstanding test orders that have exceeded Hewitt's
#          90 day test deadline and export a CSV file with the information from these tuples
#       2. Extract WooCommerce CSV file contents and organize contents into a 2 dimensional 
#          array (an array of tuples)
#       3. Define the SQL query necessary to INSERT these new test_orders into the database
#       4. Update Hewitt's database with the test orders

from PyQt5.QtCore import QObject # PyQt5 functionality
import sys   # Needed for configuring the database connection
import pyodbc   # Access the MS SQL database
from configparser import ConfigParser # Separate configuration settings into separate file
import csv   # Process csv files
from datetime import date, datetime, timedelta # Import functionality for accessing today's date


# This class comprises the full CSV WooCommerce script
class wooCommerce(QObject):

    # CLASS MEMBER VARIABLES

    # Flag to indicate an error/exception was thrown
    exception_thrown = False

    # This Boolean will be set when this script finds tests that have exceeded the 90 day Hewitt deadline
    deadline_exceeded = False
    # Amount of tests exceeding the deadline found during the running of this script
    num_unreturned = 0

    # Create a function to query the database, find any test_orders that are past due,
    # and export a CSV file to Hewitt with pertinent file information
    def find_deadline_exceeded(self):
        try:

            ############ BEGIN BY CONNECTING THIS SCRIPT TO HEWITT'S DATABASE ##########################
            
            config = ConfigParser()
            # This will determine if its the .exe or just the script
            if 'dist' in sys.path[0]: config.read('config.ini')
            else: config.read(sys.path[0] + '\\config.ini')
            # Transforms the config infor into a dictionary for the program to use
            config_dict = dict(config.items())

            # Connect to the database
            hewitt_db = pyodbc.connect(
            # NOTE : all database credentials removed for security purposes
            )
            ##########################################################################################

            ####### CAPTURE ALL OF THE STUDENT INFO FOR PAST DUE EXAMS ###############################

            # The cursor function is returns a control structure that enables the 
            # execution of queries against the DB
            my_cursor = hewitt_db.cursor()
            # Pull in all test orders that have exceeded Hewitt's 90 day deadline
            # NOTE: the deadline_exceeded VIEW within Hewitt's DB was constructed specifically
            #       for this script so selecting everything from this VIEW is what we want
            
            my_cursor.execute('SELECT * FROM deadline_exceeded_tester') 
            # Save the tuple information in a variable that we can iterate over.  The fetchall()
            # function returns a tuple of tuples (the internal tuples consisting of individual
            # test order attributes)
            unreturned_tests = my_cursor.fetchall()
            
            # Now that we have a data structure holding all of our student data, we can begin
            # to work with the data. First, how many tests have exceeded the 90 day deadline?
            self.num_unreturned = len(unreturned_tests)
            
            # Are there any past due exams?
            if self.num_unreturned > 0:
                # Set the flag so that we can communicate with the GUI that we found past due exams
                self.deadline_exceeded = True
                
                # How many columns will each CSV tuple have?
                num_attributes = len(unreturned_tests[0])

                ##### EXPORT A CSV WITH PERTINENT STUDENT INFO SO HEWITT CAN CONTACT THESE TUPLES IF NECESSARY ###

                   ### FORMAT THE CSV FILEPATH ###

                # TODO: REFORMAT THE FILE PATH

                # TEMP FILE PATH SO I COULD TEST ON MY OWN MACHINE
                Hewitt_csv_export_path = 'C:\\Users\\Snediker\\Documents\\Whitworth\\Last call\\Software Engineering\\past_due_tests.csv'
                
                # Format the file path here and save in a variable

                ########### CONVERT NULL VALUES TO EMPTY STRINGS ###################################

                # Python hates null values (or what they term 'None types')
                # We begin interacting with the data by converting all nulls to empty strings, which
                # is something that Python can deal with
                # We need to use an array in order to rebuild the tuple
                no_null_attributes = []
                # For each individual tuple:
                for ind_tup in range(self.num_unreturned):
                    # check it for nulls and if found, convert to '':
                    # However modifying an existing tuple is not permitted in Python, so we'll have to
                    # do some extra work to create new tuples.  This is necessary because if this script
                    # ever assigns an attribute that has a null value, it will crash the script
                    # Track the attributes within this tuple that hold null values
                    null_indices =[]
                    for attribute in range(num_attributes):
                        # Does this index hold a null value?
                        if unreturned_tests[ind_tup][attribute] is None:
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
                            inner_tuple.append(unreturned_tests[ind_tup][index])
                    # Finally, add this inner array as a newly created null-less tuple into the new outer array    
                    no_null_attributes.append(inner_tuple)
                    
                    # Now we have a Python array corresponding to the data tuple with null values replaced by empty strings
                ###############################################################################################################

                ############ BEGIN LOOP TO BUILD PAST DUE TEST TUPLES #########################################################
                
                # Now it's time to export the CSV file : instantiate a Python list to hold each
                # row of the CSV file:
                past_due_holder = []
                
                # Each loop processes 1 past due test tuple and places each attribute within a tuple
                for row in range(self.num_unreturned):
                    # NOTE: the CSV file that we export will have 7 columns in the following configuration:

                    #  [0]: student_first_name  [1]: student_last_name  [2]: date_printed  [3]:  account_id  [4]: student_id  [5]: test_id  [6]:  group_id

                    # Build a tuple with the exact order you want these attributes to appear in the CSV file
                    csv_tuple = (no_null_attributes[row][0], no_null_attributes[row][1], no_null_attributes[row][2], no_null_attributes[row][3], no_null_attributes[row][4], no_null_attributes[row][5], no_null_attributes[row][6])
                    # Place thise tuple within the row holder array
                    past_due_holder.append(csv_tuple)
                ################################################################################################################
                
                ############### SEND THIS CSV FILE TO THE DESIGNATED CSV DIRECTORY #############################################

                # The following code appends each of these rows to an existing CSV file within Hewitt's system.
                # NOTE: 'a+' is the flag for appending
                with open(Hewitt_csv_export_path, 'a+', newline='') as out:
                    # Create a csv file writer object that we can use to create our CSV file
                    file_writer = csv.writer(out)
                    # Iterate through each shipping label, writing the corresponding attributes into each column
                    for row in past_due_holder:
                        file_writer.writerow(row)
                ###############################################################################################################

                ##### UPDATE OUTSTANDING_ORDER COLUMN OF TEST_ORDER TABLE FOR THESE PAST DUE EXAMS ############################
            
                # NOTE: within the test_order table an outstanding_order value of 0 
                # corresponds to all tests that haven't been printed, a value of 1 corresponds to
                # all test_orders that have been printed but not shipped yet, a value of 2 corresponds
                # to all test_orders that have been shipped but not yet returned, a value of 3
                # corresponds to all test_orders that have been returned, and a value of 4 means that
                # the customer/student didn't return the test_order in time (there's a 90 day deadline).
                #  
                # First we create a Python array to hold the data we will need to make sure we only
                # change the outstanding_order attribute for the tests that we just determined are 
                # past due.  Fortunately, the deadline_exceeded DB VIEW captures the relevant 
                # information-we just need to make sure to be careful to utilize the correct attributes.

                # Create an array that will hold the tuples needed to query the database
                update_data = []
                # For each unreturned test, create a query tuple
                for iter in range(self.num_unreturned):
                    # account_id - index 3, student_id - index 4, and test_id - index 5
                    update_data.append((no_null_attributes[iter][3], no_null_attributes[iter][4], no_null_attributes[iter][5]))

                # Now we build the UPDATE query that we will use to update the outstanding_order attributes
                # NOTE: Here we use string formatting syntax (?).  The ?s's get populated with the
                #       corresponding tuple data provided in the executemany function call.  So, for
                #       example we've created an array of 3 value tuples (account_id, student_id, test_id)
                #       When we call the execute many function it will go through each tuple in the array
                #       and populate the WHERE clause parameters with the corresponding attribute data  
                # NOTE: The "my_cursor.execute('SELECT * FROM deadline_exceeded;')" function call
                #       from early in this script accessed the VIEW 'deadline_exceeded'.  Database views are
                #       kind of like virtual copies of data and cannot be directly modified.  Therefore,
                #       when we are updating database tuples, we have to access the actual tables where the
                #       original data resides
                
                # Define the SQL query that will update the date_printed attribute for these tests to today

                update_query = """UPDATE test_order_tester 
                                SET outstanding_order = 4 
                                WHERE account_id = ? AND student_id = ? AND test_id = ?;"""
                # The executemany function will execute multiple UPDATE commands within 1 function call
                my_cursor.executemany(update_query, update_data)
                # HAVE TO COMMIT YOUR DATABASE UPDATES
                hewitt_db.commit()

        except:
            self.exception_thrown = True

    # This function will open the WooCommerce CSV file of new test_orders, pull and parse all of the column data
    # and add the test orders into Hewitt's test_order table in their database
    def add_online_orders(self):
        try:

            ############ BEGIN BY CONNECTING THIS SCRIPT TO HEWITT'S DATABASE ##########################

            config = ConfigParser()
            # This will determine if its the .exe or just the script
            if 'dist' in sys.path[0]: config.read('config.ini')
            else: config.read(sys.path[0] + '\\config.ini')
            # Transforms the config infor into a dictionary for the program to use
            config_dict = dict(config.items())

            # Connects to the database
            hewitt_db = pyodbc.connect(
            # NOTE : all database credentials removed for security purposes
            )
            ##########################################################################################

            ########## CONFIGURE THE FILEPATH WHERE THE WOOCOMMERCE CSV FILE IS LOCATED ##############

            # The following is the official filepath provided by Hewitt that maps to the location 
            # on Hewitt's system where the WooCommerce CSV file will always be located:

            # # temporary open statement to so that I could test functionality my own local machine
            temp_woo = 'C:\\Users\\Snediker\\Documents\\Whitworth\\Last call\\Software Engineering\\Hewitt\\New Orders\\May21.csv'
            woo_order = open(temp_woo, 'rt')
            read_this_file = csv.reader(woo_order)

            # Open CSV

            # Now that we're in, the power of extraction can take place.  We construct a 2-dimensional
            # array.  It will be an array holding each row of the csv file.  Each row is an individual
            # student test order and the data needs to be loaded into Hewitt's database
            test_order_holder = []
            for row in read_this_file:
                test_order_holder.append(row)
            ############################################################################################

            ########## CONVERT CSV DATA INTO DB FORMAT #################################################

            # Our one constant value is the amount of columns in this file.  It will ALWAYS contain
            # 7 columns
            NUM_COLUMNS = 7

            # The following is a breakdown of the columns provided in Hewitt's CSV WooCommerce files:

            #   Group ID    ID      Student First Name      Student Last Name       Grade       Date        Date Ordered
            # 
            #   -GROUP ID will be the 6 digit account_id of the group leader plus a dash followed
            #       by the quantity of exams in that group order.  Example: 123456-1
            #   -ID consists of the 6 digit parent/guardian account_id and the 2 digit student_id
            #       in the following format: 165165-2
            #   -Student First Name is self explanatory
            #   -Student Last Name is self explanatory
            #   -Grade of the student taking the exam (this value will be the dot that gets marked on their scantron)
            #   -Date : this field requires an important distinction.  This date field represents the 
            #           date that the customer is requesting test materials/curriculum.  If this field 
            #           is left blank, we have been instructed to assign this test order a print_on_date of
            #           the current date.  If this field is not blank, that means the customer is requesting
            #           a delayed shipping date.  Therefore, we have been instructed to assign these test orders
            #           a print_on_date of 18 days before the date provided in this field.
            #   -Date ordered is provided so that Hewitt can maintain an accurate record of when the test was ordered
            #   -Group qty is the amount of test orders for group orders.  This field will be blank for 
            #         orders that are not affiliated with a group

            # Therefore, each internal array will hold 7 indices and the index breakdown is as follows:

            #       [0] : Group ID
            #       [1] : ID
            #       [2] : Student First Name
            #       [3] : Student Last Name
            #       [4] : Grade
            #       [5] : Date
            #       [6] : Date ordered
            
            # NOTE: the first internal array can be ignored because it contains the row of column
            #       headers.  When we begin building database tuples, we need to skip over index 0.
            # 
            # We will need to use today's date in this process
            today = date.today().strftime("%Y-%m-%d") 

            # Our strategy here is to conduct the entirety of this data organization process within
            # 1 for loop.  In this way, with each iteration of the loop, we can access and distribute
            # the necessary attributes into a tuple that can later be uploaded into the test_results
            # table of the database
            # 
            # How many rows does the CSV file have?
            num_tuples = len(test_order_holder)

            # Declare an array to hold the resultant tuples that will then get imported into the DB
            tuple_holder = []

            ########### GET RID OF NULL VALUES ######################################################

            # Python hates null values (or what they term 'None types')
            # We begin interacting with the data by converting all nulls to empty strings, which
            # is something that Python can deal with
            # We need to use an array in order to rebuild the tuple
            no_null_attributes = []
            # For each individual tuple:
            for ind_tup in range(1, num_tuples):
                # check it for nulls and if found, convert to '':
                # However modifying an existing tuple is not permitted in Python, so we'll have to
                # do some extra work to create new tuples.  This is necessary because if this script
                # ever assigns an attribute that has a null value, it will crash the script
                # Track the attributes within this tuple that hold null values
                null_indices =[]
                for attribute in range(NUM_COLUMNS):
                    # Does this index hold a null value?
                    if test_order_holder[ind_tup][attribute] is None:
                        # Then save this index number
                        null_indices.append(attribute)
                # Instantiate an inner array to hold this tuple's attributes
                inner_tuple = []
                # For each index, if it is an index number that held a null value, append an empty
                # string into the new array, otherwise append the original value
                for index in range(NUM_COLUMNS):
                    if index in null_indices:
                        inner_tuple.append("")
                    else:
                        inner_tuple.append(test_order_holder[ind_tup][index])
                # Finally, add this inner array as a newly created null-less tuple into the new outer array    
                no_null_attributes.append(inner_tuple)
                
                # Now we have a Python array corresponding to the data tuple with null values replaced by empty strings
            
            ########### BEGIN THE LOOP THAT WILL COMPRISE THE DATA PARSING OF THE TUPLES #####################################
            
            # Start the loop that will process each row of the CSV file and save the result into a tuple
            for row in range(num_tuples - 1): # Skip the 0th row because it's the column headers
                
                # Declare an array to hold this row's parsed and organized attributes
                this_tuple_data = []

                # It's important at this point to acknowledge the specific order that the attribute
                # holder array is going to hold the attributes in, because we will be loading these
                # attributes into very specific columns within the database

                #   index 0 :   group_id
                #   index 1 :   group_qty
                #   index 2 :   account_id
                #   index 3 :   student_id
                #   index 4 :   student_first_name
                #   index 5 :   student_last_name
                #   index 6 :   grade
                #   index 7 :   print_on_date
                #   index 8 :   date_ordered
                
                # Declare a couple Booleans that will help us process
                is_grouped = False  # Is student part of a group?
                delay_ship = False  # Is this order a delayed print order?

                # Let's begin by setting our flags

                # Is this student a part of a group?
                if no_null_attributes[row][0] != "":
                    is_grouped = True
                # Does this test need to have a delayed shipping date?
                if no_null_attributes[row][5] != "":
                    delay_ship = True

                # Now we can separate the first 2 indices into the 4 attributes that they hold 
                # (group_id, group_qty, account_id, and student_id)

                # If this student is a part of a group
                if is_grouped:
                    group_separator = no_null_attributes[row][0].split('-')
                    # Add the group id
                    this_tuple_data.append(int(group_separator[0]))
                    # Add the group test quantity
                    this_tuple_data.append((group_separator[1]))
                # Not apart of a group, but let's still fill those indices with NULL values
                else:
                    # Fill the group id column with a NULL value
                    this_tuple_data.append(None)
                    # Fill the group quantity column with a NULL value
                    this_tuple_data.append(None)
                # account_id and student_id held in index 1
                student_separator = no_null_attributes[row][1].split('-')
                # Add the account_id
                this_tuple_data.append(student_separator[0])
                # Add the student_id
                this_tuple_data.append(student_separator[1])

                # Now that we've parsed those first 2 indices, the rest will be a cake walk

                # Add the student first name (index 2)
                this_tuple_data.append(no_null_attributes[row][2])
                # Add the student last name (index 3)
                this_tuple_data.append(no_null_attributes[row][3])
                # Add the student grade (index 4)
                this_tuple_data.append(no_null_attributes[row][4])

                # Let's set a print date

                # Set the format that will be compatible with SQL Date objects: YYYY-MM-DD
                format = "%Y-%m-%d"

                # If it's a delayed ship, we do some math
                if delay_ship:
                    # Let's convert the date within this CSV column to an actual date object
                    date_object = datetime.strptime(no_null_attributes[row][5], format).date()
                    # Per Hewitt's request, we subtract 18 days from the delayed ship date to
                    # calculate the print_on_date
                    print_on_date = date_object - timedelta(days=18)
                    
                # It's not a delay ship, give it an immediate print_on_date 
                else:

                    print_on_date = today
                    # Convert to date object
                    print_on_date = datetime.strptime(print_on_date, format).date()
                # Load the print_on_date into the array
                this_tuple_data.append(print_on_date)

                # Convert the CSV column date ordered to a date object compatible with SQL
                date_ordered = datetime.strptime(no_null_attributes[row][6], format).date()
                # Add date ordered to the array
                this_tuple_data.append(date_ordered)

                # Lastly, we need to make sure that these test_orders get pre-loaded with the 
                # oustanding_order flag being set to 0!
                outstanding_flag = 0
                this_tuple_data.append(outstanding_flag)

                # We should now have each of our tuples built within an array
                ############################################################################################

                ########## CONVERT ARRAYS INTO TUPLES ######################################################
                #   index 0 :   group_id
                #   index 1 :   group_qty
                #   index 2 :   account_id
                #   index 3 :   student_id
                #   index 4 :   student_first_name
                #   index 5 :   student_last_name
                #   index 6 :   grade
                #   index 7 :   print_on_date
                #   index 8 :   date_ordered
                #   index 9 :   outstanding_order
                # The SQL syntax needed to import data into the database requires tuples:
                # We load this tuple putting the attributes in the same order as they show up in the DB:
                #                          0: account_id,       1: student_id,          2: print_on_date,  3:outstanding_order, 4: group_id,         5: group_qty,      6:date_ordered,      7: grade,        8: student_first_name 9: student_last_name
                tuple_holder.append((int(this_tuple_data[2]), int(this_tuple_data[3]), this_tuple_data[7], this_tuple_data[9], this_tuple_data[0], this_tuple_data[1], this_tuple_data[8], this_tuple_data[6], this_tuple_data[4], this_tuple_data[5]))

                # We now have an array holding all of the tuples of data that need to be entered into the DB
            #################################################################################################

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

            # Instantiate DB object
            my_cursor = hewitt_db.cursor()


            # Define the SQL query that will update the date_printed attribute for these tests to today
            update_query = """INSERT INTO test_order_tester (account_id, student_id, print_on_date, outstanding_order, group_id, group_qty, date_ordered, grade, student_first_name, student_last_name)
                            VALUES(?,?,?,?,?,?,?,?,?,?);"""
            # The executemany function will execute multiple UPDATE commands within 1 function call
            my_cursor.executemany(update_query, tuple_holder) 
            # HAVE TO COMMIT YOUR DATABASE UPDATES
            hewitt_db.commit()
        ####################################################################################################
        
        ############## EXCEPTION/ERROR HANDLING ############################################################

        # It's important that we implement some exception handling for Hewitt's scripts.  When you're
        # working with large scale databases, the chances of errors and discrepancies are very high.
        # Python uses try/exception blocks (as opposed to C's try/catch blocks). Here, we've placed 
        # the entire script within a try block, and if anything goes wrong throughout the execution 
        # of the script, we will send an error message to the desktop application provoking Hewitt to 
        # look into the error.  For this file, since we are reading from a CSV file, parsing the 
        # column data, and importing new test orders into the database, the following is a small 
        # list of potential errors that might arise:
        #       -Some sort of file path error connecting to the CSV files
        #       -Internet connectivity issue
        #       -Error connecting to MS SQL database
        #       -Somehow the CSV file didn't get saved as a .CSV extension
        #       -Some sort of formatting error within the data in the CSV file
        #       -Date columns in the CSV file not saved in YYYY-MM-DD format
        #       -Miscellaneous database inconsistency with specific student tuple/s

        # If something went wrong in the script, the compiler will enter the following sequence to
        # throw a notification to the desktop app to let Hewitt know of the error.
        except:
            self.exception_thrown = True
        #####################################################################################################

# # NOTE: the following is a class instantiation and function call used only for testing 
# my_example = wooCommerce()
# my_example.find_deadline_exceeded()
# my_example.add_online_orders()