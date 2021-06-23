# Shane Snediker
# Hewitt Research Foundation Desktop Application
# File last updated: 6-22-2021

# This file contains the Python script for exporting test order shipping information
# to a CSV file.  The following is the breakdown of the attributes that will be extracted
# from the database VIEW called "print_ship_labels":
# 
#   [0] : customer1_first_name
#   [1] : customer1_last_name
#   [2] : address1
#   [3] : city1
#   [4] : state1
#   [5] : zipcode1
#   [6] : plus4_1
#   [7] : ship_first_name1
#   [8] : ship_last_name1
#   [9] : ship_company
#   [10] : ship_address
#   [11] : ship_city
#   [12] : ship_state
#   [13] : ship_zipcode
#   [14] : ship_plus4
#   [15] : email1
#   [16] : phone1
#   [17] : account_id
#   [18] : student_id   NOTE: this attribute is retrieved only because it is required to make the SQL query work properly
#   [19] : test_id  NOTE: this attribute is retrieved only because it is required for updating the outstanding_order attribute
#   [20] : group_id NOTE: this attribute is retrieved only for ORDER BY functionality
#   [21] : student_first_name   NOTE: this attribute is retrieved only for ORDER BY functionality
#   [22] : student_last_name    NOTE: this attribute is retrieved only for ORDER BY functionality

# NOTE: the last 5 attributes will not be used in shipping labels, they are extracted to help
#       this script uniquely identify tuples so that the outstanding_order attribute of each
#       one can be modified from a 1 to 2 signifying that the labels have been printed and 
#       tests sent and so that the query will be able to order the selection results to add
#       functional convenience for Hewitt. 

# And the following is a list of the column headers of the CSV file that we will output:

# SHIPPING FIRST NAME    SHIPPING LAST NAME    SHIPPING ADDRESS     CITY    STATE     ZIPCODE     SHIPPING COMPANY      EMAIL     PHONE     ACCOUNT ID    DATE

# Therefore the flow of this script is:

#       1. Connect to database and extract the print_ship_labels VIEW
#       2. Separate group shipping labels from individual shipping labels
#       3. Access the shipping information of the group leader for the group orders
#       4. Combine the group and individual orders back into a single array of ship label tuples
#       5. For each tuple, determine if there is a unique shipping address
#       6. Generate tuples containing the data corresponding to the CSV column headers listed above
#       7. Instantiate and export a CSV file containing this information
#       8. Update the outstanding_order attribute for each of these label tuples to a value of 2

# Import libraries
from PyQt5.QtCore import QObject    # PyQt5 functionality
import sys   # Needed for configuring the database connection
import pyodbc # Connect to MS SQL
from configparser import ConfigParser # Separate configuration settings into separate file
import csv   # Process csv files
from datetime import date   # Capture today's date

class ship_labels(QObject):

    # CLASS MEMBER VARIABLES
    exception_thrown = False

    # Throw this script into a function so that the desktop app can utilize it
    def export_shipping_csv(self):
        try:

            ########### FORMAT THE FILE NAME AND THE PATH WHERE WE WILL EXPORT THE CSV ##################

            # TEMP FILE PATH SO WE COULD TEST ON OUR OWN MACHINE
            Hewitt_csv_export_path = "C:\\Users\\Snediker\\Documents\\Whitworth\\Last call\\Software Engineering\\ready_shipper.csv"

            ###############################################################################################

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
            my_cursor.execute('SELECT * FROM print_ship_tester;') 
            # Save the tuple information in a variable that we can iterate over.  The fetchall()
            # function returns a tuple of tuples (the internal tuples consisting of individual
            # test order attributes)
            current_ship_labels = my_cursor.fetchall()
            
            # How many shipping labels will we be printing today?
            num_labels = len(current_ship_labels)
            # How many columns does each label tuple have?
            num_attributes = len(current_ship_labels[0])
            
            ########### CONVERT NULL VALUES TO EMPTY STRINGS ###################################

            # Python hates null values (or what they term 'None types')
            # We begin interacting with the data by converting all nulls to empty strings, which
            # is something that Python can deal with
            # We need to use an array in order to rebuild the tuple
            no_null_attributes = []
            # For each individual tuple:
            for ind_tup in range(num_labels):
                # check it for nulls and if found, convert to '':
                # However modifying an existing tuple is not permitted in Python, so we'll have to
                # do some extra work to create new tuples.  This is necessary because if this script
                # ever assigns an attribute that has a null value, it will crash the script
                # Track the attributes within this tuple that hold null values
                null_indices =[]
                for attribute in range(num_attributes):
                    # Does this index hold a null value?
                    if current_ship_labels[ind_tup][attribute] is None:
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
                        inner_tuple.append(current_ship_labels[ind_tup][index])
                # Finally, add this inner array as a newly created null-less tuple into the new outer array    
                no_null_attributes.append(inner_tuple)
                
                # Now we have a Python array corresponding to the data tuple with null values replaced by empty strings
            #########################################################################################################################################
            
            ## BUILD THE UPDATE_DATA ARRAY THAT WILL HOLD THE CREDENTIALS OF EACH STUDENT WHO WILL NEED THEIR OUTSTANDING_ORDER ATTRIBUTE UPDATED ###

            # First we create a Python array to hold the data we will need to make sure we only
            # change the outstanding_order attribute for the tests that we just printed
            # Our SQL test to make sure that we're only updating the specific test tuples of
            # the test orders that we just printed consists of querying the test_order table
            # and updating the outstanding_order attribute to 2 for all tuples with matching
            # account_id's, student_id's and test_id's as the ones from this test order print job
            update_data = []
            for iter in range(num_labels):
                # account_id - index 17, student_id - index 18, and test_id - index 19
                update_data.append((no_null_attributes[iter][17], no_null_attributes[iter][18], no_null_attributes[iter][19]))
            #########################################################################################################################################
            
            ############# SEPARATE GROUP STUDENT TUPLES FROM INDIVIDUAL STUDENT TUPLES ##############################################################

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
            # newly declared group_tuples array.  NOTE: group_id is located at index 20. 
            for i in range(len(no_null_attributes)):
                # Remember that we've switched NULL values to empty strings to make Python happy
                if no_null_attributes[i][20] != "":
                    # If this is the first instance of a non-NULL group_id in this array
                    if not(begin):
                        # Then save the index and set the flag
                        idx_grp_begin = i
                        begin = True
                    # Add this tuple to the group_tuples array
                    group_tuples.append(no_null_attributes[i])
            # Now delete all of the group tuples from the original 2 dimensional array so that 
            # it will be left with only the non-group tuples and we will have 2 separate arrays
            # that we can work with and manipulate for printing little labels
            del no_null_attributes[idx_grp_begin : num_labels]
            #########################################################################################################################################

            ############ IDENTIFY UNIQUE GROUP_ID'S/ACCOUNT_ID'S AND GET RID OF DUPLICATES ######################################################### 

            # UNIQUE GROUP_ID'S

            # Let's initialize a list with which we can hold unique group_id's (duplicates get left out)
            grp_id_holder = []
            # Start a loop to capture every unique group_id in the group_tuples array
            for unique_grp in range(len(group_tuples)):
                if (group_tuples[unique_grp][20] not in grp_id_holder):
                    grp_id_holder.append(group_tuples[unique_grp][20])
            
            # UNIQUE INDIVIDUAL ACCOUNT_ID'S

            # Initialize a list with which we can hold unique account_id's (duplicates are left out)
            acct_id_holder = []
            # Let's also save the indices where duplicate account_id's reside so that we can delete them later
            duplicate_idxs = []
            # Start a loop to capture every unique account_id in the array that is currently 
            # holding only individual student test order tuples
            for unique_acct in range(len(no_null_attributes)):
                if (no_null_attributes[unique_acct][17] not in acct_id_holder):
                    acct_id_holder.append(no_null_attributes[unique_acct][17])
                else:
                    duplicate_idxs.append(unique_acct)
            ########################################################################################################################################

            ######### PREPARE UNIQUE GROUP TUPLES AND INDIVIDUAL TUPLES TO BE PRINTED ##############################################################

            # INDIVIDUAL TUPLES

            # New array to save unique tuples
            ind_saver = []
            # Trim down the individual account_id array by saving only unique tuples
            for iter in range(len(no_null_attributes)):
                if iter not in duplicate_idxs:
                    ind_saver.append(no_null_attributes[iter])
            
            # Clean up
            no_null_attributes.clear()

            # GROUP TUPLES

            # For the group tuples, we are going to have to query the database for each of the unique 
            # group_id's to capture the corresponding mailing_list information (NOTE: a student's group_id
            # corresponds to the account_id of the group leader.  We need the group leader's mailing info)
            # Clear out the group_tuples list so that it will be ready to hold the group leaders' info
            group_tuples.clear()
            # Start loop to capture group leaders' info
            for query in range(len(grp_id_holder)):
                my_cursor.execute('''SELECT m.customer1_first_name, m.customer1_last_name, m.address1, m.city1, m.state1, m.zipcode1, m.plus4_1, m.ship_first_name1, m.ship_last_name1, m.ship_company, m.ship_address, m.ship_city, m.ship_state, m.ship_zipcode, m.ship_plus4, m.email1, m.phone1, m.account_id 
                FROM mailing_list_tester m
                WHERE m.account_id = ''' + str(grp_id_holder[query]) + ''';''')
                # Now pull this query and append it to our group_tuples list
                this_tuple = my_cursor.fetchall()
                group_tuples.append(this_tuple)
            #########################################################################################################################################

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
                no_null_attributes.append(inner_tuple)
            
            # Now we have a Python array corresponding to the data tuple with null values replaced by empty strings
            #########################################################################################################################################

            ################## COMBINE THE 2 ARRAYS INTO 1 ARRAY ####################################################################################
            
            # Let's combine our individual account ship label tuples with the group account ship label tuples
            # so that we can use 1 loop to build the CSV file tuples and then export them to the CSV file
            for iter in range(len(ind_saver)):
                no_null_attributes.append(ind_saver[iter])
            
            # Clean up
            del ind_saver
            del group_tuples
            ########################################################################################################################################

            ####### BEGIN LOOP TO BUILD SHIP LABEL TUPLES ###########################################################################################

            # Instantiate an array to hold each row of the the csv file
            shipping_csv_row_holder = []

            # How many shipping labels are in this batch?
            num_labels = len(no_null_attributes)

            # REMINDER OF WHAT EACH INDEX WITHIN TUPLE ARRAYS CORRESPONDS WITH
            # NOTE: the group_id tuples won't contain indices 18 - 22 because they
            #       were compiled with a different query than the individual account
            #       tuples and those attributes won't be added to the CSV export file

            #   [0] : customer1_first_name
            #   [1] : customer1_last_name
            #   [2] : address1
            #   [3] : city1
            #   [4] : state1
            #   [5] : zipcode1
            #   [6] : plus4_1
            #   [7] : ship_first_name1
            #   [8] : ship_last_name1
            #   [9] : ship_company
            #   [10] : ship_address
            #   [11] : ship_city
            #   [12] : ship_state
            #   [13] : ship_zipcode
            #   [14] : ship_plus4
            #   [15] : email1
            #   [16] : phone1
            #   [17] : account_id
            #   [18] : student_id   NOTE: this attribute was retrieved only because it 
            #                             is required to make the SQL query work properly
            #   [19] : test_id      NOTE : this attribute is retrieved only because it is 
            #                              needed in the WHERE clause of the SQL query that 
            #                              updates the outstanding_order attribute in the database
            #   [20] : group_id     NOTE : this attribute was 
            #   [21] : student_first_name   NOTE: this attribute was retrieved only for ORDER BY functionality
            #   [22] : student_last_name    NOTE: this attribute was retrieved only for ORDER BY functionality

            # Each loop processes 1 shipping label tuple and places all of it's required data within a tuple
            for label in range(num_labels):
                
                ######### DETERMINE IF THIS SHIPPING LABEL HAS A UNIQUE SHIPPING ADDRESS ############################################################

                # If the ship_first_name1 attribute is not null, then we need to use the shipping names
                use_ship_name = False
                if no_null_attributes[label][7] != "":
                    use_ship_name = True
                
                # If the ship_address attribute is not null, then we need to use the shipping info
                use_ship_info = False
                if no_null_attributes[label][10] != "":
                    use_ship_info = True
                #####################################################################################################################################

                ######## DETERMINE IF THE STATE ATTRIBUTE IS NULL (FOR OLD PARADOX TUPLES) ##########################################################

                city = ""
                state = ""
                zipcode = ""

                # Is there a unique shipping address for this label?
                if use_ship_info:
                    # If it's a newer order (one entered into the system after May of 2021) it will
                    # have separate city and state attributes.  However, every tuple entered into the
                    # Paradox system has a combined city_state attribute.  We need to separate those
                    # files for Ready Shipper
                    # If the state attribute isn't null, then we already know the city and state for this label:
                    if no_null_attributes[label][12] != "":
                        city = no_null_attributes[label][11]
                        state = no_null_attributes[label][12]
                    # However, for all the old orders, parse out the state from the city into separate columns
                    else:
                        city_state_holder = no_null_attributes[label][11].split(' ')
                        # The state will always be the last index of this new array
                        state = city_state_holder[len(city_state_holder) - 1]
                        # Now build the city by combining each of the earlier indices up to but not including the last index
                        for index in range(len(city_state_holder) - 1):
                            # Remember to leave a space in between words of a city
                            city += ' ' + city_state_holder[index]
                    # Check whether or not we're using separate shipping info and then determine if this
                    # tuple has a plus 4 that needs to be added to the zipcode
                    # Does this tuple have a separate shipping address?
                    zipcode = str(no_null_attributes[label][13])
                    if(no_null_attributes[label][14] != ""):
                        zipcode += "-" + str(no_null_attributes[label][14])

                # Nope, let's use the standard shipping info then
                else:
                    # If it's a newer order (one entered into the system after May of 2021) it will
                    # have separate city and state attributes.  However, every tuple entered into the
                    # Paradox system has a combined city_state attribute.  We need to separate those
                    # files for Ready Shipper
                    # If the state attribute isn't null, then we already know the city and state for this label:
                    if no_null_attributes[label][4] != "":
                        city = no_null_attributes[label][3]
                        state = no_null_attributes[label][4]
                    # However, for all the old orders, parse out the state from the city into separate columns
                    else:
                        city_state_holder = no_null_attributes[label][3].split(' ')
                        # The state will always be the last index of this new array
                        state = city_state_holder[len(city_state_holder) - 1]
                        # Now build the city by combining each of the earlier indices up to but not including the last index
                        for index in range(len(city_state_holder) - 1):
                            # Remember to leave a space in between words of a city
                            city += ' ' + city_state_holder[index]
                    # Check whether or not we're using separate shipping info and then determine if this
                    # tuple has a plus 4 that needs to be added to the zipcode
                    # Does this tuple have a separate shipping address?
                    zipcode = str(no_null_attributes[label][5])
                    if(no_null_attributes[label][6] != ""):
                        zipcode += "-" + str(no_null_attributes[label][6])
                ####################################################################################

                ################ BUILD CSV TUPLES ##################################################

                # Now that we've determined whether or not this label has a unique address, we can
                # load the pertinent data into a tuple that will later be exported to a CSV file
                # If this label has a unique address

                # The column headers within the CSV file that we create are as follows:
                #
                #   SHIPPING FIRST NAME     SHIPPING LAST NAME      SHIPPING ADDRESS    CITY    STATE   ZIPCODE     SHIPPING COMPANY    EMAIL    PHONE  ACCOUNT_ID  DATE

                # Let's establish today's date because it is the last column in the CSV file (it will be used
                # to help Hewitt establish a PK for the ready shipper application)
                today = date.today().strftime("%Y-%m-%d") 
                
                if (use_ship_info):

                    # Is there also a separate shipping name associated with this order?
                    if(use_ship_name):
                        #               ship_first_name               ship_last_name                  ship_address       ship_city ship_state ship_zipcode      ship_company                    email1                           phone1                         account_id
                        csv_tuple =(no_null_attributes[label][7], no_null_attributes[label][8], no_null_attributes[label][10], city, state, zipcode, no_null_attributes[label][9], no_null_attributes[label][15], no_null_attributes[label][16], no_null_attributes[label][17], today)
                    # There isn't a separate shipping name, so use customer1 first and last name
                    else:
                        # This version is the same as the one above, except it uses the standard mailing_list customer name because the shipping name is null
                        #             customer1_first_name           customer1_last_name                ship_address        ship_city ship_state ship_zipcode      ship_company                   email1                        phone1                            account_id
                        csv_tuple = (no_null_attributes[label][0], no_null_attributes[label][1], no_null_attributes[label][10], city, state, zipcode, no_null_attributes[label][9], no_null_attributes[label][15], no_null_attributes[label][16], no_null_attributes[label][17], today)
                # There is not unique shipping information, so use the standard address from the mailing list
                else:
                    # Use standard mailing info since there is no unique shipping address
                    #              customer1_first_name           customer1_last_name              address1               city1 state1 zipcode1      ship_company                    email1                             phone1                    account_id
                    csv_tuple = (no_null_attributes[label][0], no_null_attributes[label][1], no_null_attributes[label][2], city, state, zipcode, no_null_attributes[label][9], no_null_attributes[label][15], no_null_attributes[label][16], no_null_attributes[label][17], today)
                
                # Add this tuple to the array
                shipping_csv_row_holder.append(csv_tuple)

            # The following code creates a CSV file with 1 header row of column headers followed directly
            # by a row for each shipping label that needs to be printed
            
            with open(Hewitt_csv_export_path, 'w', newline='') as out:
                # Create a csv file writer object that we can use to create our CSV file
                file_writer = csv.writer(out)
                # Use the file writer object to generate the header row of the CSV file
                file_writer.writerow(['SHIPPING FIRST NAME', 'SHIPPING LAST NAME', 'SHIPPING ADDRESS', 'CITY', 'STATE', 'ZIPCODE', 'SHIPPING COMPANY', 'EMAIL', 'PHONE', 'ACCOUNT ID', 'DATE'])
                # Iterate through each shipping label, writing the corresponding attributes into each column
                for row in shipping_csv_row_holder:
                    file_writer.writerow(row)
            ###############################################################################################
            
            ##### UPDATE OUTSTANDING_ORDER COLUMN OF TEST_ORDER TABLE FOR THESE SHIPPING LABELS ##########
                
            # Now that Hewitt has confirmed that they've printed all of the labels that they
            # need, we can update the outstanding_order value for each of these tuples from
            # a 1 to a 2.  NOTE: within the test_order table an outstanding_order value of 0 
            # corresponds to all tests that haven't been printed, a value of 1 corresponds to
            # all test_orders that have been printed but not shipped yet, a value of 2 corresponds
            # to all test_orders that have been shipped but not yet returned, a value of 3
            # corresponds to all test_orders that have been returned, and a value of 4 means that
            # the customer/student didn't return the test_order in time (there's a 90 day deadline).

            # Now we build the UPDATE query that we will use to update the outstanding_order attributes
            # NOTE: Here we use string formatting syntax (?).  The ?s's get populated with the
            #       corresponding tuple data provided in the executemany function call.  So, for
            #       example we've created an array of 3 value tuples (account_id, student_id, test_id)
            #       When we call the execute many function it will go through each tuple in the array
            #       and populate the WHERE clause parameters with the corresponding attribute data  
            # NOTE: The "my_cursor.execute('SELECT * FROM Hewitt_DB.print_labels;')" function call
            #       from early in this script accessed the VIEW 'print_labels'.  Database views are
            #       kind of like virtual copies of data and cannot be directly modified.  Therefore,
            #       when we are updating database tuples, we have to access the actual tables where the
            #       original data resides
            
            # Define the SQL query that will update the date_printed attribute for these tests to today
            
            update_query = """UPDATE test_order_tester 
                            SET outstanding_order = 2 
                            WHERE account_id = ? AND student_id = ? AND test_id = ?;"""
            # The executemany function will execute multiple UPDATE commands within 1 function call
            my_cursor.executemany(update_query, update_data) 
            # HAVE TO COMMIT YOUR DATABASE UPDATES
            hewitt_db.commit()

        ############# EXCEPTION/ERROR HANDLING ##########################################################

        # It's important that we implement some exception handling for Hewitt's scripts.  When you're
        # working with large scale databases, the chances of errors and discrepancies are very high.
        # Python uses try/exception blocks (as opposed to C's try/catch blocks). Here, we've placed 
        # the entire script within a try block, and if anything goes wrong throughout the execution 
        # of the script, we will send an error message to the desktop application provoking Hewitt to 
        # look into the error.  For this file, since we are reading from a CSV file, parsing the 
        # column data, and importing new test orders into the database, the following is a small 
        # list of potential errors that might arise:
        #       -Some sort of error with the filepath where this script sends the shipping CSV file
        #       -Error connecting to the database
        #       -Internet connectivity issue
        #       -Some sort of student account database inconsistency

        # If something went wrong in the script, the compiler will enter the following sequence to
        # throw a notification to the desktop app to let Hewitt know of the error.
        except:
            self.exception_thrown = True
        
# NOTE: the following is a class instantiation and function call used only for testing 
# my_example = ship_labels()
# my_example.export_shipping_csv()