# Shane Snediker for Hewitt Research Foundation
# Enhance the look of the desktop application that the Whitworth Software Engineering team built
# Last updated: 6-22-2021

# Within this file is the source code for the primary PyQt5 GUI application that we are constructing
# to enhance the aesthetics of Hewitt's scripting desktop application.  The desktop application
# will be used by Hewitt to run the printing and CSV processing scripts needed for Hewitt's day
# to day operations

########### DOCUMENTATION ####################################################################
#
#   Unlike the TKinter Python library that we used originally to try to construct a GUI
#   for Hewitt's Desktop Application, there is a lot of online documentation and help for
#   the PyQt library.  The following are a few sources of documentation that I found to be
#   very helpful as I constructed this interface:
#
#   Basic PyQt5 tuturial (including environment and installation setup): 
#       https://guiguide.readthedocs.io/en/latest/gui/qt.html
#
#   QMessageBox widget tutorial:      https://www.tutorialspoint.com/pyqt5/pyqt5_qmessagebox.htm
#   "                " documentation: https://doc.qt.io/qtforpython-5/PySide2/QtWidgets/QMessageBox.html
#
#   QThread tutorial:      https://realpython.com/python-pyqt-qthread/
#   "     " documentation: https://doc.qt.io/qtforpython/PySide6/QtCore/QThread.html
#
#   Some YouTube Videos were also very helpful during my initial stages of understanding
#   how to construct the application using a foundation of multithreading.  A few of the
#   more helpful YouTube URL's are:
#       YouTube QThread (15 min): https://youtu.be/G7ffF0U36b0
#       YouTube QThread (8 min):  https://youtu.be/k5tIk7w50L4
#       YouTube QThread (16 min): https://youtu.be/eYJTcLBQKug
#       YouTube QThread (33 min): https://youtu.be/Hwk242UMFR8
#
#   The following Stack Overflow threads were also extremely helpful for understanding QThreads:
#       https://stackoverflow.com/questions/52973090/pyqt5-signal-communication-between-worker-thread-and-main-window-is-not-working
#       https://stackoverflow.com/questions/37687463/single-worker-thread-for-all-tasks-or-multiple-specific-workers

#
#   Last, but certainly not least is the tutorial at the following link.  It really was this 
#   tutorial that got me over the edge in understanding how to run scripts on separate threads
#   that can run concurrently without interrupting the GUI main window event handler.  The core
#   structure of the application that I have constructed is based in a large part around this 
#   tutorial:
#      https://www.semicolonworld.com/question/58279/pyqt5-qthread-signal-not-working-gui-freeze

##################################################################################################

# Import libraries
import sys
import os
import time
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from add_online_orders import wooCommerce
from print_tests import print_tests
from print_labels import little_labels
from export_shipping_csv import ship_labels


################# DECLARE COLOR CONSTANTS #######################################################

# We want to use solid, consistent Hewitt branding for this application, so we will incorporate
# their color schemes.  Hewitt has provided me with a branding guide with their rgb color scheme. 
# Hewitt's shade of blue from within their branding emblem is rgb(40, 77, 99)
hewitt_blue = QtGui.QColor(40, 77, 99)
# Hewitt's shade of green from within their emblem is rgb(61, 112, 115)
hewitt_green = QtGui.QColor(61, 112, 115)
# The purple from within their emblem is rgb(116, 86, 116)
hewitt_purple = QtGui.QColor(116, 86, 116) 
# The grey/white from within their branding emblems is rgb(217, 217, 217)
hewitt_white = QtGui.QColor(217, 217, 217)
#################################################################################################

############# BRIEF QTHREAD DESCRIPTION #########################################################

#   Within the PyQt5 framework, every object that we send to the application screen is a child
#   of the QObject base class.  The PyQt5 library is an event-driven framework, which means that
#   every widget added to a GUI is connected to the boundaries of an event handler that controls
#   the overall flow of the GUI.  Windows and objects being sent to the screen are all queued 
#   within the event handler.  The problem that I was having in the early stages of app 
#   development involved allowing Hewitt to initiate the running of their scripts with by 
#   pressing a button.  Hewitt's scripts take time to run, and the event of the scripts running
#   was causing the main window of the application to time out and subsequently leading to a
#   frozen GUI.  Thus, I had to implement the QThread library so that when Hewitt runs a script,
#   the script can run on its own thread and the main window can stay operational unimpeded.
#
#   Within the QThread definition is a powerful system of what is called signals and slots.  
#   Signals and slots is how separate threads can communicate with each other and allow an
#   event from 1 thread to impact or effect events from another thread.  Signals are class
#   member variables that can be emitted at specific times within a thread to signify an
#   important event.  Signals are connected to slots, which are class methods.  When a signal
#   is emitted, it triggers the slot that it is connected to.  For example, when Hewitt runs 
#   one of their scripts, the thread that it is running on will emit specific signals at very
#   specific junctures within the script.  The signals will be connected to the primary window
#   thread such that other popup windows and dialogue widgets can appear as a result of what
#   is taking place in the script threads.  
#
#   Final NOTE: WITHIN PYQT5, ALL GUI COMPONENTS MUST COME FROM THE PRIMARY ORIGINAL CLASS

############## EXCEPTION HANDLING ##############################################################

# If the application raises uncaught exception, print the pertinent info
def trap_exc_during_debug(*args):
    print(args)

# Install an exception hook: without this, uncaught exception would cause application to exit
sys.excepthook = trap_exc_during_debug
#################################################################################################

########### DEFINE WORKER CLASS THAT IS USED TO RUN HEWITT'S SCRIPTS ############################

# We encapsulate our GUI components within a class structure.  This allows us to utilize OOP
# to make connections between the varying components of the GUI so that they can interact and
# manipulate each other in the desired ways

# Define a class that will initiate and fuel the operations of the 4 Hewitt scripts 
# NOTE: this class must inherit from the QObject class in order to be able to emit signals, 
#       connect signals to slots and operate within a QThread object
class script_runner(QObject):

    # CLASS MEMBER VARIABLES

    # Boolean flag used to confirm that Hewitt pressed 'Ok' to begin a script
    confirm_script_begin = False
    
    # PyQt5 QThread Signals
    begin_script_sig = pyqtSignal()     # Emitted upon script commencement
    found_past_due_sig = pyqtSignal()   # Emitted after the database is checked for outstanding 
    #                                     test orders that have exceeded Hewitt's 90 day dealine
    orders_added_sig = pyqtSignal()     # Emitted upon successful completion of the add_online_orders function
    reprint_scantron_sig = pyqtSignal() # Emitted in order to determine if Hewitt needs to reprint any scantrons
    reprint_label_sig = pyqtSignal()    # Emitted in order to determine if Hewitt needs to reprint little labels
    export_success_sig = pyqtSignal()   # Emitted in order to tell Hewitt that the CSV export script finished successfully
    exception_sig = pyqtSignal()        # Emitted if a script throws an exception
    script_complete_sig = pyqtSignal()  # Emitted after a script completes to terminate the thread

    # Class constructor defines the object that we will use to run Hewitt's scripts
    def __init__(self):
        super().__init__()

        # Give this script runner class an instance of each script class so that
        # we can access its functionality
        self.addOrders = wooCommerce()
        self.scantrons = print_tests()
        self.labels = little_labels()
        self.shipper = ship_labels()

    # CLASS METHODS/SLOTS

    # WOOCOMMERCE SCRIPT RUNNER
    @pyqtSlot() # PyQt5 decorator function
    def run_woo_orders_script(self):
        
        ############# CHECK FOR PAST DUE EXAMS ##########################################################

        # Begin by checking Hewitt's database for any outstanding test orders that have 
        # eclipsed their 90 day deadline
        self.addOrders.find_deadline_exceeded()
        # Did the find_deadline_exceeded script have any errors?
        if self.addOrders.exception_thrown:
            # Emit the error notification window
            self.exception_sig.emit()
            self.script_complete_sig.emit()
        # Nope, everything went fine! Let Hewitt know it:
        else:
            # We need to pause the running of this script to allow 
            # Emit the message window that will inform Hewitt whether we found any past due test orders
            # NOTE: we don't want the add_online_orders function to start until Hewitt has pressed
            #       'OK' on the past due function notification window, so the following signal will
            #       handle notifying Hewitt of whether or not any past due exams were found (as well
            #       as how many were found).  As soon as Hewitt clicks 'OK', the script will continue 
            #       by adding the CSV WooCommerce orders to the DB
            self.found_past_due_sig.emit()
        #####################################################################################################

    # SCANTRON PRINT SCRIPT RUNNER
    @pyqtSlot() # PyQt5 decorator function
    # Method that is used to run the printing of scantrons
    def run_scantron_script(self):
        
        # Run the script to print student scantron tests
        self.scantrons.print_scantrons()

        # Now that the script has ran, determine if it ran successfully or if it 
        # ran into some sort of error or exception
        # If an error happened during the execution of the script, provide
        # Hewitt with a notification QMessageBox with some details
        if self.scantrons.exception_thrown:
            # Emitting the exception signal will send a QMessageBox to the GUI screen
            self.exception_sig.emit()
            self.script_complete_sig.emit()
            
        # Since an exception wasn't thrown, we know that the script ran successfully,
        # so let's confirm for Hewitt the success of the script
        else:
            # Emitting the reprint signal sends a notification to the GUI asking Hewitt
            # if they need to have any of the scantrons reprinted
            self.reprint_scantron_sig.emit()
            self.script_complete_sig.emit()

    # PRINT LITTLE LABELS SCRIPT RUNNER 
    @pyqtSlot() # PyQt5 decorator function
    def run_labels_script(self):

        # Run the script to print little labels
        self.labels.print_labels()

        # Now that the script has ran, determine if it ran successfully or if it 
        # ran into some sort of error or exception

        # If an error happened during the execution of the script, provide
        # Hewitt with a notification QMessageBox with some details
        if self.labels.exception_thrown:
            # Emitting the exception signal will send a QMessageBox to the GUI screen
            self.exception_sig.emit()
            self.script_complete_sig.emit()
            
        # Since an exception wasn't thrown, we know that the script ran successfully,
        # so let's confirm for Hewitt the success of the script
        else:
            # Emitting the reprint signal sends a notification to the GUI asking Hewitt
            # if they need to have any of the scantrons reprinted
            self.reprint_label_sig.emit()
            self.script_complete_sig.emit()
        

    # EXPORT SHIPPING INFO CSV SCRIPT RUNNER
    @pyqtSlot() # PyQt5 decorator function
    def run_ship_info_script(self):

        # Run the shipping info CSV export script
        self.shipper.export_shipping_csv()

        # Now that the script has ran, determine if it ran successfully or if it 
        # ran into some sort of error or exception

        # If an error happened during the execution of the script, provide
        # Hewitt with a notification QMessageBox with some details
        if self.shipper.exception_thrown:
            # Emitting the exception signal will send a QMessageBox to the GUI screen
            self.exception_sig.emit()
            self.script_complete_sig.emit()
            
        # Since an exception wasn't thrown, we know that the script ran successfully,
        # so let's confirm for Hewitt the success of the script
        else:
            # Emitting the export success signal
            self.export_success_sig.emit()
            self.script_complete_sig.emit()

##################################################################################################

################ MAIN WINDOW CLASS DEFINITION ####################################################

class main_window(QWidget):

    # CLASS METHODS/SIGNALS

    # Confirm Script is ready to begin signal
    confirm_begin_sig = pyqtSignal()

    # Class constructor defines the characteristics of the main application window
    def __init__(self):
        super().__init__()

        self.__threads = None

        ############## DEFINE THE MAIN WINDOW'S CHARACTER ##########################################################################

           ### CSS STYLING ###

        # Here we initialize a long string variable that will hold the CSS styling rules for each of our
        # individual widget elements.  An element is connected to specific CSS rules by its PyQt widget name
        # So here we define styling for:
        #   QWidget : the main window frame of this application we set to a black background
        #   QInputDialog : we set the frame of the input dialog message boxes to white
        #   QLabel : The greeting at the top of the main window frame to which we give font character and spacing specifications
        #   QPushButton : Each of the 5 buttons on the main application window we provide with a color gradient characteristic, as
        #                 well as font specs and a nice border radius that rounds the edges of the buttons so they don't look 
        #                 so generic and abrupt
        #   QPushButton::hover : Here we provide some CSS styling so that when Hewitt hovers over a button its characteristics will
        #                        will change to show that the button is being selected
        self.app_css = '''
        QWidget{background-color: black; color: ''' + hewitt_white.name() + ''';}
        QInputDialog{background-color: ''' + hewitt_white.name() + '''; color: ''' + hewitt_white.name() + ''';}
        QLabel{font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 30px; color: ''' + hewitt_white.name() + '''; font-weight: 900;}
        QPushButton{font-family: Georgia, serif; font-size: 20px; background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 ''' + hewitt_purple.name() + ''', stop:1 ''' + hewitt_green.name() + '''); border-radius: 21px; color: ''' + hewitt_white.name() + ''';}
        QPushButton::hover{color: ''' + hewitt_purple.name() + '''; background-color: ''' + hewitt_white.name() + ''';}
        '''
        
           ### WINDOW DESIGN ###      

        # Define the size and location of the window
        # Args: x location of upper left window corner, y location of upper left window corner, 
        #       horizontal pixel length of window, vertical pixel length of window
        self.setGeometry(500, 200, 900, 700)

        # Disable window maximize function so that the application window will remain a fixed size
        self.setFixedSize(900, 700)
        
        # Define the title of the desktop application window
        self.setWindowTitle("                         Hewitt Script Manager")

        # Replace generic window icon with a Hewitt personalized image
        self.setWindowIcon(QtGui.QIcon('img\\icon_image.jpg'))

        # Apply widget CSS styling
        self.setStyleSheet(self.app_css)

        ################# ADD AND DEFINE GREETING LABEL ##############################################################################
        
        # Now we can place a label on the window that will greet the Hewitt administrators
        # when they use this application
        self.greeting = QtWidgets.QLabel(self)
        # Define the text of the label
        self.greeting.setText("Hello Hewitt. \n Welcome to your printing and \n    data import/export center!")
        # Make sure the label is long enough to fit our greeting
        self.greeting.resize(600, 150)
        # Specify the location of the label (x location, y location within the window)
        self.greeting.move(140, 20)
        self.greeting.setAlignment(QtCore.Qt.AlignCenter)
        
        # Add a cool little shadow effect to the greeting label
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(100)
        shadow.setColor(hewitt_green)
        self.greeting.setGraphicsEffect(shadow)

        ##############################################################################################################################

        ################# ADD AND DEFINE SCRIPT BUTTONS ##############################################################################

        # NOTE : I've decided to lay out the buttons on the GUI main page in the order that
        #        I think Hewitt may typically run the scripts: Add online WooCommerce orders
        #        - print scantrons - print little labels - export shipping CSV info

        # NOTE : Because each button is going to call the same start_scripts function, I've 
        #        added each button to a PyQt5 object called a QButtonGroup.  A QButtonGroup
        #        holds a group of QPushButton objects and provides functionality for connecting
        #        multiple buttons to the same slot function and still being able to identify which
        #        of the buttons has been clicked within an application.  So, for each button the
        #        setCheckable() function is used to make the buttons available to be accessed 
        #        within the QButtonGroup

        # Instantiate and define the button that will be used to add WooCommerce orders to Hewitt's database.  
        self.wooCommerce_button = QtWidgets.QPushButton('Add WooCommerce Orders', self)
        self.wooCommerce_button.setCheckable(True)
        self.wooCommerce_button.resize(350, 50)
        self.wooCommerce_button.move(260, 250)
        
        # Instantiate and define the button that will be used to print student scantron exams
        self.print_scantrons_button = QtWidgets.QPushButton('Print Scantrons', self)
        self.print_scantrons_button.setCheckable(True)
        self.print_scantrons_button.resize(350, 50)
        self.print_scantrons_button.move(260, 320)

        # Instantiate and define the button that will be used to print little organizational shipping labels
        self.print_labels_button = QtWidgets.QPushButton('Print Little Labels', self)
        self.print_labels_button.setCheckable(True)
        self.print_labels_button.resize(350, 50)
        self.print_labels_button.move(260, 390)

        # Instantiate and define the button that will be used to export shipping info CSV files
        self.export_ship_info_button = QtWidgets.QPushButton('Export Shipping CSV', self)
        self.export_ship_info_button.setCheckable(True)
        self.export_ship_info_button.resize(350, 50)
        self.export_ship_info_button.move(260, 460)

        # Instantiate and define a button that will be used to close Hewitt's desktop application
        self.close_button = QtWidgets.QPushButton('Close Application', self)
        self.close_button.resize(350, 50)
        self.close_button.move(260, 600)
        self.close_button.clicked.connect(self.close_application)

        # Create a QButtonGroup object that will hold the primary Hewitt script buttons
        self.button_group = QButtonGroup()
        self.button_group.setExclusive(True)    # Necessary for differentiating clicked buttons
        # Add each script button to the QButtonGroup
        self.button_group.addButton(self.wooCommerce_button)
        self.button_group.addButton(self.print_scantrons_button)
        self.button_group.addButton(self.print_labels_button)
        self.button_group.addButton(self.export_ship_info_button)

        # Connect the QButtonGroup to the start_scripts function
        self.button_group.buttonClicked.connect(self.start_scripts)
        ##############################################################################################

    # Class method that initializes the threads and script_runner object that we use to run the scripts
    # This method disables the desktop buttons (so that Hewitt won't be able to accidently and
    # inadvertently start another script while one is currently working by pressing a button), 
    # instantiates a script_runner object (which provides access to Hewitt's 4 script classes),
    # begins a new QThread object external to the primary GUI thread, moves the script_runner
    # object to the new thread, connects the script_runner QThread signals to their corresponding
    # slots and then starts the thread 
    def start_scripts(self):
        # Begin by disabling the GUI buttons so that Hewitt cannot inadvertently initiate another
        # script by accidently clicking a button
        self.wooCommerce_button.setDisabled(True)
        self.print_scantrons_button.setDisabled(True)
        self.print_labels_button.setDisabled(True)
        self.export_ship_info_button.setDisabled(True)
        self.close_button.setDisabled(True)

        # Declare a Python list to hold the thread, script_runner object pair
        self.__threads = []

        # Instantiate a script_runner object that we can use to access script functions
        script_worker = script_runner()
        # Instantiate a QThread object that we can send this script function to
        script_thread = QThread()
        # Save the thread/worker tuple pair in our Python list
        self.__threads.append((script_thread, script_worker))
        # Move the script_runner object to its own thread so that the running of the script
        # won't freeze up the main GUI window  
        script_worker.moveToThread(script_thread)

        # Connect the script_runner signals to their corresponding slots
        # The lambda functions allow us to pass a script_worker object as a parameter so
        # that we can access the member variables of the corresponding script classes within the slots
        script_worker.begin_script_sig.connect(lambda: self.initial_messageBox(script_worker))
        script_worker.found_past_due_sig.connect(lambda: self.is_past_due(script_worker))
        script_worker.orders_added_sig.connect(lambda: self.add_wooCommerce_orders(script_worker))
        script_worker.reprint_scantron_sig.connect(lambda: self.reprint_scantron_dialogue(script_worker))
        script_worker.reprint_label_sig.connect(lambda: self.reprint_label_dialogue(script_worker))
        # The following 3 signals do not need to reference the script_runner class, so lambda function is not necessary
        script_worker.export_success_sig.connect(self.export_success)
        script_worker.exception_sig.connect(self.application_error)
        script_worker.script_complete_sig.connect(self.terminate_thread)

        # Differentiate which button was clicked so that the the correct script function
        # can be called for the correct button and connect the script thread to the 
        # corresponding script_runner method that consists of the full Hewitt script

        # Was the Add WooCommerce Button clicked?
        if (self.button_group.checkedButton().text() == "Add WooCommerce Orders"):
            script_thread.started.connect(script_worker.run_woo_orders_script)
        # Was the Print Scantrons Button clicked?            
        elif (self.button_group.checkedButton().text() == "Print Scantrons"):
            script_thread.started.connect(script_worker.run_scantron_script)
        # Was the Print Little Labels Button clicked?
        elif (self.button_group.checkedButton().text() == "Print Little Labels"):
            script_thread.started.connect(script_worker.run_labels_script)
        # Was the Export Ship Info Button clicked?
        elif (self.button_group.checkedButton().text() == "Export Shipping CSV"):
            script_thread.started.connect(script_worker.run_ship_info_script)

        # Begin by emitting script initialization signal to make sure Hewitt is ready to begin
        script_worker.begin_script_sig.emit()
        time.sleep(1)
        # If Hewitt happened to cancel out or X out from the script initialization message
        # box, make sure that the script doesn't still run.  Therefore, we've created a 
        # Boolean variable confirm_script_begin in the script_worker script_runner() object.  
        # The script initialization QMessageBoxes will only turn the confirm_script_begin True
        # if Hewitt presses the 'Ok' button, otherwise it will stay false.  So let's check to
        # make sure that Hewitt is ready to start the script 
        if script_worker.confirm_script_begin:  
            # Start the script thread's event loop
            script_thread.start()  
        # The confirm_script_begin is still false, which means Hewitt must not be ready to start
        # the script (they've exited out of the initialization window without clicking 'Ok'). 
        # Let's at least re-enable the GUI buttons so that Hewitt can still interact with the 
        # main window if they so desire.
        else:
            # Turn the buttons back on
            self.wooCommerce_button.setEnabled(True)
            self.print_scantrons_button.setEnabled(True)
            self.print_labels_button.setEnabled(True)
            self.export_ship_info_button.setEnabled(True)
            self.close_button.setEnabled(True)
    ################################################################################################
    
    # Class method used to close the desktop application.  The close application button will be
    # connected to this method and when Hewitt presses the button, the application will close.
    def close_application(self):
        self.close()

    ############ SLOTS #############################################################################

    # Define a slot method that will send a QMessageBox widget to the screen to allow Hewitt
    # to initialize the running of the script.
    @pyqtSlot() # PyQt5 decorator function
    def initial_messageBox(self, reference):
        ############ PROVIDE HEWITT SCRIPT INITIALIZATION PRIVILEGES ##############################

        # Let's add a confirmation message to the screen that they will press in order to 
        # start the script     
        start_script = QMessageBox()
        # Set the message title
        start_script.setWindowTitle("            Script Initialization")
        # Provide Hewitt branding by changing the generic error window icon to Hewitt emblem
        start_script.setWindowIcon(QtGui.QIcon('img/icon_image.jpg'))
        # Give a little CSS styling to the message
        start_script.setStyleSheet("QLabel {font-size: 18px; color: black;} ")
        # Set the text of the message
        start_script.setText("When you are ready to start this script, please press the OK button.")
        # Add an option for Hewitt to cancel
        start_script.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        start_script.setDefaultButton(QMessageBox.Ok)
        # Make the QMessageBox appear
        present_msg = start_script.exec_()
        # Identify whether or not Hewitt pressed the 'Ok' button, because if they didn't
        # we can't let the script begin
        if present_msg == QMessageBox.Ok:
            # Hewitt is officially ready to start this script, set the start flag
            reference.confirm_script_begin = True

    # Define a slot method that will send a notification to the GUI informing Hewitt whether
    # or not our check for outstanding test orders that have exceeded the 90 day deadline 
    # retrieved any tuple results
    @pyqtSlot() # PyQt5 decorator function
    def is_past_due(self, reference):

        # Don't forget to re-enable the GUI buttons as soon as the script is completed
        self.wooCommerce_button.setEnabled(True)
        self.print_scantrons_button.setEnabled(True)
        self.print_labels_button.setEnabled(True)
        self.export_ship_info_button.setEnabled(True)
        self.close_button.setEnabled(True)

        # Let's add a confirmation message to the screen that they will press in order to 
        # start the script     
        past_due = QtWidgets.QMessageBox(self)
        # Set the message title
        past_due.setWindowTitle("            Check For Past Due Test Orders")
        # Provide Hewitt branding by changing the generic error window icon to Hewitt emblem
        past_due.setWindowIcon(QtGui.QIcon('img/icon_image.jpg'))
        # Give a little CSS styling to the message
        past_due.setStyleSheet("QMessageBox {background-color: " + hewitt_white.name() + ";} QLabel {font-size: 18px; color: black; background-color: " + hewitt_white.name() + ";} ")
        # The specific text of the screen will change depending on whether any outstanding past due
        # test orders were find, so did we find any?
        if reference.addOrders.deadline_exceeded:
            # Tell Hewitt that we found some outstanding past due test exams and sent the student info to the CSV file
            past_due.setText("We found " + str(reference.addOrders.num_unreturned) + " past due test orders and have changed the outstanding_order attribute from a value of 2 to a value of 4. We sent all relevant information for these past due orders to the CSV file found at \nU:\Work\Testing\CSV_Output\past_due_tests.csv")
        else:
            # Tell Hewitt that there aren't any outstanding past due exams at this time
            past_due.setText("We've checked the database and there are no outstanding past due test orders at this time.")
        # Make the QMessageBox appear
        present_msg = past_due.exec_()
        # Identify the moment when the WooCommerce script has successfully completed checking
        # the DB for outstanding past due test orders so that then and only then will the script
        # begin to add test orders to the DB
        if present_msg == QMessageBox.Ok:
            # Now we've confirmed that Hewitt has pressed 'OK' on the past due test order 
            # notification window, so we can commence adding orders to the DB
            reference.addOrders.add_online_orders()

            # Did we run into any issues adding orders into the database?
            if reference.addOrders.exception_thrown:
                # Then emit the error notification window
                reference.exception_sig.emit()
                reference.script_complete_sig.emit()
            # Nope! Everything went great.  Inform Hewitt that the DB has been updated with the new orders.
            else:
                # Give the GUI a few moments to rest before immediately prompting Hewitt of success
                time.sleep(1)
                # Emit the signal that will present Hewitt with confirmation that they've successfully added online orders
                reference.orders_added_sig.emit()
                reference.script_complete_sig.emit()

    # Define a slot method that will send Hewitt confirmation that their WooCommerce orders
    # have been successfully added to their database
    @pyqtSlot() # PyQt5 decorator function
    def add_wooCommerce_orders(self, reference):

        # Don't forget to re-enable the GUI buttons s soon as the script is completed
        self.wooCommerce_button.setEnabled(True)
        self.print_scantrons_button.setEnabled(True)
        self.print_labels_button.setEnabled(True)
        self.export_ship_info_button.setEnabled(True)
        self.close_button.setEnabled(True)

        # Instantiate success message 
        success_msg = QMessageBox()
        # Set the message title
        success_msg.setWindowTitle("            Success!")
        # Provide Hewitt branding by changing the generic error window icon to Hewitt emblem
        success_msg.setWindowIcon(QtGui.QIcon('img/icon_image.PNG'))
        # Give a little CSS styling to the message
        success_msg.setStyleSheet("QButton {color:rgb(62, 109, 112); font-family: Georgia; font-size:26px; background-color: rgb(81, 51, 96)} QPushButton::hover {color: rgb(75, 181, 67); background-color: black} QLabel {font-size: 18px; color: rgb(75, 181, 67)} ")
        # Set the text of the message
        success_msg.setText("The test orders have successfully been added to the database.")
        # Send message to the window
        present_msg = success_msg.exec_()

    # Define a slot method that will send a dialogue interface to the screen so that Hewitt
    # can reprint scantrons if necessary.  The method will provoke an input widget asking
    # first if Hewitt successfully printed all of their student exams.  If they did print all
    # needed exams, the script will peacefully exit.  Otherwise, an input dialogue will be 
    # presented provoking Hewitt to input the name of the last student whose exam successfully
    # printed so that we can begin the script at the next student in the queue
    @pyqtSlot() # PyQt5 decorator function
    def reprint_scantron_dialogue(self, reference):
        # We begin by giving Hewitt the option to end the script or reprint scantrons
        # We use a series of QInputDialog widgets to receive user input from Hewitt
        # QInputDialog widgets are the best way to interact with the GUI user

        # First let the printing proceed for a little while before automatically sending
        # a QInputDialog to the screen
        time.sleep(15)  # 15 seconds of downtime
        # Define a tuple with the answers that will populate the first QInputDialog widget
        yes_or_no = ("Yes", "No")
        # QInputDialog's getItem() function returns the user's selected answer as well as a 
        # Boolean value corresponding to whether the user clicked the "ok" button or not
        # The arguments for the following getItem() function are as follows: a widget object
        # where the QInputDialog box will attach itself, the QInputDialog box title, the label
        # that will appear above the drop down menu, the number of the current item (default is 0)
        # and a Boolean value signifying whether the QInputDialog will be editable
        choice, ok = QInputDialog.getItem(self, "                Need a reprint?", "Did you receive a complete print job?", yes_or_no, 0, False)
        # If Hewitt received all of their scantron tests
        if choice == "Yes" and ok:
            # Then go ahead and send a successful script completion notification to the window

            # Don't forget to re-enable the GUI buttons s soon as the script is completed
            self.wooCommerce_button.setEnabled(True)
            self.print_scantrons_button.setEnabled(True)
            self.print_labels_button.setEnabled(True)
            self.export_ship_info_button.setEnabled(True)
            self.close_button.setEnabled(True)

            # Instantiate success message 
            success_msg = QMessageBox()
            # Set the message title
            success_msg.setWindowTitle("            Success!")
            # Provide Hewitt branding by changing the generic error window icon to Hewitt emblem
            success_msg.setWindowIcon(QtGui.QIcon('img/icon_image.jpg'))
            # Give a little CSS styling to the message
            success_msg.setStyleSheet("QButton {color:rgb(62, 109, 112); font-family: Georgia; font-size:26px; background-color: rgb(81, 51, 96)} QPushButton::hover {color: rgb(75, 181, 67); background-color: black} QLabel {font-size: 18px; color: rgb(75, 181, 67)} ")
            # Set the text of the message
            success_msg.setText("The test orders have successfully printed and been assigned today's date for their date_printed attribute within the database.")
            # Send message to the window
            present_msg = success_msg.exec_()
        # Otherwise, Hewitt needs to reprint some of the scantrons that didn't get printed
        else:
            # Send a QInputDialog to the GUI to allow Hewitt to input the name of the last
            # student whose test was successfully printed
            student_name, ok = QInputDialog.getText(self, "              Provide student name", "Please type in the full name of the last student whose \ntest successfully printed and click 'Ok' (leave blank if \nno tests have successfully printed)")
            # We've saved the student's name in the student_name variable, now separate it 
            # into a list of [0] = first name and [1] = last name
            name_separator = student_name.split(' ')
            # Let's do a little error handling and take care of the chance that Hewitt mistakenly
            # forgets to put a space in between the first and the last name 
            if len(name_separator) == 1 and student_name != '' and ok:
                student_name, ok = QInputDialog.getText(self, "              Provide student name", "Oops! You have to put a space in between the first and last name please")
                name_separator = student_name.split(' ')
            # If they leave it blank, then we need to run the whole script again, but let's add
            # an empty string to the name list so that both the first and last names will have
            # values associated with them, be them empty string values
            elif student_name == '' and ok:
                name_separator.append('')
            
            # Now we've taken in the required information from Hewitt to reprint scantrons
            # Let's set some of the variables within our print_test class reference, which
            # will tell the print_scantrons function where to start the reprint job
            
            # Set the reprint flag to TRUE
            reference.scantrons.need_reprints = True
            # Assign the first name of the last student to have a successful test print
            reference.scantrons.first_name = name_separator[0]
            # Assign the last name of the last student to have a successful test print
            reference.scantrons.last_name = name_separator[1]
            # Rerun the print script
            reference.scantrons.print_scantrons()
            # Wait a few moments to let the print job take place
            time.sleep(8)
            # Don't forget to re-enable the GUI buttons as soon as the script is completed
            self.wooCommerce_button.setEnabled(True)
            self.print_scantrons_button.setEnabled(True)
            self.print_labels_button.setEnabled(True)
            self.export_ship_info_button.setEnabled(True)
            self.close_button.setEnabled(True)
            # Better double check that even the reprint produced all of the tests
            reference.reprint_scantron_sig.emit()     

    # Define a slot method that will send a dialogue interface to the screen so that Hewitt
    # can reprint little labels if necessary.  The method will provoke an input widget asking
    # first if Hewitt successfully printed all of their organizational labels.  If they did 
    # print all needed labels, the script will peacefully exit.  Otherwise, an input dialogue 
    # will be presented provoking Hewitt to input the account number of the last label that 
    # successfully printed so that we can begin the script at the next label in the queue
    @pyqtSlot() # PyQt5 decorator function
    def reprint_label_dialogue(self, reference):

        # We begin by giving Hewitt the option to end the script or reprint labels
        # We use a series of QInputDialog widgets to receive user input from Hewitt
        # QInputDialog widgets are the best way to interact with the GUI user

        # First let the printing proceed for a little while before automatically sending
        # a QInputDialog to the screen
        time.sleep(10)  # 10 seconds of downtime
        # Define a tuple with the answers that will populate the first QInputDialog widget
        yes_or_no = ("Yes", "No")
        # QInputDialog's getItem() function returns the user's selected answer as well as a 
        # Boolean value corresponding to whether the user clicked the "ok" button or not
        # The arguments for the following getItem() function are as follows: a widget object
        # where the QInputDialog box will attach itself, the QInputDialog box title, the label
        # that will appear above the drop down menu, the number of the current item (default is 0)
        # and a Boolean value signifying whether the QInputDialog will be editable
        choice, ok = QInputDialog.getItem(self, "                Need a reprint?", "Did you receive a complete print job?", yes_or_no, 0, False)
        # If Hewitt received all of their scantron tests
        if choice == "Yes" and ok:
            # Then go ahead and update the outstanding_order attribute to a 1 for each of the
            # student tuples who just had a scantron/label printed within the database
            reference.labels.update_outstanding()

            # If the outstanding_order attributes were updated successfully, we can send a success
            # message to the GUI screen, otherwise we better send an error notification
            if reference.labels.exception_thrown:
                # Don't forget to re-enable the GUI buttons as soon as the script is completed
                self.wooCommerce_button.setEnabled(True)
                self.print_scantrons_button.setEnabled(True)
                self.print_labels_button.setEnabled(True)
                self.export_ship_info_button.setEnabled(True)
                self.close_button.setEnabled(True) 

                # An exception occurred while we were trying to update the outstanding_order
                # attribute within the database, send an alert message to warn Hewitt
                # Instantiate an error message
                error_msg = QMessageBox()
                # Set the error message title
                error_msg.setWindowTitle("             Script Error")
                # Provide Hewitt branding by changing the generic error window icon to Hewitt emblem
                error_msg.setWindowIcon(QtGui.QIcon('img/icon_image.PNG'))
                # Place a red warning symbol in the message box
                error_msg.setIcon(QMessageBox.Critical)
                # The default button and textual CSS styling for the message box is set to the CSS 
                # rules of the primary desktop application window.  Let's provide unique styling
                error_msg.setStyleSheet("QPushButton {color:rgb(62, 109, 112); font-family: Georgia; font-size:26px; background-color: rgb(81, 51, 96)} QPushButton::hover {color: red; background-color: black} QLabel {font-size: 18px; color:red;}")
                # For this message box, we will have a main announcement and add an information box 
                # with a more detailed accounting of this error.  The information box can be accessed
                # by Hewitt with the click of a "more details" type button
                # Set main announcement text:
                error_msg.setText("Oops! Something went wrong trying to update the outstanding_order attributes in the database.")
                # Set information box text:
                error_msg.setDetailedText("This error message indicates that there is a discrepancy in the data that is preventing the script from updating the database.  It is recommended that you manually look into this issue.  The following is a list of possible causes of error (depending on which script you are currently trying to execute): \n -Internet connectivity issue \n -Database connection error \n -Student account database inconsistency.")
                # In order to send the message box to the screen, we've gotta call the following function
                present_msg = error_msg.exec_()
                
            else:

                # Don't forget to re-enable the GUI buttons as soon as the script is completed
                self.wooCommerce_button.setEnabled(True)
                self.print_scantrons_button.setEnabled(True)
                self.print_labels_button.setEnabled(True)
                self.export_ship_info_button.setEnabled(True)
                self.close_button.setEnabled(True)

                # Instantiate success message 
                success_msg = QMessageBox()
                # Set the message title
                success_msg.setWindowTitle("            Success!")
                # Provide Hewitt branding by changing the generic error window icon to Hewitt emblem
                success_msg.setWindowIcon(QtGui.QIcon('img/icon_image.jpg'))
                # Give a little CSS styling to the message
                success_msg.setStyleSheet("QButton {color:rgb(62, 109, 112); font-family: Georgia; font-size:26px; background-color: rgb(81, 51, 96)} QPushButton::hover {color: rgb(75, 181, 67); background-color: black} QLabel {font-size: 18px; color: rgb(75, 181, 67)} ")
                # Set the text of the message
                success_msg.setText("The outstanding_order attributes for these student test orders have successfully been changed from a value of 0 to a value of 1 within the database.")
                # Send message to the window
                present_msg = success_msg.exec_()
        
        # Otherwise, Hewitt needs to reprint some of the scantrons that didn't get printed
        else:
            
            # Send a QInputDialog to the GUI to allow Hewitt to input the last account number 
            # of the family/group/account whose label was successfully printed
            account_num, ok = QInputDialog.getInt(self, "              Provide account number", "Please type in the account number of the \nlast label that successfully printed and click \n'Ok' (leave value of zero if no labels have successfully printed)")
            
            # Now we've taken in the required information from Hewitt to reprint labels
            # Let's set some of the variables within our little_labels class reference, which
            # will tell the print_labels function where to start the reprint job
            
            # Set the reprint flag to TRUE
            reference.labels.need_reprints = True

            # If for some reason none of the little labels printed, Hewitt will input a value
            # of zero.  If they do this, we can leave the account_num variable alone and this
            # will allow the print_labels function to go with the default reprint_index of 0,
            # which will cause the function to reprint all of the labels in the queue
            if account_num != 0:
                # Assign the account_id of the last label that successfully printed
                reference.labels.acct_num = account_num
            # Rerun the print script
            reference.labels.print_labels()
            # Wait a few moments to let the print job take place
            time.sleep(6)
            # Don't forget to re-enable the GUI buttons as soon as the script is completed
            self.wooCommerce_button.setEnabled(True)
            self.print_scantrons_button.setEnabled(True)
            self.print_labels_button.setEnabled(True)
            self.export_ship_info_button.setEnabled(True)
            self.close_button.setEnabled(True)
            # Better double check that even the reprint produced all of the tests
            reference.reprint_label_sig.emit()

    # Define a slot method that will be used to send Hewitt a success notification regarding
    # the shipping info CSV export script
    @pyqtSlot() # PyQt5 decorator function
    def export_success(self):
        # Turn the rest of the buttons back on
        self.wooCommerce_button.setEnabled(True)
        self.print_scantrons_button.setEnabled(True)
        self.print_labels_button.setEnabled(True)
        self.export_ship_info_button.setEnabled(True)
        self.close_button.setEnabled(True)
        # Instantiate success message 
        success_msg = QMessageBox()
        # Set the message title
        success_msg.setWindowTitle("            Success!")
        # Provide Hewitt branding by changing the generic error window icon to Hewitt emblem
        success_msg.setWindowIcon(QtGui.QIcon('img/icon_image.PNG'))
        # Give a little CSS styling to the message
        success_msg.setStyleSheet("QButton {color:rgb(62, 109, 112); font-family: Georgia; font-size:26px; background-color: rgb(81, 51, 96)} QPushButton::hover {color: rgb(75, 181, 67); background-color: black} QLabel {font-size: 18px; color: rgb(75, 181, 67)} ")
        # Set the text of the message
        success_msg.setText("The shipping label information has been exported to the following ready shipper path: \n'U:\Work\Testing\CSV_Output\LabelsImport.csv'.  \nThe outstanding_order attribute for each of the corresponding test orders has also been changed from a value of 1 to a value of 2.")
        # Send message to the window
        present_msg = success_msg.exec_()

    # Define a slot method that will send a QMessageBox widget to the screen in the event that
    # one of the scripts throws an exception.  This will prevent the application from quitting
    # unexpectedly without providing Hewitt with a little bit of details
    @pyqtSlot() # PyQt5 decorator function
    def application_error(self):
        # As soon as the script is completed, make sure to re-enable the GUI buttons
        self.wooCommerce_button.setEnabled(True)
        self.print_scantrons_button.setEnabled(True)
        self.print_labels_button.setEnabled(True)
        self.export_ship_info_button.setEnabled(True)
        self.close_button.setEnabled(True)
        # Instantiate an error message
        error_msg = QMessageBox()
        # Set the error message title
        error_msg.setWindowTitle("             Script Error")
        # Provide Hewitt branding by changing the generic error window icon to Hewitt emblem
        error_msg.setWindowIcon(QtGui.QIcon('img/icon_image.PNG'))
        # Place a red warning symbol in the message box
        error_msg.setIcon(QMessageBox.Critical)
        # The default button and textual CSS styling for the message box is set to the CSS 
        # rules of the primary desktop application window.  Let's provide unique styling
        error_msg.setStyleSheet("QPushButton {color:rgb(62, 109, 112); font-family: Georgia; font-size:26px; background-color: rgb(81, 51, 96)} QPushButton::hover {color: red; background-color: black} QLabel {font-size: 18px; color:red;}")
        # For this message box, we will have a main announcement and add an information box 
        # with a more detailed accounting of this error.  The information box can be accessed
        # by Hewitt with the click of a "more details" type button
        # Set main announcement text:
        error_msg.setText("Oops! Something went wrong during the execution of this script.")
        # Set information box text:
        error_msg.setDetailedText("This error message indicates that there is a discrepancy in the data that is preventing the script from executing.  It is recommended that you manually look into this issue.  The following is a list of possible causes of error (depending on which script you are currently trying to execute): \n -Internet connectivity issue \n -Database connection error \n -Error connecting to a printer \n -CSV file path error \n -Student account database inconsistency.")
        # In order to send the message box to the screen, we've gotta call the following function
        present_msg = error_msg.exec_()
        
    # Define a slot method that will kill the script_runner thread after the script has
    # finished running.  This will allow Hewitt to run a script and with the main GUI 
    # window open, after the script has completed, push another button to start another
    # script.  This way we only ever have 1 additional thread running at a time above and
    # beyond the primary default GUI thread.
    @pyqtSlot() # PyQt5 decorator function
    def terminate_thread(self):
        # Access the list that is holding our worker thread
        for thread, worker in self.__threads:
            # Kill the thread
            thread.quit()
            # There's always a little awkward downtime, so call the wait function to wait
            # for the thread to fully die
            thread.wait()

if __name__ == "__main__":
    app = QApplication([])

    form = main_window()
    form.show()

    sys.exit(app.exec_())
