import logging
import logging.handlers
import win32evtlogutil
import win32evtlog
from datetime import datetime
import os
import time
import win32api
import win32security
import win32con
import ctypes
import sys
from typing import Iterable

class RollingEventLogger:
    def __init__(
            self, 
            log_file_path: str, 
            log_file_base_filename: str, 
            logger_name: str, 
            logging_level: str ='w', 
            debug_in_stream: bool = False, 
            add_datetime_to_log_filname: bool = False
    ) -> None:
        """
        Returns an instance of a logger which can be used to output logs created by the operation of python classes to a file.
        
        In the absence of a file to log to, rather than erroring into a failed state, the logger object will switch over to Windows Event logging instead.

        Args:
            log_file_path (str):    The path to the output location of the log file in either C:\\path\\name or C:/path/name format 
            log_file_base_filename (str):   The base filename which should be used for the log file.
                                            Can be written with or without the file extension
                                            If written without a file extension, a .LOG file will be assumed to be the preference
            logger_name (str):  This is the name applied to the logger instance
                                This can be the same as the logger filename without a file extension
            logging_level (str):    Python logging has the following levels: Debug, Info, Warning, Error, Critical
                                    The rolling logger class utilises the following notation when setting the level:
                                    d - debug:  Returns the most amount of messages from the scripts operation
                                                As such, returns messages rated debug, along with all other message types
                                    i - information:    Similar to debug, only, will not return debug messages
                                                        Info rated logging will however return all other types of messages
                                    w - warning:    Will return warning messages and above only
                                                    Will not return either Debug or Info rated messages
                                    e - error:  Will return messages rated as error or higher
                                                As such, will not return Debug or Info, but, will return levels above
                                    c - critical:   Will only return critical rated messages
                                                    All other message types will be ignored
            debug_in_stream (bool): This will return all messages outputted by the script's operation into the console
                                    As such, if you are running this from DOS, then, the output messages will be logged to the DOS prompt
            add_datetime_to_log_filename (bool):    This will add the date followed by current time to any log files created by this class
                                                    The log will divide itself across 7 x 20mb files.
                                                    Upon the 7th file reaching 20MB, the file currently labelled as file #7 will be dropped and file 6 will be renamed and take its place
                                                    This way, any time a computer is rebooted or a problem effects the script, a new instance of a stream of 7 log files will be created
                                                    This will allow for tracking of script performance and for the recall of all previous data along with any error messages that may have been captured
                                                    If the file becomes innaccessible but the script continues to function however, it will continue to utilise the original log file
        """

        self.log_file_path = log_file_path
        self.log_file_base_filename, self.log_file_base_ext = os.path.splitext(log_file_base_filename)

        self.event_source = __name__
        self.logging_level = logging_level.lower()
        self.handler_base_name = str(logger_name)
        self.is_logging_to_eventlog = False
        
        # If the option to include a date and time to the log file output is set to true, add it to the log file filename
        match add_datetime_to_log_filname:
            case True:
            # #  Assume .log as the extension if none provided, else, add user provided ext
                match self.log_file_base_ext:
                    case '':
                        self.log_file_path_and_file_handler_filename = f"{self.log_file_path}/{self.log_file_base_filename}_{datetime.now():%Y-%m-%d_%H-%M-%S}.log"
                    case _:
                        self.log_file_path_and_file_handler_filename = f"{self.log_file_path}/{self.log_file_base_filename}_{datetime.now():%Y-%m-%d_%H-%M-%S}{self.log_file_base_ext}"
            case False:
                #  Assume .log as the extension if none provided, else, add user provided ext
                match self.log_file_base_ext: 
                    case '':
                        self.log_file_path_and_file_handler_filename = f"{self.log_file_path}/{self.log_file_base_filename}.log"
                    case _:
                        self.log_file_path_and_file_handler_filename = f"{self.log_file_path}/{self.log_file_base_filename}{self.log_file_base_ext}"
        
        # Configure a 'catch all' logger device to afix multiple logging handle types to
        self.logger = logging.getLogger(logger_name)
        self.logger.setLevel(logging.DEBUG)

        # Configure a formatter for the messages ouputted by the logging handlers
        #self.formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
        self.formatter = logging.Formatter('%(asctime)s - [%(levelname)s] - %(name)s - %(lineno)d - %(message)s')

        # Create logger which outputs to the python stream (i.e. to a DOS window if using command line to test run the utility),
        # sctivate it when debug is set to true
        if debug_in_stream:
            self.stream_handler = logging.StreamHandler()
            self._configure_handler(
                self.stream_handler,
                self.formatter,
                log_level='d',
                handler_name = f'{self.handler_base_name}_STR'
            )

        # Create the log directory if it doesn't exist
        try:
            os.makedirs(log_file_path, exist_ok=True)

            # If log file directory exists or has been created, attempt to create the file handler, switching to event log if necessary
            # Attempt to configure a 20mb rotating file handler
            #   - newest file will always be named as below 
            #   - older files will have a number following the filename
            #   - the handler allows for a max of 7 consecutave backup files before deleting file 7 for good
            #   - this will continue indefinately until the script is brought to a halt 
            self.file_handler = logging.handlers.RotatingFileHandler(filename=self.log_file_path_and_file_handler_filename, maxBytes=20971520, backupCount=7)
        
        except Exception as e:
            # Rotating logging file not working, switch to windows event log
            self._switch_to_eventlog(
                'Failed to create rotating log file. Logging to Windows event log instead.\nCheck access and permissions are available to script operating functional user at: {self.log_file_path}\nError event data captured: {e}'
            )

        else:
            self._configure_handler(
                self.file_handler, 
                self.formatter, 
                log_level = self.logging_level, 
                handler_name = f'{self.handler_base_name}_RFH'
            )
               
    def _configure_handler(
            self, 
            handler: logging.Handler, 
            formatter: logging.Formatter, 
            log_level: str ='',
            handler_name: str = ''
        ) -> None:
        
        log_level_lcase = log_level.lower()

        if handler_name != '':
            handler.name = handler_name
        handler.setFormatter(formatter)
        self._select_logging_level(handler, log_level_lcase)
        self.logger.addHandler(handler)

    def _select_logging_level(
            self, 
            handler: logging.Handler, 
            log_level: str
        ) -> None:
        """
            This is used to allow the user to select a default logging level for the handlers which are configured using this class\n
               - If left blank, then, this will default to 'w' which is the 'warning level'\n
               - It can be configured to Debug which will output details of every operation completed by the class\n
               - Debug in this sense is different to the debug_in_stream setting as it applies to the handlers only\n
               - Setting the debug_in_stream level will cause all messages to be outputted to a DOS prompt if ran manually for e.g.\n
        """
        select_log_lev = {
            'd': logging.DEBUG,
            'i': logging.INFO,
            'w': logging.WARNING,
            'e': logging.ERROR,
            'c': logging.CRITICAL
            }
        set_lev =  select_log_lev[log_level]
        handler.setLevel(set_lev)

    def _check_for_and_remove_handler_of_specified_type(
            self, 
            handler_name: str,
            handler_type: str
        ) -> str:
        """
        This has been created to cover the potential need to add more log file handler types in the future\n
            If you are creating a handler type to be used within this class which could potentially fail mid-use\n 
            this method can be called to remove it from the logger at the point failure is detected\n
                - There are 2 variables:\n
                    1 - Name - this requires that the object has been created as a named handler\n
                    2 - Type - this is a string representing the abreviation of the handler type\n
                - RFH - Rotating file handler\n
                - EVT - Windows Event Log Handler\n 
        """

        # compile a list of any handlers configured within the logger
        for handler in self.logger.handlers:
            #attempt to remove handler of specified type
            try:
                # assess which type of logger needs to be removed, then, remove it:
                if isinstance(handler, logging.handlers.RotatingFileHandler) and handler.get_name() == handler_name:
                    self.logger.removeHandler(self.file_handler) 
                    self.logger.log(10, 'Successful removal of Rotating File Handler between switching logging methods')
                    self.handler_removed = str(self.file_handler)
                    return 'removal_success'
                elif isinstance(handler, logging.handlers.NTEventLogHandler) and handler.get_name() == handler_name:
                    self.logger.removeHandler(self.event_handler)
                    self.logger.log(10, 'Successful removal of Windows Event Log Handler between switching logging methods')
                    self.handler_removed = str(self.event_handler)
                    return 'removal_success'
            except Exception as e:
                self.logger.log(40, 
                                f'Error removing {self.handler_removed}, named {handler_name}, of type {handler_type}!\nThe following error was recorded: {e}'
                            )
                self.output_evtlog(
                    self.event_source, 
                    win32evtlog.EVENTLOG_ERROR_TYPE, 
                    0,
                    0,
                    [
                        f'Failed to remove {self.handler_removed} handler named {handler_name} from logger.', 
                        'Check script configuration', 
                        f'Thefollowing error was recorded: {e}'
                    ],
                    'Failed\0to\0remove\0handler\0from\0logger\0Check\0script\0config.'.encode("ascii")
                )
                return f'removal_failed'
        # if no handlers found return a status which indicates this
        self.logger.log(10, 'No handlers found at time of switching logs')
        return 'no_handlers_found'

    def _check_file_access(
            self
        ) -> None:
        """
            This is used every time an event is logged as a quick way to assess whether the log file is accessible.\n
            If it is not, it will call the switch_to_eventlog method to either drop the file handler and activeate event log logging\n  
            or, it will simply see that event log logging is enabled, and skip out of the method
        """
        try:
            with open(self.log_file_path_and_file_handler_filename, 'a'):               
                pass
            self._switch_to_file()
        except (FileNotFoundError, OSError, IOError) as e:
            self._switch_to_eventlog(str(e))


    def _switch_to_eventlog(
            self, 
            error_data: str
        ):
        """
            On issues being found with file logging, this method will be\n 
            called as an if all else fails method of capturing process data
        """

        # Check if already logging to event log
        if self.is_logging_to_eventlog:
            return

        # check for and remove any active rotating file handlers
        try:
            handler_remove_state = self._check_for_and_remove_handler_of_specified_type(f'{self.handler_base_name}_RFH', 'rfh')  
        except AttributeError:
            handler_remove_state = 'no_handlers_found'

        match handler_remove_state:
            case 'no_handlers_found' | 'removal_success':
                self.event_handler = logging.handlers.NTEventLogHandler(self.event_source, logtype='Application')
                self._configure_handler(
                    self.event_handler, 
                    self.formatter, 
                    log_level=self.logging_level,
                    handler_name = f'{self.handler_base_name}_EVT' 
                )

                # Update flag to indicate that we are now logging to the Windows Event log using the logger
                self.is_logging_to_eventlog = True

                # Log an event to the event log
                self.logger.log(30,
                        'Failed to access rolling log file. Switching to Windows Event logging.\nCheck access and permissions are available to script operating functional user at: {self.log_file_path_and_file_handler_filename}\nThefollowing error was recorded: {error_data}'
                )
                self.output_evtlog(
                    self.event_source, 
                    win32evtlog.EVENTLOG_ERROR_TYPE, 
                    0,
                    0,
                    [
                        'Failed to access rolling log file. Switching to Windows event logging.', 
                        f'Check access and permissions are available to script operating functional user at: {self.log_file_path_and_file_handler_filename}', 
                        f'Thefollowing error was recorded: {error_data}'
                    ],
                    'Failed\0to\0access\0rolling\0log\0file.\0Logging\0to\0event\0log\0instead.'.encode("ascii")
                )    
            case _:
                # At this point it is difficule to know what to do
                # As the file logging has failed and it is not possible to remove the rotating file handler from the logger
                # The script has no other choice but to end here as there's a requirement to log, yet, no possibility to log
                
                # Output a message box to screen with a simple ok button
                title = 'File copy logging failure'
                message = f'{os.path.basename(__file__)} suffered an unrecoverable error while logging activiy. Check Windows event log for details. Closing script'
                style = 0  # Ok button only
                error_popup(title, message , style)
                sys.exit(f'Failed to clear {self.file_handler.name} post failure to log')
            
    def _switch_to_file(
            self
        ) -> None:
        """
        This will be called upon to log process data to file upon\n 
        'favourable conditions' being detected for file based logging to resume
        """

        # Check if already logging to file
        if not self.is_logging_to_eventlog:
            return

        # check for and remove any active rotating file handlers
        try:
            handler_remove_state = self._check_for_and_remove_handler_of_specified_type(f'{self.handler_base_name}_EVT', 'evt')
        except AttributeError:
            handler_remove_state = 'no_handlers_found'

        # Configure the event log handler
        match handler_remove_state:
            case 'no_handlers_found' | 'removal_success':
                # Add the file handler back to the logger
                try:
                    # Attempt to configure a 20mb rotating file handler
                    #   - newest file will always be named as below 
                    #   - older files will have a number following the filename
                    #   - the handler allows for a max of 7 consecutave backup files before deleting file 7 for good
                    #   - this will continue indefinately until the script is brought to a halt 
                    self.file_handler = logging.handlers.RotatingFileHandler(
                        filename=self.log_file_path_and_file_handler_filename, 
                        maxBytes=20971520, 
                        backupCount=7
                    )
                except Exception as e:
                    # If something goes wrong configuring file logging, configure the event log handler
                    self._switch_to_eventlog(
                        f'A failed attempt was recorded when re-enabling rotating log file logging.\nIntermittent folder access may be available to {self.log_file_path_and_file_handler_filename}.\nFailed to create log file: {e}'
                    )
                else:  
                    # Make sure that an event has been logged to the event log updating that file logging resuming
                    self.output_evtlog(
                        self.event_source, 
                        win32evtlog.EVENTLOG_INFORMATION_TYPE, 
                        0,
                        0,
                        [
                            'Resuming logging to rolling log file.', 
                            f'Check {self.log_file_path_and_file_handler_filename} for further log data.'
                         ],
                        'Resuming\0logging\0to\0rolling\0log\0file.'.encode("ascii")
                    )

                    # Reconfigure the file handler
                    self._configure_handler(
                        self.file_handler, 
                        self.formatter, 
                        log_level=self.logging_level, 
                        handler_name=f'{self.handler_base_name}_RFH'
                    )       
                    
                    # Update flag to indicate file logging has resumed
                    self.is_logging_to_eventlog = False

                    # Log an event to the event log
                    self.logger.log(
                        30, 
                        f'Resuming logging to rolling log file.\nCheck {self.log_file_path_and_file_handler_filename} for further log data.'
                    )
            case _:
                # At this point it is difficult to know what to do
                # As the file logging has failed and it is not possible to remove the rotating file handler from the logger
                # The script has no other choice but to end here as there's a requirement to log, yet, no possibility to log
                
                # Output a message box to screen with a simple ok button
                title = 'File copy logging failure'
                message = f'{os.path.basename(__file__)} suffered an unrecoverable error while logging activiy. Check Windows event log for details. Closing script'
                style = 0  # Ok button only
                error_popup(title, message , style)
                sys.exit(f'Failed to clear {self.event_handler.name} post failure to log')      

    def log(
            self, 
            level: str, 
            message: str
        ) -> None:
        # Check if we need to switch to event log
        self._check_file_access()
        
        # Log the message using the current logger
        getattr(self.logger, level)(message)

        # Check if we need to switch back to file
        self._check_file_access()
    
    def output_evtlog(
            self, 
            applicationNameIn: str, 
            eventIDIn: int, 
            categoryIn: int, 
            myTypeIn: int, 
            descrIn: Iterable[str], 
            dataIn: bytes
        ) -> None:
        ph = win32api.GetCurrentProcess()
        th = win32security.OpenProcessToken(ph, win32con.TOKEN_READ)
        my_sid = win32security.GetTokenInformation(th, win32security.TokenUser)[0]

        applicationName = applicationNameIn
        eventID = eventIDIn
        category = categoryIn	
        myType = myTypeIn
        descr = descrIn
        data = dataIn

        win32evtlogutil.ReportEvent(
            applicationName, 
            eventID, 
            eventCategory=category, 
            eventType=myType, 
            strings=descr, 
            data=data, 
            sid=my_sid
        )
    
class error_popup:
    def __init__(
            self, 
            title: str, 
            text: str, 
            style: int, 
            hwnd=0
        ) -> None:
        """
        This is used to initialis the class, but, also call the popup

        Button styles:
            0 : OK
            1 : OK | Cancel
            2 : Abort | Retry | Ignore
            3 : Yes | No | Cancel
            4 : Yes | No
            5 : Retry | No 
            6 : Cancel | Try Again | Continue

        To also change icon, add these values to previous number
            16 Stop-sign icon
            32 Question-mark icon
            48 Exclamation-point icon
            64 Information-sign icon consisting of an 'i' in a circle 

        Return values
            IDABORT 3 - The Abort button was selected.
            IDCANCEL 2 - The Cancel button was selected.
            IDCONTINUE 11 - The Continue button was selected.
            IDIGNORE 5 - The Ignore button was selected.
            IDNO 7 - The No button was selected. 
            IDOK 1 - The OK button was selected.
            IDRETRY 4 - The Retry button was selected.
            IDTRYAGAIN 10 - The Try Again button was selected.
            IDYES 6 - The Yes button was selected.

        The Hwnd option is usually left as 0 - it can be used to work out the currently open window and such like
        Something like the following will display the active window name:
          - win32gui.GetWindowText(win32gui.GetForegroundWindow())
        Can use something like this to direct the message box to top of the currently active window:
           - test = ctypes.windll.user32.MessageBoxW(win32gui.GetForegroundWindow(),"Your text", "Your title", 22)
           - using the variable 'test' will allow you to capture the button clicked.
        """
        self._display_message_box(hwnd, title, text, style)
    
    def _display_message_box(self, 
                             hwnd, 
                             title, 
                             text, 
                             style):
        return ctypes.windll.user32.MessageBoxW(hwnd, text, title, style)
        


## Example as to how to utilse the logger
############################################################

class MyClass:
    def __init__(self, logger):
        self.logger = logger
        #RollingEventLogger('logs/app.log', 'myapp', 'MyApp')

    def my_method(self):
        self.logger.log('debug', 'This is a debug message from MyClass')

        # ...

        self.logger.log('error', 'An error occurred in MyClass')

        # ...
        for _ in range(100):
             self.logger.log('info', 'Info MyClass loop message')
             time.sleep(5)


## Auto trigger statement
if __name__ == '__main__':
    log_file_path = 'D:/MyFiles/temp/test'   #'logs/app.log'
    log_file_base_filename = "Testing.log"
    logger_name = 'myapp'
    log_level = 'd'
    stream_debug = True
    include_datetime = True

    logger = RollingEventLogger(
        log_file_path, 
        log_file_base_filename, 
        logger_name, 
        logging_level = log_level, 
        debug_in_stream = stream_debug,
        add_datetime_to_log_filname = include_datetime
    )

    # Log a message at the DEBUG level
    logger.log('debug', 'This is a debug message')

    # Log a message at the INFO level
    logger.log('info', 'This is an info message')

    # Log a message at the WARNING level
    logger.log('warning', 'This is a warning message')

    # Log a message at the ERROR level
    logger.log('error', 'This is an error message')

    # Log a message at the CRITICAL level
    logger.log('critical', 'This is a critical message')

    my_class = MyClass(logger)
    my_class.my_method()

    