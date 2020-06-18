# -*- coding: utf-8 -*-
"""
Created on Tue Jun 16 11:12:57 2020

@author: Kyle Kimsey for BBT Community @ bearbulltraders.com
@webpage: k-sqrd.com/me

Description:
    Simple script to take in a layout filename, grab the focus of TradeIdeas, and
    then load that layout. 
    
Usage:
    Use Windows Task Scheduler for scheduling the layout shifts at specified time
        - Call the file using python with:
            python.exe tiswapapp.py -path "C:/my_layouts" -layout "my_layout_file.lti"
            
        - Optional Debug Mode can be activated by appending "--debug" parameters:
            python.exe tiswapapp.py -path "C:/my_layouts" -layout "my_layout_file.lti" --debug
                > This should write to a debug.txt file in the script directory.
      
Requirements:
    - Python 3.8.3+
    - pywinauto   ( help docs: https://pywinauto.readthedocs.io/en/latest/code/pywinauto.application.html )
        - pyWin32
        - comtypes
        - six    
    
    Install requirements with "pip install -U pywinauto"
    
"""

from pywinauto.application import Application
from datetime import datetime
import win32gui
import win32com
import time
import os
import argparse


# CONFIGURATION
TI_PATH = "C:\\Program Files\\Trade-Ideas\\Trade-Ideas Pro AI\\"
TI_EXE = "TIPro.exe"
TI_PATHEXE = f"{TI_PATH}{TI_EXE}"
#LAYOUT_PATH = "C:\Temp"
#LAYOUT_NAME = "BBT_SCANNER_NUMBER2.LTI"

DEFAULT_TI_USER = os.path.join(os.path.expanduser('~'), 'Documents\TradeIdeasPro')
DEBUG = True

# cmd prompt -> start /b /d "C:\Program Files\Trade-Ideas\Trade-Ideas Pro AI\" TIPro.exe -NorthAmerica.xml
TI_SHELLPATH = f'cmd /c start /b /d "{TI_PATH}" {TI_EXE} -NorthAmerica.xml'


"""
WARNING: Program begins here, only change if you know what you are doing.

APPLICATION:
    
"""

def pdebug(msg):
    # Quick debug print out if the global DEBUG is set to True, otherwise it does nothing.
    if DEBUG:
        print(f'-- DEBUG: {msg}')
        log_to_file(msg)
  

def log_to_file(msg, category='debug', logfile='debug.txt'):
    # Quick debug log to file function.
    with open(logfile, "a") as file_object:
        #Construct the log item
        #stack = inspect.stack()
        #function = get_stack_info(stack)
        timestamp = datetime.now().strftime("%m/%d/%Y, %H:%M:%S")
        file_object.write(f"{timestamp} - {category.upper()} - {msg}\n")
                     

def get_window_focus_retry(wnd, wait=.3, retries=30):
    """
    Parameters
    ----------
    wnd : String
        Text to search within the window titles for, case insensitive. 
    wait : FLOAT, optional
        Time to wait between retries in seconds. The default is .3.
    retries : INT, optional
        Number of attempts to make before returning False. The default is 30.

    Returns
    -------
    bool
        If the program is found and in focus, it will return TRUE. Otherwise
        if the the amount of retries is exhausted, it will return FALSE.

    """    
    for n in range(retries):
        _window = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        
        if wnd.lower() not in _window.lower():
            time.sleep(wait)
        else:
            return True     
    return False





if __name__ == "__main__":
         
    # Create the parser
    parser = argparse.ArgumentParser(prog="Swap TradeIdeas Layouts", 
                                     usage="script.py -path 'C:\path-location\'  -layout 'my_layout.lti'",
                                     description="Change the loaded layout for Trade-Ideas Pro \
                                     by passing in an [optional] path and layout filename.")
    
    # Add the arguments
    parser.add_argument("-path",
                            metavar = "path",
                            type=str,
                            help="The directory path to where the to-be-loaded Layout is located. \
                                e.g. C:\my_layout_shenanigans\secret_sauce",
                            default="")
    
    parser.add_argument("-layout",
                            metavar = "layout",
                            type=str,
                            required=True,
                            help="The the name of the layout file to load, e.g. MY_TI_LAYOUT.LTI",
                            default="")
    
    parser.add_argument("--debug",
                            metavar = "Debug Mode",
                            type=bool,
                            nargs='?',
                            const=True,
                            help="TRUE if present: Include extra console print outs during processing.",
                            default=False)
    
    
    # Execute the parse_args() method
    args = parser.parse_args()
       
    DEBUG = args.debug
    LAYOUT_PATH = os.path.normcase(args.path)
    LAYOUT_NAME = args.layout
    
    # HELPERS
    _FN = os.path.join(LAYOUT_PATH, LAYOUT_NAME)
    
    # Quick sanity check on the file.
    if not os.path.exists(_FN):
        pdebug('File not found at passed in location. Attempting to locate using the default TI layout folder.')
        
        _FN = os.path.join(DEFAULT_TI_USER, LAYOUT_NAME)
        
        if not os.path.exists(_FN):
            print('\n File was not found in any known location, please check the parameters you submitted and try again.')
            print(f"\n Location: {_FN} \n")
            pdebug(f'File not found, please check the parameters submitted. Parameters: {_FN}')
            quit()
            
            
    print(f"\n Location: {_FN} \n")
                  
    
    try:
        app = Application(backend="uia").connect(path=TI_PATHEXE, title="TradeIdeasPro")
        pdebug("Already running, attaching to the process ID.")
        
    except:
        pdebug("Application not found, trying to start it.")
        os.system(TI_SHELLPATH)
        time.sleep(5)
        app = Application(backend="uia").connect(path=TI_PATHEXE, title="TradeIdeasPro")
           
    # Grab a shell for keyboard entry
    kbshell = win32com.client.Dispatch("WScript.Shell")
    
    # Check if the program is running.
    if app.is_process_running():
        pdebug("Confirming Trade-Ideas is running.")
           
        # Grab the right program, and select File -> Load Layout
        dlg = app.window(name_re=".*Trade-Ideas.*")
        dlg.set_focus().type_keys("%FL")
         
        # Doing this manually seems to be a lot quicker than with the PyWinAuto hooks.  
        if get_window_focus_retry('Open'):             
            kbshell.SendKeys(_FN, 0)
            time.sleep(.2)
            kbshell.SendKeys("~")
        else:
            pdebug('Can not capture the open window, doing nothing.')       
        
    # Close up shop, we're done here. 
    quit()
