# -*- coding: utf-8 -*-
"""
Created on Tue Jun 16 11:12:57 2020

@author: Kyle Kimsey for BBT Community @ bearbulltraders.com
@webpage: k-sqrd.com/me


IMPORTANT NOTE:
    - Requires Python 3.8.3+ to function.
    - This was merged with an EXE version to make deployment easier, could spend time rewriting this to
        clean up the code at a later date.


Description:
    Simple script to take in a layout filename, grab the focus of TradeIdeas, and
    then load that layout. 
    
Usage:
    Use Windows Task Scheduler for scheduling the layout shifts at specified time
        - Call the file using python with:
            python.exe tilayoutswap.py -path "C:/my_layouts" -layout "my_layout_file.lti"
            
        - Optional Debug Mode can be activated by appending "--debug" parameters:
            python.exe tilayoutswap.py -path "C:/my_layouts" -layout "my_layout_file.lti" --debug
                > This should write to a debug.txt file in the script directory.
    EXE Version functions the same, just replace the python.exe script.py portion with the .exe.
        e.g. tilayoutswap.exe -path "C:/my_layouts" -layout "my_layout_file.lti"
      
Requirements (for script version):
    - Python 3.8.3+
    - pywinauto   ( help docs: https://pywinauto.readthedocs.io/en/latest/code/pywinauto.application.html )
        - pyWin32
        - comtypes
        - six    
    
    Install requirements with "pip install -U pywinauto"
    
Requirements (for EXE version):
    - Should run self-sufficient.
    
Note:
    TI Run Syntax is: 
        cmd prompt -> start /b /d "C:\Program Files\Trade-Ideas\Trade-Ideas Pro AI\" TIPro.exe -NorthAmerica.xml
"""

from pywinauto.application import Application
from datetime import datetime
import win32gui
import win32com
import time
import os
import argparse
import configparser


# CONFIGURATION 
config = configparser.ConfigParser()
defaultcfg = configparser.ConfigParser()

defaultcfg.add_section('user')
defaultcfg.add_section('tradeideas')

defaultcfg['user']['layout_directory'] = "C:\Temp"
defaultcfg['user']['debug_mode'] = "False"
defaultcfg['tradeideas']['ti_path'] = "C:\\Program Files\\Trade-Ideas\\Trade-Ideas Pro AI\\"
defaultcfg['tradeideas']['ti_exe'] = "TIPro.exe"
defaultcfg['tradeideas']['ti_default_user_directory'] = os.path.join(os.path.expanduser('~'), 'Documents\TradeIdeasPro')
defaultcfg['tradeideas']['ti_shellpath'] = f'cmd /c start /b /d "[TIPATH]" [TIEXE] -NorthAmerica.xml'

config.read('config.ini')

## Config Helpers

def merge_configs(config, defaults):
    # Merges the default configs with the user given configs.
    mergedcfg = defaults
    for section in config.sections():
        for name, value in config.items(section):
            mergedcfg.set(section, name, value)
    return mergedcfg
            
def config_shellpath(cfg):
    # Take in the active config and replace the placeholders with values post load.
    _shellpath = cfg['tradeideas']['ti_shellpath']
    return _shellpath.replace("[TIPATH]", cfg['tradeideas']['ti_path']).replace("[TIEXE]", cfg['tradeideas']['ti_exe'])
    
cfg = merge_configs(config, defaultcfg)

with open('config.ini', 'w') as cfgfile:
    cfg.write(cfgfile)

# Config Variables for Config [Legacy from old version]
TI_PATH = cfg['tradeideas']['ti_path']
TI_EXE = cfg['tradeideas']['ti_exe']
TI_PATHEXE = f"{TI_PATH}{TI_EXE}"
DEFAULT_TI_USER = cfg['tradeideas']['ti_default_user_directory']
TI_SHELLPATH = cfg['tradeideas']['ti_shellpath']

LAYOUT_PATH = cfg['user']['layout_directory']
DEBUG = cfg['user']['debug_mode']

"""
WARNING: Program begins here, only change if you know what you are doing.

APPLICATION:
    
"""

def pdebug(msg):
    # Quick debug print out if the global DEBUG is set to True, otherwise it does nothing.
   
    if DEBUG:
        print(f'-- DEBUG: {msg}')
        log_to_file(msg)
  
def quit(code):
    sys.exit(code)

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
                                     usage="program.exe -path 'C:\path-location\'  -layout 'my_layout.lti'",
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

    if args.debug:
        DEBUG = args.debug
    
    if args.path:
        LAYOUT_PATH = args.path
               
    LAYOUT_PATH = os.path.normcase(LAYOUT_PATH)
    LAYOUT_NAME = args.layout
    LAYOUT_NAME = LAYOUT_NAME.replace("'", "").replace('"', '')
    
    # HELPERS
    _FN = os.path.join(LAYOUT_PATH, LAYOUT_NAME)
    
    # Quick sanity check on the file.
    if not os.path.exists(_FN):
        pdebug(f'File not found at passed in location --> {_FN}. Attempting to locate using the default TI layout folder.')
        
        _FN = os.path.join(DEFAULT_TI_USER, LAYOUT_NAME)
        
        if not os.path.exists(_FN):
            print('\n File was not found in any known location, please check the parameters you submitted and try again.')
            print(f"\n Location: {_FN} \n")
            pdebug(f'File not found, please check the parameters submitted. Parameters: {_FN}')
            quit(1)
            
            
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
        pdebug("Confirming Trade-Ideas is running ... ")
           
        # Grab the right program, and select File -> Load Layout
        dlg = app.window(name_re=".*Trade-Ideas.*")
        dlg.set_focus().type_keys("%FL")
         
        # Doing this manually seems to be a lot quicker than with the PyWinAuto hooks.  
        if get_window_focus_retry('Open'):             
            kbshell.SendKeys(_FN, 0)
            time.sleep(.2)
            kbshell.SendKeys("~")
        else:
            pdebug('Can not capture the open window, closing.')
            quit(1)       
        
    # Close up shop, we're done here. 
    quit(0)