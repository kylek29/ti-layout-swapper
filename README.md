# ti-layout-swapper


# Description
A command line script using Python (req. 3.8.3+) to allow for scheduled (Windows Task Scheduler / etc) swapping of Trade-Ideas Pro layouts.
    
# Usage
Use Windows Task Scheduler for scheduling the layout shifts at specified time
* Call the file using python with:  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; python.exe tilayoutswap.py -path "C:/my_layouts" -layout "my_layout_file.lti"
            
* Optional Debug Mode can be activated by appending "--debug" parameters:  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; python.exe tilayoutswap.py -path "C:/my_layouts" -layout "my_layout_file.lti" --debug  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; * *This should write to a debug.txt file in the script directory.*

* For the EXE version you can use the same arguments, just call 'tilayoutswap.exe' instead of 'python.exe tilayoutswap.py'
      
# Config
- Upon first run of the .exe, a default config.ini file will be created in the .exe's directory, options are:
    - [user] layout_directory :: This is the default directory we should look into for layout files. Defaults to C:\Temp
    - [user] debug_mode :: Turns the extended debug logging on or off. Defaults to False (other option is True)
    - [tradeideas] ti_path :: The location for the TradeIdeas EXE. Default is "C:\Program Files\Trade-Ideas\Trade-Ideas Pro AI\"
    - [tradeideas] ti_exe :: The name for the Trade Ideas EXE. Default is "TIPro.exe" -- Change this if you have a different version.
    - [tradeideas] ti_default_user_directory :: The location of where TradeIdeas saves layouts by default. Defaults to: "C:\Users\{user}\Documents\TradeIdeasPro"
    - [tradeideas] ti_shellpath :: Parameters for calling TI from the cmd line (shell), you likely don't need to change this. Uses ti_path and ti_exe from above.

      
# Requirements
* **Python 3.8.3+**
- pywinauto   ( help docs: https://pywinauto.readthedocs.io/en/latest/code/pywinauto.application.html )
    - pyWin32
    - comtypes
    - six     
    - Install requirements with "pip install -U pywinauto"
    
    
# Build EXE Instructions
To build an EXE you need to use PyInstaller. You may want to setup a separate vanilla virtual environment, if you use Anaconda 3 you can do this like:

* conda create --no-default-packages -n 'pybase38tiswap' python=3.8.3
* conda activate pybase38tiswap

* pip install -U pywinauto
* pip install pyinstaller
* pip install argparse
* pip install configparser

Then run the script from the tilayoutswap.py scripts directory with:

* pyinstaller --onefile tilayoutswap.py --icon tilayoutswap.ico

NOTE:
- The above is the recommended method as building directly from a base Anaconda environment can include every module that is installed and the resulting .EXE maybe hundreds of megs big. Doing a bare environment allows you to significantly reduce the size (for me it was reduced from 300mb to 14mb).
