# ti-layout-swapper
A command line script using Python (req. 3.8.3+) to allow for scheduled (Windows Task Scheduler / etc) swapping of Trade-Ideas Pro layouts.

# Description:
    A command line script using Python (req. 3.8.3+) to allow for scheduled (Windows Task Scheduler / etc) swapping of Trade-Ideas Pro layouts.
    
# Usage:
    Use Windows Task Scheduler for scheduling the layout shifts at specified time
        - Call the file using python with:
            python.exe tiswapapp.py -path "C:/my_layouts" -layout "my_layout_file.lti"
            
        - Optional Debug Mode can be activated by appending "--debug" parameters:
            python.exe tiswapapp.py -path "C:/my_layouts" -layout "my_layout_file.lti" --debug
                > This should write to a debug.txt file in the script directory.
      
# Requirements:
    - Python 3.8.3+
    - pywinauto   ( help docs: https://pywinauto.readthedocs.io/en/latest/code/pywinauto.application.html )
        - pyWin32
        - comtypes
        - six    
    
    - Install requirements with "pip install -U pywinauto"
