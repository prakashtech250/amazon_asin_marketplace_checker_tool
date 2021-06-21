# amazon_asin_marketplace_checker_tool
Amazon asin marktetplace checker tool

Steps:

    1. Install python
        I assumed that you have install python on your operating system.
        To check if python is installed properly, open terminal/command_prompt and type command:
            > python --version or python3 --version
                Python 3.8.5
                
    2. Extract the zip file that I provide. (suppose the zip file is named as script.zip)
        Open terminal/command prompt and change directory inside the folder that you extracted script.zip
            > cd path_where_you_extract_script/script

    3. Install pip (pip is package manager in python to install all the modules that we are using in our project)
            > python get-pip.py or python3 get-pip.py
    
    4. Install all the modules listed on requirements.txt by typing command. (out of 4 commnads, one should work)
            > pip install -r requirements.txt or
            > pip3 install -r requirements.txt or 
            > python -m pip install -r requirements.txt or
            > python3 -m pip install -r requirements.txt

    5. Inside folder you can see a file (asin.txt). you can update list of asin in that file.

    6. Run the tools.
            > python main.py or 
            > python3 main.py

    