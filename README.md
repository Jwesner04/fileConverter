# fileConverter
Custom file parsing that reads and writes specific sections of files and outputs to a summary word document

#--------------------------------------------------------------------------------------#
# CLONE REPOSITORY ON LOCAL COMPUTER (LINUX)
#--------------------------------------------------------------------------------------#
    - Install git on computer if you don't have it using command line 
            -- 'sudo apt-get install git-all'
    - Go to a specific folder within command line that you want
      the repository to be using the command:
            -- 'cd /myFolder'
    - Once you are in your folder, type:
            -- 'git clone https://Jwesner04@bitbucket.org/Jwesner04/fileconverter.git' 


#--------------------------------------------------------------------------------------#
# RUN .EXE FILE IN COMMAND LINE
#--------------------------------------------------------------------------------------#
   - Once the repository is cloned on your computer 'cd' into the repository folder
   - Then run the command via command line './fileConverter'

#--------------------------------------------------------------------------------------#
# ALTERNATIVE
#--------------------------------------------------------------------------------------#
   - Assuming you have python already installed on your computer, otherwise install
     python.
   - Get dependencies on linux (tested on Ubuntu) running with python 2.7
        - open command line...into default directory is fine and run the following
          commands:
        - sudo apt-get -y install python-pip
	- sudo pip install lxml
	- sudo pip install python-docx
	- sudo pip install docxtpl
	- sudo apt-get install python-qt4

   - Make file executable on local machine using .py file:
	- cd to folder the repository is in
	- 'chmod +x fileToWordConverter.py' --in command line
	- './fileToWordConverter.py' --run this command and file will be executable