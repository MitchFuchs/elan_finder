# elan_finder
This GUI provides a search engine for ELAN files (.eaf EAF Annotation Format) in a specified directory.
Based on your ELAN annotations, it allows direct access to your ELAN through the GUI and can save you valuable time on researching your ELAN database.

# demo 
https://user-images.githubusercontent.com/73831423/144033054-50075f25-97c6-48c4-98fb-c16542d3326b.mp4

# use case example 
You possess a large database of ELAN-annotated videos and need to access / view your annotations without having to open each file individually. 

# instructions
1) After starting the application, choose the directory in which all your ELAN files are stored. 
2) You can then filter the linguistic types and their related annotation values. 
3) The table shows a list of all matching results in your database.
4) Double-click any row to display the video of your selected annotation. 

# dependencies
Using our env freeze: 
- `pip install -r requirements.txt`

Or manually install:
- `[pip or conda] install opencv`
- `[pip or conda] install pillow`
- `[pip or conda] install pandas`
- `[pip or conda] install xlsxwriter`

# dev

To freeze the venv:
- `pip list --format=freeze > requirements.txt`

To build executable on mac (app witll be in build folder):
- `pyinstaller --onefile --hidden-import cmath main.py`

# contact
Feel free to reach out at michael@neptune-consulting.ch
Cheers, 
Mitch. 
