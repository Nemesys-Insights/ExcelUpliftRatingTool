# Uplift Evaluation Workbook Creator


## Requirements
- Windows
- Internet connection (required for first time setup).
- Python 3.13+
- A BBG benchmark dataset, complete with responses from the model to be benchmarked.
- Known number of evaluators to evaluate the responses
- Basic knowledge of python and command line interfaces.

 
# Installation and Use

## Initial Setup
The followings steps only have to be completed the first time the program is installed. 

1. Open a command prompt instance inside the project directory where this readme file is located.
    - Open the project directory in the File Explorer. Then type "cmd" in the address bar and a command prompt window will open in the correct location.
2. Type or paste `python -m venv venv` into the command prompt window and then hit the enter key. This will create a python virtual environment in the project directory.
3. Type or paste `.\venv\Scripts\activate` into the command prompt window, and then hit the enter key. This will activate the virtual environment you just created.
4. Type or paste `pip install -r requirements.txt` into the command prompt window, and then hit the enter key. This will install the python modules required for the project.

## Modifying the Workbook Generation Template
All workbooks are generated based on the eval_template.xlsx file in the template folder. Changes to this file will apply to any newly created workbooks. While formatting changes are encouraged, changes to the structure of the template are not compatible with the workbook generation program.

## Using the Workbook Generation Interface
1. Open a command prompt instance inside the project directory where this readme file is located. 
2. Type or paste `python create.py` into the command prompt window, and then hit the enter key. This will open the user interface.
3. For the Input Dataset, use the file browser to select the completed BBG benchmark dataset that you wish to distribute for evaluation.
4. Use the file browser to select your desired output folder. Ideally, this folder should be empty.  
5. Input the number of evaluators that you plan to engage.
6. Specify the number of evaluations desired per prompt-response pair in the dataset.  This value should be predetermined based on project requirements, but the minimum recommended value is 3. 
7. Click Run.
8. Wait a few moments while the workbooks generate.

In addition to the workbooks, the program generates an assignment mapping.xlsx file. This file has a tab labelled Assignments, which
details which reviewers have been assigned to each prompt-response pair in the benchmark dataset. It also has a tab labelled Summary, which indicates the number of assigned prompt-response pairs per reviewer.

## Using the Workbook Aggregation Interface
Once all of the workbooks have been completed by evaluators, this part of the tool offers an easy way of compiling all of the data into a single file in preparation for analysis. 

1. Place all of the workbooks from the evaluators into a single folder.
2. Open a command prompt instance inside the project directory where this readme file is located. 
3. Type or paste `python aggregate.py` into the command prompt window, and then hit the enter key. This will open the user interface.
4. Use the file browser to select the correct input folder, where all of the evaluator workbooks are saved.
5. Clcik Run.
6. Wait a few moments while the data is aggregated.

The program will notify you when the aggregation process is complete and inform you where the outputs of the aggregation process are.

Inside the output folder, you will find:
- A workbook containing the aggregated data named combined_clean_data.xlsx.
- A workbook listing all of the issues with the responses (e.g. missing or unexpected values) from the entire dataset.
- A folder containing all of the input workbooks, each with a new tab annotating particular issues with the responses (e.g. missing or unexpected values).



