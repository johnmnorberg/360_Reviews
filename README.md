# 360_Reviews
This was my first Python project to develop an automated means to collate and distribute 360 reviews. 360 reviews allow employees to answer questions about their coworkers and themselves. Answers are collected, and can be anonymously given back to each employee by their supervisor. The process allows individuals to hear feedback from their peers that their peers may not be comfortable giving. It also allows for some self reflection. The goal of 360s is to provide constructive criticism and praise.

I wrote this in early 2019, and it was successfully deployed by my employer. The code has been modified slightly to make it generic for re-use. The goal of the project was to find a free and time-saving method to collect, collate, and distribute answers from about 15-20 people.

When it was deployed, I used a two folder framework, as follows:

* Template_folder/
    * Create_360s_Template.exe
    * Employees.xlsx
    * Questions.xlsx
    * 360_Reviews_Template.xlsx (Created by the .exe file)

* Compiling_folder/
    * 360_Reviews.exe
    * Employees.xlsx (must be a duplicate from above)
    * Questions.xlsx (must be a duplicate from above)
    * Employee_A.xlsx
    * Employee_B.xlsx
    * Employee_C.xlsx
    * ...
    
In this repository, I am providing the .py files used to create the .exe files as well as the relevant employee and questions .xlsx files. I converted the .py files to .exe as Python was not guaranteed to be installed on the computers that would be executing the process as Python was not needed by any other employee within the company.

The Create_360s_Template.py script simply creates a form for individuals to electroically copmlete. The completed forms were then emailed back to myself for compiling purposes. I was entrusted to not read any of the submissions as I (being a supervisor) would only have access to feedback about my subordinates.

The 360_Reviews.py script took in every submission, parsed out the data, and created four separate documents corresponding with each employee type (line-level, supervisor, administration, and director). For example, all feedback about line-level employees were collected, and organized into one file. Each piece of feedback had the author's name associated with it. Once completed, each form was automatically emailed to the associated supervisor's email address, then subsequently deleted to keep data in the compiling folder segregated as much as possible. It was up to the supervisor to save the collated feedback in their own appropriate folder. Individual submissions (.xlsx) files were then archived. For example, feedback about administrative employees was collated, emailed to the director, then the collated file was deleted.
    
