# Hertie School Course Evaluation Desktop App

![image](./app.png)

This is an app to support Hertie School's Curricular Affairs team to automate part of the course evaluation procedure. If you are part of the team, there are Windows, Mac and Linux-compatible versions of the app available in the releases branch of this repo: [Download Desktop App](https://github.com/hertie-data-science-lab/hertie-course-eval/releases)

Prototype for testing is available on Heroku at: https://hertie-course-eval.herokuapp.com/

However, all actuall reporting and data processing should be used through the Desktop app to guarantee data protection and privacy compliance. 

--- 

### For whoever that maintains this in the future, some notes: 

* MPP/MIA and EMPA courses are different in their question forms and should be processed differently; 
* Edge cases to keep track of: some courses have more than one instructors and their scores need to be averaged;
* Raw files are all in .xls format;
* Float formatting is wrong and needs to be converted from "," to "." for processing;
* This repo is public as there's no sensitive information and for the team to access the download files. Please do not upload any test files or actual reports into this repo. 
