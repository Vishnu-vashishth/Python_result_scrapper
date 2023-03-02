import selenium 
import time
from selenium import webdriver 
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
#list of roll numbers
roll_numbers = range(20020004001,20020004120)

# create an empty list to store student information
students = []

# open chrome browser


# base url of website
base_url = "https://jcboseustymca.co.in/Forms/Student/ResultStudents.aspx"

# semester you want to select
semester = '03'

#iterate through roll numbers
driver = webdriver.Chrome()
for roll_number in roll_numbers:
    # go to the website
    driver.get("https://jcboseustymca.co.in/Forms/Student/ResultStudents.aspx")

    # find the input field for roll number
    roll_no_input = driver.find_element('id','txtRollNo')

    # enter roll number
    roll_no_input.send_keys(roll_number)
    
    # find the semester dropdown element
    semester_dropdown = Select(driver.find_element('id','ddlSem'))
    # select the desired semester
    semester_dropdown.select_by_value(semester)

    # find the submit button
    submit_button = driver.find_element('id','btnResult')

    # click the submit button

    submit_button.click()
    
    try :
         elem= driver.find_element('id','lblMessage')
    except (ValueError, TypeError):
        print("")
        continue



    # wait for the result to appear
    # time.sleep(2)
    if elem.text == "No Record Found. Please enter correct roll no or result may not be validate.":
        continue
    driver.get('https://jcboseustymca.co.in/Forms/Student/PrintReportCardNew.aspx')

    # extract student information
    name = driver.find_element('id','lblname').text
    roll_number = driver.find_element('id','lblRollNo').text
    result = driver.find_element('id','lblResult').text

    # add student information to list
    students.append({
        'name': name,
        'roll_number': roll_number,
        'result': result
    })
    # print(students)
    # time.sleep(5)
# create a DataFrame from the student information list


    # print("hello")
df = pd.DataFrame(students)
# print(df)
# export DataFrame to Excel
df.to_excel('students.xlsx', index="true")

# close the browser
driver.quit()
