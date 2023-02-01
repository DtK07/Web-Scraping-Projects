from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
import openpyxl
from selenium.webdriver.chrome.options import Options

read_df = pd.read_excel('Switch Up.xlsx', sheet_name='Switch Up')
switch_up_url = read_df['Switch Up URL']
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
course_list = []
for url in switch_up_url:
    print(url)
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url)
    time.sleep(3)
    #bootcamp names
    bootcamp_name = driver.find_element(By.XPATH, '//div[@class="bootcamp-header__name"]/h1').text
    #bootcamp locations
    location_list_element = driver.find_elements(By.XPATH,'//bootcamp-courses[@class = "bootcamp-locations uid-1cf0696c568338fb"]/span[1]/span')
    locations = []
    for location in location_list_element:
        locations.append(location.text)
    rating = driver.find_element(By.XPATH, '//div[@class="overall-rating__rating"]/h4').text
    if driver.find_elements(By.XPATH, '//bootcamp-courses//div[@class="mdc-layout-grid__cell mdc-layout-grid__cell--span-12-desktop"]'):
        Isnextdisabled = False
        #course menu expansion
        while not Isnextdisabled:
            try:
                courses_list_expand = driver.find_elements(By.XPATH,'//div[@class="courses"]/div[@class = "course section-spacing__bottom--1"]')
                for exp_button in courses_list_expand:
                    exp_button.find_element(By.XPATH,'//div[@class = "courses-dropdown toggle-btn toggle-bound"]/button[@class = "btn__fab btn__fab--closed expand"]').click()
                course_tabs = driver.find_elements(By.XPATH, '//div[@class="course section-spacing__bottom--1"]')
                for course in course_tabs:
                    course_name = course.text.split("\n")[0]
                    if "Cost:" in course.text:
                        course_cost = course.text[course.text.find("Cost:") + len("Cost:"):].split("\n")[0].strip()
                    else:
                        course_cost = None
                    if "Duration:" in course.text:
                        course_duration = course.text[course.text.find("Duration:") + len("Duration:"):].split("\n")[0].strip()
                    else:
                        course_duration = None
                    if "Subjects:" in course.text:
                        course_subjects = course.text.split("\n")[-1]
                    else:
                        course_subjects = None
                    if "Course Description:" and "Subjects:" in course.text:
                        course_desc = course.text[course.text.find("Course Description:") + len("Course Description") + 2: course.text.find("Subjects:")].strip()
                    elif "Course Description" in course.text:
                        course_desc = course.text[course.text.find("Course Description:") + len("Course Description") + 2:].strip()
                    elif "Course Desription" not in course.text:
                        course_desc = None
                    courses = {"Bootcamp Name": bootcamp_name,
                               "Bootcamp Locaton": locations,
                               "Bootcamp Rating": rating,
                               "Course Name": course_name,
                               "Course Cost": course_cost,
                               "Course Duration": course_duration,
                               "Course Desc": course_desc,
                               "Course Subjects": course_subjects}
                    course_list.append(courses)
            except:
                course_tabs = driver.find_elements(By.XPATH, '//div[@class="course section-spacing__bottom--1"]')
                for course in course_tabs:
                    course_name = course.text.split("\n")[0]
                    if "Cost:" in course.text:
                        course_cost = course.text[course.text.find("Cost:") + len("Cost:"):].split("\n")[0].strip()
                    else:
                        course_cost = None
                    if "Duration:" in course.text:
                        course_duration = course.text[course.text.find("Duration:") + len("Duration:"):].split("\n")[0].strip()
                    else:
                        course_duration = None
                    course_subjects = None
                    course_desc = None
                    courses = {"Bootcamp Name": bootcamp_name,
                               "Bootcamp Locaton": locations,
                               "Bootcamp Rating": rating,
                               "Course Name": course_name,
                               "Course Cost": course_cost,
                               "Course Duration": course_duration,
                               "Course Desc": course_desc,
                               "Course Subjects": course_subjects}
                    course_list.append(courses)
            try:
                driver.find_element(By.XPATH,'//bootcamp-courses/span/div[@class="text--centered vert-space__top"]/div[@class = "mdc-layout-grid__cell mdc-layout-grid__cell--span-12-desktop"]/nav[@class="pagination text--centered"]/span/a[contains( text( ),"Next")]').click()
                time.sleep(3)
            except:
                Isnextdisabled = True
        driver.close()
    else:
        try:
            courses_list_expand = driver.find_elements(By.XPATH,'//div[@class="courses"]/div[@class = "course section-spacing__bottom--1"]')
            for exp_button in courses_list_expand:
                exp_button.find_element(By.XPATH,'//div[@class = "courses-dropdown toggle-btn toggle-bound"]/button[@class = "btn__fab btn__fab--closed expand"]').click()
            course_tabs = driver.find_elements(By.XPATH, '//div[@class="course section-spacing__bottom--1"]')
            for course in course_tabs:
                course_name = course.text.split("\n")[0]
                if "Cost:" in course.text:
                    course_cost = course.text[course.text.find("Cost:") + len("Cost:"):].split("\n")[0].strip()
                else:
                    course_cost = None
                if "Duration:" in course.text:
                    course_duration = course.text[course.text.find("Duration:") + len("Duration:"):].split("\n")[
                        0].strip()
                else:
                    course_duration = None
                if "Subjects:" in course.text:
                    course_subjects = course.text.split("\n")[-1]
                else:
                    course_subjects = None
                if "Course Description:" and "Subjects:" in course.text:
                    course_desc = course.text[course.text.find("Course Description:") + len(
                        "Course Description") + 2: course.text.find("Subjects:")].strip()
                elif "Course Description" in course.text:
                    course_desc = course.text[
                                  course.text.find("Course Description:") + len("Course Description") + 2:].strip()
                elif "Course Desription" not in course.text:
                    course_desc = None
                courses = {"Bootcamp Name": bootcamp_name,
                           "Bootcamp Locaton": locations,
                           "Bootcamp Rating": rating,
                           "Course Name": course_name,
                           "Course Cost": course_cost,
                           "Course Duration": course_duration,
                           "Course Desc": course_desc,
                           "Course Subjects": course_subjects}
                course_list.append(courses)
        except:
            course_tabs = driver.find_elements(By.XPATH, '//div[@class="course section-spacing__bottom--1"]')
            for course in course_tabs:
                course_name = course.text.split("\n")[0]
                if "Cost:" in course.text:
                    course_cost = course.text[course.text.find("Cost:") + len("Cost:"):].split("\n")[0].strip()
                else:
                    course_cost = None
                if "Duration:" in course.text:
                    course_duration = course.text[course.text.find("Duration:") + len("Duration:"):].split("\n")[
                        0].strip()
                else:
                    course_duration = None
                course_subjects = None
                course_desc = None
                courses = {"Bootcamp Name": bootcamp_name,
                           "Bootcamp Locaton": locations,
                           "Bootcamp Rating": rating,
                           "Course Name": course_name,
                           "Course Cost": course_cost,
                           "Course Duration": course_duration,
                           "Course Desc": course_desc,
                           "Course Subjects": course_subjects}
                course_list.append(courses)
        driver.close()
df = pd.DataFrame(course_list)
with pd.ExcelWriter('courses.xlsx',mode='a', if_sheet_exists = 'overlay') as writer:
    df.to_excel(writer, sheet_name='Courses1', index = False)






