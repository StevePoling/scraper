# scrape links from AmazonStaging.DematicTraining.com
# https://amazonstaging.dematictraining.com/login/index.php?saml=off
import argparse
from gooey import Gooey, GooeyParser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter

@Gooey 
def main():
    desc = "Scrape Moodle site for MCIDs"
    parser = GooeyParser(description=desc)
    file_help_msg = "Select the name of the Excel file you want created/overwritten."
    parser.add_argument('-f', '--filename', default='courses.xlsx', help=file_help_msg, widget="FileChooser")
    parser.add_argument('-u', '--username', default="admin", type=str, help='Username of Moodle administrator')
    parser.add_argument('-p', '--password', default="", type=str, help='Password of Moodle administrator')
    args = parser.parse_args()
    excel_filename = args.filename
    username = args.username
    password = args.password

    login_url = 'https://amazonstaging.dematictraining.com/login/index.php?saml=off'
    url = 'https://amazonstaging.dematictraining.com/course/index.php'
    driver = webdriver.Chrome()

    print("logging in")
    driver.get(login_url)
    driver.find_element_by_id("username").send_keys(username)
    driver.find_element_by_id("password").send_keys(password)
    driver.find_element_by_id("loginbtn").click()
    driver.get(url)
    driver.find_element_by_class_name("collapseexpand").click()

    print("creating workbook")
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 50)
    worksheet.write('A1', 'Course Name')
    worksheet.write('B1', 'Link')

    try:
        print("expanding categories")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "9410 BOR Overview German"))
        )

        course_divs = driver.find_elements_by_class_name("coursename")
        i = 1
        for course in course_divs:
            name = course.text
            href = course.find_element_by_xpath('a').get_attribute('href')
            worksheet.write(i,0,name)
            worksheet.write(i,1,href)
            i = i + 1

        categories = driver.find_elements_by_css_selector("h4.categoryname a")
        category_urls = []
        for category in categories:
            category_url = category.get_attribute('href')
            category_urls.append(category_url)

        print("expanding sub-categories")
        for category_url in category_urls:
            driver.get(category_url)
            course_divs = driver.find_elements_by_class_name("coursename")
            for course in course_divs:
                name = course.text
                href = course.find_element_by_xpath('a').get_attribute('href')
                worksheet.write(i,0,name)
                worksheet.write(i,1,href)
                i = i + 1

        workbook.close()
        print("success")
    finally:
        driver.quit()

if __name__ == '__main__':
    main()