from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

workbook = openpyxl.Workbook()
sheet = workbook.active

driver = webdriver.Chrome()

# Load the webpage
driver.get('https://irp.nih.gov/our-research/principal-investigators/name')
for elements in driver.find_elements(By.CLASS_NAME, "teaserlist__item"):
    for el in elements.find_elements(By.CLASS_NAME, 'pilist__item'):
        first, last, email, company, position, website, phone, street, address2, town, state, zipCode = ' '*12
        try:
            index = el.find_element(By.TAG_NAME, 'a')
            name = index.text
            name_parts = name.split(',')
            first = name_parts[1]  # column 1
            last = ",".join([name_parts[0]] + name_parts[2:])  # column 2
            link = index.get_attribute('href')
            driver.get(link)
            profile = driver.find_element(By.CLASS_NAME, 'profile')
            positionDetails = profile.find_element(By.CLASS_NAME, 'profile__content').find_element(By.CLASS_NAME,
                                                                                                   'profile__content-group')
            position = positionDetails.find_element(By.TAG_NAME, 'h2').text + ', ' + positionDetails.find_element(
                By.TAG_NAME, 'p').text  # column 7
            company = positionDetails.find_element(By.TAG_NAME, 'strong').text  # column 6
            website = profile.find_element(By.CLASS_NAME, 'profile__content').find_elements(By.CLASS_NAME,
                                                                                        'profile__content-group')[
            1].find_element(By.TAG_NAME, 'a').get_attribute('href')  # column 9

            details = profile.find_element(By.CLASS_NAME, 'profile__sidebar-wrapper')

            address = details.find_elements(By.TAG_NAME, 'p')[0].text
            street = address.split('\n')[0]
            location = address.split('\n')[-1]
            if len(location.split()) == 3:
                town = location.split()[0].replace(',', '')  # column 13
                state = location.split()[1]  # column 14
                zipCode = location.split()[2]  # column 15
            else:
                zipCode = location

            address2 = address.replace(street, '').replace(location, '').replace('\n', '')  # column 12

            phone = details.find_elements(By.TAG_NAME, 'p')[1].text  # column 10
            email = details.find_elements(By.TAG_NAME, 'p')[2].find_element(By.TAG_NAME, 'a').get_attribute(
                'href')  # column 3
        except:
            print('')
        row = [first, last, email, '', '', company, position, '', website, phone, street, address2, town, state,
               zipCode, '']
        sheet.append(row)
        print(row)

        driver.back()
workbook.save("output.xlsx")
driver.quit()
