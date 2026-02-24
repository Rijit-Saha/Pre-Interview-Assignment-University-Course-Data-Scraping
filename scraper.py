import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook


def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=chrome_options)
    return driver



universities = [
    {"university_id": "U001", "university_name": "University of Helsinki", "country": "Finland", "city": "Helsinki", "website": "https://www.helsinki.fi/en"},
    {"university_id": "U002", "university_name": "University of Edinburgh", "country": "United Kingdom", "city": "Edinburgh", "website": "https://www.ed.ac.uk"},
    {"university_id": "U003", "university_name": "University of Amsterdam", "country": "Netherlands", "city": "Amsterdam", "website": "https://www.uva.nl"},
    {"university_id": "U004", "university_name": "University of Sydney", "country": "Australia", "city": "Sydney", "website": "https://www.sydney.edu.au"},
    {"university_id": "U005", "university_name": "University of Alberta", "country": "Canada", "city": "Edmonton", "website": "https://www.ualberta.ca"},
]



def scrape_courses(driver, university):
    courses = []
    driver.get(university["website"])
    time.sleep(3)

    elements = driver.find_elements(By.TAG_NAME, "a")
    course_counter = 1

    for element in elements:
        text = element.text.strip()
        if len(text) > 15 and course_counter <= 5:
            course = {
                "course_id": f"C{university['university_id'][1:]}{course_counter}",
                "university_id": university["university_id"],
                "course_name": text,
                "level": "Not Specified",
                "discipline": "General",
                "duration": "Not Available",
                "fees": "Not Available",
                "eligibility": "Not Available"
            }
            courses.append(course)
            course_counter += 1

    return courses



def main():
    driver = setup_driver()
    all_courses = []

    for uni in universities:
        print(f"Scraping {uni['university_name']}...")
        try:
            uni_courses = scrape_courses(driver, uni)
            all_courses.extend(uni_courses)
        except Exception as e:
            print(f"Error scraping {uni['university_name']}: {e}")

    driver.quit()

    universities_df = pd.DataFrame(universities)
    courses_df = pd.DataFrame(all_courses)

    courses_df.drop_duplicates(subset=["course_name", "university_id"], inplace=True)
    courses_df.fillna("Not Available", inplace=True)

    output_path = os.path.join(os.path.dirname(__file__), "University_Course_Data.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        universities_df.to_excel(writer, sheet_name="Universities", index=False)
        courses_df.to_excel(writer, sheet_name="Courses", index=False)

    print("Excel file created successfully!")


if __name__ == "__main__":
    main()
