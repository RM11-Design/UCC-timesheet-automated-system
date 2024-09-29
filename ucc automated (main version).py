from docxtpl import DocxTemplate
from datetime import datetime
import win32com.client

doc = DocxTemplate("Hourly Timesheet template.docx")

two_months = input("Enter the two months that you have worked: ")

week_one = int(input("Enter number of hours worked for week 1: "))
week_two = int(input("Enter number of hours worked for week 2: "))
week_three = int(input("Enter number of hours worked for week 3: "))
week_four = int(input("Enter number of hours worked for week 4: "))
week_five = int(input("Enter number of hours worked for week 5: "))

no_of_days = int(input("Enter number of days worked: "))
# hrs = float(input("Enter number of hours worked: "))

# Preforms calculations
hrs = week_one + week_two + week_three + week_four + week_five
wage_and_hours = 13.12 * hrs
holiday_pay = wage_and_hours * 0.08
total_salary = wage_and_hours + holiday_pay

# This is harded coded.
boss = "John Smith"

# Automatically puts the date of today
today_date = datetime.today().strftime("%d %b, %y")

# All the data is passed into a dictionary (all_info) and used to populate placeholders
all_info = {"two_months": two_months,
            "no_of_days": no_of_days,
              "week_one": week_one,
              "week_two" : week_two,
              "week_three" : week_three,
              "week_four" : week_four,
              "week_five" : week_five,
              "hrs": hrs,
              "wage_and_hours": round(wage_and_hours,2),
              "holiday_pay": round(holiday_pay,2),
              "total_salary": round(total_salary,2),
              "boss": boss,
              "today_date": today_date}

doc.render(all_info)
doc.save("Hourly Timesheet (updated).docx")

# Automatically sends email with attachment using Outlook API? Probably

ol = win32com.client.Dispatch('Outlook.Application')

olmailitem = 0x0

newmail = ol.CreateItem(olmailitem)
newmail.Subject = 'Testing Mail'
newmail.To = 'randomemail@gmail.ie'
newmail.Body = "Test 1"

# Put in double slashes as the code won't take the attachment (unicode error).
attach = "C:\\Users\\user1\\OneDrive\\Desktop\\Python\\UCC timesheet automated\\Hourly Timesheet.docx"
newmail.Attachments.Add(attach)

# newmail.Send()

# References
# https://docxtpl.readthedocs.io/en/latest/
# https://www.makeuseof.com/send-outlook-emails-using-python/ for email automation
