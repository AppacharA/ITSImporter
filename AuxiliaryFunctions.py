#auxiliary functions for this program
from math import floor, ceil
import re
import datetime
from openpyxl import Workbook
from openpyxl.utils import coordinate_from_string, column_index_from_string


def dateParser(dateString): #input should be a string in format mm/dd/yy or mm-dd-yy
	values = re.split('\D', dateString, 2) #split string at most twice (into 3 resulting values), based on the decimals. 
	date = datetime.date(int(values[2]), int(values[0]), int(values[1]))
	return date

def getStudent(workerList, workerName): #given the student name in a cell, finds the full name of the worker for Shiftplanning.
#when called, workerList will be a dictionary of workers (See ScheduleImporter.py, under the section commented "Get List of All Workers")
	
	if workerName == "": #Skips blank cells.
		return
	else:
		lastName = workerList.get(workerName)
		firstName = workerName
		if firstName == "Sophia M." or firstName == "Sophia B.": #Built in Sophia Processing!
			firstName = "Sophia"
		student = firstName + " " + lastName
		return student

def getDay(col): #returns Day as column number. This might actually be a redundant method but hey, it works!

	Day = ceil(col/2) + 2

	return Day

def getPosition(col): #For a given day, if the carltech is in the left column they're at the Helpdesk (CarlTech) and for the right column they're the Library (Libe CarlTech). Modify this code if position names change in the future.
	if (col%2) != 0:
		position = "CarlTech"
	else:
		position = "Libe CarlTech"

	return position

def getShiftTime(col, row): gets shift time given a certain cell.
	day = ceil(col/2)
	time = ""
	if row > 30 and day < 5: #if we're at 6:00pm on Mon-Thur
		if row > 30 and row < 33:
			time = "6:00pm-7:00pm"

		elif row > 33 and row < 36:
			time = "7:00pm-8:00pm"

		elif row > 36 and row < 39:
			time = "8:00pm-9:00pm"

		elif row > 39 and row < 42:
			time = "9:00pm-10:00pm"

		elif row > 42 and row < 45:
			time = "10:00pm-11:00pm"
		elif row > 45 and row < 48:
			time = "11:00pm-12:00am"
		elif row > 48 and row < 51:
			time = "12:00am-1:00am"

	elif row > 42 and day > 4 and day < 7: #if we're past 10:00pm on a friday or saturday
		return

	elif day == 1 or day == 3: #mon/wed morning schedule
		if row > 2 and row < 6:
			time = "8:00am-9:45am"
		elif row > 6 and row < 10:
			time = "9:45am-11:05am"
		elif row >10 and row < 14:
			time = "11:05am-12:25pm"
		elif row > 14 and row < 18:
			time = "12:25pm-1:45pm"
		elif row > 18 and row < 22:
			time = "1:45pm-3:05pm"
		elif row > 22 and row < 26:
			time = "3:05pm-4:25pm"
		elif row > 26 and row < 30:
			time = "4:25pm-6:00pm"

	elif day == 2 or day == 4: #tue/thurs morning schedule
		if row > 2 and row < 6:
			time = "8:00am-10:05am"
		elif row > 6 and row < 10:
			time = "10:05am-12:00pm"
		elif row >10 and row < 14:
			time = "12:00pm-1:10pm"
		elif row > 14 and row < 18:
			time = "1:10pm-3:05pm"
		elif row > 18 and row < 22:
			time = "3:05pm-5:00pm"
		elif row > 26 and row < 30:
			time = "5:00pm-6:00pm"

	elif day == 5: #The friday schedule
		if row > 2 and row < 6:
			time = "8:00am-9:35am"
		elif row > 6 and row < 10:
			time = "9:35am-10:45am"
		elif row >10 and row < 14:
			time = "10:45am-11:55am"
		elif row > 14 and row < 18:
			time = "11:55am-1:05pm"
		elif row > 18 and row < 22:
			time = "1:05pm-2:15pm"
		elif row > 22 and row < 26:
			time = "2:15pm-3:25pm"
		elif row > 26 and row < 30:
			time = "3:25pm-4:30pm"
		elif row > 30 and row < 33:
			time = "4:30pm-6:00pm"
		elif row > 33 and row < 36:
			time = "6:00pm-7:00pm"
		elif row > 36 and row < 39:
			time = "7:00pm-8:00pm"
		elif row > 39 and row < 42:
			time = "8:00pm-9:00pm"

	elif day == 6: #the saturday schedule
		if row > 2 and row < 6:
			time = "10:00am-11:00am"
		elif row > 6 and row < 10:
			time = "11:00am-12:00pm"
		elif row >10 and row < 14:
			time = "12:00pm-1:00pm"
		elif row > 14 and row < 18:
			time = "1:00pm-2:00pm"
		elif row > 18 and row < 22:
			time = "2:00pm-3:00pm"
		elif row > 22 and row < 26:
			time = "3:00pm-4:00pm"
		elif row > 26 and row < 30:
			time = "4:00pm-5:00pm"
		elif row > 30 and row < 33:
			time = "5:00pm-6:00pm"
		elif row > 33 and row < 36:
			time = "6:00pm-7:00pm"
		elif row > 36 and row < 39:
			time = "7:00pm-8:00pm"
		elif row > 39 and row < 42:
			time = "8:00pm-9:00pm"

	elif day == 7: #the sunday schedule
		if row > 2 and row < 6:
			time = "10:00am-11:00am"
		elif row > 6 and row < 10:
			time = "11:00am-12:00pm"
		elif row >10 and row < 14:
			time = "12:00pm-1:00pm"
		elif row > 14 and row < 18:
			time = "1:00pm-2:00pm"
		elif row > 18 and row < 22:
			time = "2:00pm-3:00pm"
		elif row > 22 and row < 26:
			time = "3:00pm-4:00pm"
		elif row > 26 and row < 30:
			time = "4:00pm-5:00pm"
		elif row > 30 and row < 33:
			time = "5:00pm-6:00pm"
		elif row > 33 and row < 36:
			time = "6:00pm-7:00pm"
		elif row > 36 and row < 39:
			time = "7:00pm-8:00pm"
		elif row > 39 and row < 42:
			time = "8:00pm-9:00pm"

		elif row > 42 and row < 45:
			time = "9:00pm-10:00pm"
		elif row > 45 and row < 48:
			time = "10:00pm-11:00pm"
		elif row > 48 and row < 51:
			time = "11:00pm-12:00am"
		elif row > 51:
			time = "12:00am-1:00am"
	return time



