"""
Cover Letter Generator by Shafiqur Khan
"""

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import os

template = "tShafiqurRKhan_CoverLetter.docx" #Cover Letter template with MailMerge->MergeField's
document = MailMerge(template)

#
#print(document.get_merge_fields())
#{'requirement', 'Trade', 'CompanyC', 'CompanyB', 'CompanyA', 'HiringManager', 'technology', 'CityProvincePostalCode', 'position', 'Company', 'location', 'tech_field', 'TaskA', 'TaskB', 'TaskC', 'SkillB', 'PositionB', 'team_focus', 'StreetAddress', 'SkillC', 'PositionC', 'SkillA', 'PositionA'}

os.system('cls')

print("*** CL Gen by Shafiqur Khan *** \n")

document.merge(
	HiringManager=raw_input("(1/23) What is the name of the Hiring Manager? [Type \"Hiring Manager\" if a name is unavailable] "),
	Company=raw_input("(2/23) What is the full name of the company? "),
	StreetAddress=raw_input("(3/23) What is the street address of the company? "),
	CityProvincePostalCode=raw_input("(4/23) What is the [City, Province, Postal Code] of the company? "),
	technology=raw_input("(5/23) What technology/service does the company excel in? "),
	location=raw_input("(6/23) What is the primary country/place of operation of the company? "),
	position=raw_input("(7/23) Which position are you applying for in the company? "),
	tech_field=raw_input("(8/23) What is a relevant technology/field you are passionate about? "),
	requirement=raw_input("(9/23) What kind of educational/experiential background does this role require? "),
	team_focus=raw_input("(10/23) What is the primary focus or objective of the team you will be working with? "),
	Trade=raw_input("(11/23) Fill in the blank: You require [a/an ______] (e.g., a programmer): "),
	PositionA=raw_input("(12/23) Fill in the blank: During my time as [Position A] ... "),
	CompanyA=raw_input("(13/23) Fill in the blank: ... at [Company A], ... "),
	SkillA=raw_input("(14/23) Fill in the blank: ... I showed [Skill A] ... "),
	TaskA=raw_input("(15/23) Fill in the blank: ... by [Task A]. "),
	PositionB=raw_input("(16/23) Fill in the blank: During my time as [Position B] ... "),
	CompanyB=raw_input("(17/23) Fill in the blank: ... at [Company B], ... "),
	SkillB=raw_input("(18/23) Fill in the blank: ... I showed [Skill B] ... "),
	TaskB=raw_input("(19/23) Fill in the blank: ... by [Task B]. "),
	PositionC=raw_input("(20/23) Fill in the blank: During my time as [Position C] ... "),
	CompanyC=raw_input("(21/23) Fill in the blank: ... at [Company C], ... "),
	SkillC=raw_input("(22/23) Fill in the blank: ... I showed [Skill C] ... "),
	TaskC=raw_input("(23/23) Fill in the blank: ... by [Task C]. ")
	)

OutputFileName = "ShafiqurRKhan_CoverLetter.docx" #Output file
	
document.write(OutputFileName) 

print("\n" + OutputFileName + " is generated! \nPlease review thoroughly before submission!")

os.system('pause')