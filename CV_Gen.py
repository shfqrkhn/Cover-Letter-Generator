"""
CV Generator by Shafiqur Khan
"""

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import os

template = "tShafiqurRKhan_Resume.docx" #CV template with MailMerge->MergeField's
document = MailMerge(template)

#print(document.get_merge_fields())
#{'Skill1', 'Skill2', 'Skill3'}

os.system('cls')

print("*** CV Gen by Shafiqur Khan *** \n")

document.merge(
	Skill1=raw_input("(1/3) Desired skill 1: "),
	Skill2=raw_input("(2/3) Desired skill 2: "),
	Skill3=raw_input("(3/3) Desired skill 3: ")
	)

OutputFileName = "ShafiqurRKhan_Resume.docx" #Output file
	
document.write(OutputFileName) 

print("\n" + OutputFileName + " is generated! \nPlease review thoroughly before submission!")

os.system('pause')