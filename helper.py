from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Cm
import re
from docx.shared import Pt 

def format_size_and_font(format_obj):
    format_obj.font.size = Pt(12)

def format_alignment(para_obj, inches=0.5):
    para_obj.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para_obj.paragraph_format.first_line_indent = Inches(inches)
    return para_obj

def format_fill_in_info(format_obj, data):
    replaceList = re.findall("\{.*?\}", format_obj)
    for replace in replaceList:
        replaceTo = replace.strip("{").strip("}")
        format_obj = re.sub(replace, str(data[replaceTo]), format_obj)
    return format_obj


def askYesNo(promptString=None):
    print(promptString)
    answer = input()
    return False if not answer.isdigit() else int(answer)

#Ask for a string
def askInput(promptString=None):
    print(promptString)
    return input()

def askForChoice(choice_obj, promptString=None):
    print(promptString, ", one of the following:")
    choices = []
    count = 0
    for key, value in choice_obj.items():
        string = '{:>12} : {:>12}'.format(count, key)
        print(string)
        choices.append(value)
        count += 1
    choice = input()
    return choices[int(choice)]

def askForChoices(choice_obj, promptString=None):
    print(promptString,", one or more of the following. Enter -1 to skip:")
    choices = []
    
    for key, value in choice_obj.items():
        choices.append(key)
    
    returnChoices = []
    while(1):
        count = 0
        for aChoice in choices:
            string = '{:>12} : {:>12}'.format(count, aChoice)
            print(string)
            count += 1
        rawInput = input()
        if rawInput == "a": # choose all
            returnChoices = choices[:4] # hardcoded 4
            break
        elif rawInput != "":
            indexes = [int(i) for i in rawInput.split(" ")]
            for i in indexes:
                returnChoices.append(choices[i])
            break
        else:
            break

    
    returnChoices = [choice_obj[choice] for choice in returnChoices]
    return returnChoices


def sanitize_name(names):
    returnName = ""
    for name in names:
        name = name \
        .replace("(", "").replace(")", "") \
        .replace(" ", "").replace("\/", "").replace("\\", "") \
        .replace(".","").replace(",", "")
        returnName += name + "_"
    returnName += ".docx"
    return returnName