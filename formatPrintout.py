from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
import subprocess
import json
import datetime
import random


def testPrint():
    print("Connected to formatter")

correct_answers = {
    "a1" : "2",
    "a2" : "37",
    "a3" : "2"
}


count =0


answer_lookup = {
    "en": {
        "a1": {
            "1": "do",
            "2": "do not",
        },
        "a3": {
            "1": "how you see and identify yourself",
            "2": "how others see and identify you",
            "3": "how you see and identify yourself and how others see and identify you", 
            "4": "neither how you see and identify yourself nor how others see and identify you"
        },
        "a2": {
            "0": "are",
            "1": "are",
            "2": "are",
            "3": "are",
            "4": "are",
            "5": "are",
            "6": "are",
            "7": "are",
            "8": "are",
            "9": "are", 
            "10": "are",
            "11": "are",
            "12": "are",
            "13": "are",
            "14": "are",
            "15": "are", 
            "16": "are",
            "17": "are",
            "18": "are",
            "19": "are",
            "20": "are", 
            "21": "are", 
            "22": "are",
            "23": "are", 
            "24": "are", 
            "25": "are", 
            "26": "are", 
            "27": "are", 
            "28": "are", 
            "29": "are",
            "30": "are", 
            "31": "are", 
            "32": "are", 
            "33": "are",
            "34": "are", 
            "35": "are", 
            "36": "are", 
            "37": "are not" 
        },
    },
    "es": {
        "a1": {
            "1": "Si",
            "2": "No"
        },
        "a3": {
            "1": "cómo te ves e identificas",
            "2": "como otres te ven e identifican",
            "3": "cómo te ves e identificas y como otres te ven e identifican",
            "4": "nada"
        },
               
        "a2": {
            "0": "puede",
            "1": "puede",
            "2": "puede",
            "3": "puede",
            "4": "puede",
            "5": "puede",
            "6": "puede",
            "7": "puede",
            "8": "puede", 
            "9": "puede", 
            "10": "puede",
            "11": "puede",
            "12": "puede",
            "13": "puede",
            "14": "puede",
            "15": "puede", 
            "16": "puede", 
            "17": "puede",
            "18": "puede",
            "19": "puede",
            "20": "puede", 
            "21": "puede", 
            "22": "puede", 
            "23": "puede", 
            "24": "puede", 
            "25": "puede", 
            "26": "puede", 
            "27": "puede", 
            "28": "puede", 
            "29": "puede", 
            "30": "puede", 
            "31": "puede", 
            "32": "puede", 
            "33": "puede", 
            "34": "puede", 
            "35": "puede", 
            "36": "puede", 
            "37": "no puede" 
        },
    }
}


from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def checkedElement():
    elm = OxmlElement('w:checked')
    elm.set(qn('w:val'),"true")
    return elm

zipcode_data_lookup = {}
country_index_score = {}

# with open('/home/pi/zipcode_data.txt') as json_file:
#     zipcode_data_lookup = json.load(json_file)

with open('country_json_data.json', 'r') as myfile:
    data=myfile.read()

# parse file
country_index_score = json.loads(data)

def score_answers(userInfo, country_score):
    number_correct = 0
    for answer in correct_answers:
        print(answer, userInfo[answer])
        if (userInfo[answer] == correct_answers[answer]):
            print("Correct")
            number_correct +=1

    if country_score <= 23.6:
        number_correct +=1 
    return number_correct


def formatDocument(userInfo):
    print("Starting Custom Print Job")
    print(userInfo)
    doc = Document("/Users/amilcook/Documents/printouts/CivilResponsesEN.docx")
    lang = userInfo["lang"]
    if userInfo["lang"] == "es":
        doc = Document("/Users/amilcook/Documents/printouts/CivilResponsesES.docx")


    country_name = country_index_score[userInfo["countryName"]]["country_name"]
    country_score = float(country_index_score[userInfo["countryName"]]["score"])

    print(country_name, country_score)

    styles = doc.styles
    style = styles.add_style('Insertion', WD_STYLE_TYPE.PARAGRAPH)
    style = doc.styles['Insertion']
    font = style.font
    font.name = 'Helvetica'
    font.size = Pt(18)


    num_correct = score_answers(userInfo, country_score)
    qualify_status = "QUALIFY"  if num_correct == 4  else "DISQUALIFY"
    print("QUALIFY : " , qualify_status)
    print("NUM CORRECT ", num_correct)


    print("Creating Document using this info")
    print(userInfo)
    for paragraph in doc.paragraphs:
        # print(paragraph.text)

        if '[DATE]' in paragraph.text:
            paragraph.style = 'Insertion'
            paragraph.text =f"{datetime.datetime.now():%m-%d-%Y}" 

        if '[NAME]' in paragraph.text:
            paragraph.style = 'Insertion'
            paragraph.text = 'Applicant: {}'.format(userInfo["userName"])
            if lang == "es":
                paragraph.text = 'Solicitante: {}'.format(userInfo["userName"])

        if '[QUALIFYSTATUS]' in paragraph.text:
            paragraph.style = 'Insertion'
            paragraph.text = 'Result : {}'.format(qualify_status)
            if lang == "es":
                paragraph.text = "Resultado : No Califica"

        if '[Q1]' in paragraph.text:   
            q1Answer = answer_lookup[lang]["a1"][userInfo["a1"]]

            if lang == "en":
                paragraph.text = "You [{}] know or have heard of the suspects and witnesses.".format(q1Answer)
            else:
                q1Answer = answer_lookup[lang]["a1"][userInfo["a1"]]
                paragraph.text = "Usted [{}] los testigos o los sospechosos.".format(q1Answer)

        if '[Qc2]' in paragraph.text:
            questionTwoAnswer = answer_lookup[lang]["a2"][userInfo["a2"]]
            paragraph.text = "You [{}] related or know a descendent of an immigrant from the 1824 migration from the United States to Haiti/Dominican Republic.".format(questionTwoAnswer)
  
            if lang == "es":
                paragraph.text = "Usted [{}] ser un familiar o conoce a un descendiente de inmigrante de la migración de 1824 de los Estados Unidos a Haití/República Dominicana.".format(questionTwoAnswer)

        if '[Q3]' in paragraph.text:
            questionThreeAnswer = answer_lookup[lang]["a3"][userInfo["a3"]]
            print("paragraph text", paragraph.text)
            paragraph.text = "For you, [{}] affects your material conditions.".format(questionThreeAnswer)
            if lang == "es":
                paragraph.text = "Para usted, [{}] afecta sus condiciones materiales.".format(questionThreeAnswer)

#        if '[Sugar]' in paragraph.text:
#            questionFourAnswer = answer_lookup[lang]["sugarIntake"][userInfo["sugarIntake"]]
#            paragraph.text = "For you, [{}] affects your material conditions.".format(answer_lookup[lang]["sugarIntake"][userInfo["sugarIntake"]])
#            if lang == "es":
#                paragraph.text = "Su consumo semanal de azúcar es [Q3 answer].".format(questionFourAnswer)

        if '[ANSWER]' in paragraph.text:
            print("Setting Random Value 1")
            paragraph.style = 'Insertion'

            paragraph.text = "Your [{}] nationality ranks {}% in the Quality of Nationality index".format(country_name,country_score)
            if lang == "es":
                paragraph.text = "Su nacionalidad [{}] tiene un índice de {}% en el índice de calidad de la nacionalidad".format(country_name,country_score)


        if '[X out of 4]' in paragraph.text or '[X de 4]' in paragraph.text:
            print("Setting Random Value 1")
            paragraph.style = 'Insertion'

            paragraph.text = "You answered [{} out of 4] questions correctly. To be an impartial reviewer you would have to answer all the questions correctly.".format(num_correct)
            if lang == "es":
                paragraph.text = "Usted obtuvo [{} de 4] correctas. Para ser un evaluador imparcial debe responder correctamente todas las preguntas.".format(num_correct)





    doc_name = '{}.docx'.format(userInfo["userName"])
    doc.save(doc_name)
    print("Formatted And Saved Document with name {}".format(doc_name))
#    subprocess.run(["libreoffice", "--headless", "--convert-to",
#                   "pdf", "{}.docx".format(userInfo["userName"])])
#    subprocess.run(
#        ["lp", "-d", "myprinter", "{}.pdf".format(userInfo["userName"])])
#    print("Completed Format Of Document")


if __name__ == "__main__":
    try:
        formatDocument({"userName":"mattest1","userId":"d30ecf6db2dc97f363a57af1bf4f4658","a1":"1","a2":"0","a3":"3","countryName":"7","sugarIntake":"1","archivePermission":"1","lang":"en"})
    except Exception as e:
        print(e)

