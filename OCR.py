#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Aug 15 11:39:57 2022
Description of the program available at :
https://github.com/mnicolleUTC/OCR-applied-to-transcript-of-records.git
@author: nicollemathieu
"""

import os
import re
import sys
import pandas as pd
import pytesseract
import unidecode
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from pdf2image import convert_from_path


def convert_pdf(pdf_path,dpi_val,save_image = False,save_raw_text = False):
    """
    Extract text from pdf file into a list of image (jpeg format)

    Parameters
    ----------
    pdf_path : str
        Absolute or relative path to the PDF file
    dpi_val : int
        Dots per inch, can be seen as the relative resolution of the output 
        PDF, higher is better but anything above 300 is usually not discernable
        to the naked eye. Recommanded value ensuring OCR performance = 600
    save_image : bool, optional
            Enable saving of the output images produced by pdf2image.
            The default is False.
    save_raw_text : bool, optional
            Enable saving of the raw text data extracted with tesseract OCR.
            The default is False.           

    Returns
    -------
    list_pages : list
        List containing raw text data extracted with tesseract OCR from each
        pdf page
    """
    main_dir = os.getcwd()
    custom_config = r'--oem 1 --psm 4'
    #Source for custom_config = https://muthu.co/all-tesseract-ocr-options/
    #oem 1 = Neural nets LSTM engine only.
    #psm 4 = reading document as multiple columns
    pages = convert_from_path(file, dpi_val)
    list_pages = list()
    for i,page in enumerate(pages):
        #Extract data of page i
        txt = pytesseract.image_to_string(page, config=custom_config,\
                                          lang = 'fra')
        #Extract student name in the first page for saving purpose
        if i == 1:
            student_name = identify_name(txt)
        #Saving data if option(s) activated
        if save_image or save_raw_text:
            save_folder = os.path.join(main_dir,"log_folder")
            if not os.path.isdir(save_folder):
                os.mkdir(save_folder)
            if save_image :
                page.save(os.path.join(main_dir,f'{student_name}_page{i}.jpg'))
            if save_raw_text:
                with open(f'{student_name}.txt','a') as file_raw_text:
                    file_raw_text.write(txt)
        #Append page by page list_pages variable
        list_pages.append(txt)
    return list_pages

def identify_name(raw_text_data):
    """
    Identify student's name from raw text data extrated with tesseract OCR

    Parameters
    ----------
    raw_text_data : str
        Raw text data extrated with tesseract OCR

    Returns
    -------
    student_name : str
        Student's name extracted from the first page of pdf file
    """
    #Split raw text based on line break
    split_text = raw_text_data.split('\n')
    #Filter empty string from list
    split_text = list(filter(None,split_text))
    #Identify sentence before name 
    pattern_intro = 'Le directeur de l\'université de technologie de '\
                    'Compiègne (UTC), soussigné, certifie que'
    index_line_pattern = split_text.index(process.extractOne(pattern_intro,\
                         split_text, scorer=fuzz.token_sort_ratio)[0])
    
    #Extract name and return it with format name_surname
    line_name = split_text[index_line_pattern + 1]
    split_line_name = re.split(',|\s', line_name)
    split_line_name = list(filter(None,split_line_name))
    student_name = split_line_name[0].lower() + '_' + split_line_name[1].lower()
    return student_name
    
def identification_semestre_etranger(page):
    """
    Identify if the student has study a semester abroad.
    In this case identify the country and number of credits earned during this
    period

    Parameters
    ----------
    page : str
        Raw text data extracted with tesseract OCR from first page

    Returns
    -------
    erasmus_credits : int
        Number of credits earned by the student during his erasmus
    erasmus_destination : str
        Country of erasmus semester
    """
    #Identify if the student has study a semester abroad.
    pattern_erasmus = "Enseignements suivis dans le cadre de semestres "\
                      "d'études a l'étranger"
    pattern_date = "Fait a Compiegne, le"
    ratio_id = 80
    erasmus_credits = 0
    split_text = page.split('\n')
    #Identify sentence prior erasmus information
    line_pattern = process.extractOne(pattern_erasmus, split_text,\
                                    scorer = fuzz.token_sort_ratio)
    #Erasmus semester identified if fuzzywuzzy score superior to ratio_id else
    #exit of the function
    if line_pattern[1] < ratio_id:
        erasmus_credits = 0
        erasmus_destination = "None"
        return erasmus_credits,erasmus_destination
    #Identitify erasmus destination and credits
    delimiters = ["Pays Université Crédits","Fait a Compiegne, le"]
    index_delimiters = []
    for d in delimiters:
        line_pattern = process.extractOne(d, split_text,\
                                        scorer=fuzz.token_sort_ratio)
        index_delimiters.append(split_text.index(line_pattern[0]))
    #Filtering split_text to retain only erasmus information
    split_text = split_text[index_delimiters[0]+1:index_delimiters[1]]
    erasmus_information = [elt for elt in split_text if elt][0]
    split_erasmus = re.split(',|\s', erasmus_information)
    #Fetch country by taking first element of the list
    erasmus_destination = split_erasmus[0]
    #Fetch erasmus credits which is the last number of the list
    erasmus_credits = int([num for num in split_erasmus if num.isdigit()][-1])
    return erasmus_credits,erasmus_destination

def acronyme_semestre(word):
    """
    Return letter 'P' for 'printemps' or a similar word and 'A' for 'automne'
    or a similar word. Word are considered similar if fuzz.ratio > 80.
    If no fuzz.ratio are at least superior to 80 return error because the input
    word is not considered enough close to 'printemps' or 'automne' which 
    should always be the case

    Parameters
    ----------
    word : str
        Word corresponding to semester period type extracted from 
        raw text data

    Returns
    -------
    str
        Letter 'P' corresponding to 'printemps' or 
        Letter 'A' corresponding to 'automne'.

    """
    if fuzz.ratio(word,'printemps') > 80:
        return 'P'
    elif fuzz.ratio(word,'automne') > 80:
        return 'A'
    else:
        return 'ERROR'
    
def determination_periode_with_indice(liste_indice,liste_scanned):
    #Sorting of the index by ascending order
    liste_indice.sort()
    #Fetching 4 word corresponding to semester and year of beginning and end of study
    periode = [
        liste_scanned[liste_indice[0]],
        liste_scanned[liste_indice[0]+1],
        liste_scanned[liste_indice[1]],
        liste_scanned[liste_indice[1]+1],        
        ]
    return periode
    
def convert_period_to_acronym(periode):
    """
    Convert a list of four elements corresponding to semester's types and
    associated years in four digit format into an acronym with format XYY-XYY
    (X the letter corresponding to semester type and Y the year expressed in
    two digits format). Example convert the list :
    ["le printemps", '2011', "l'automne", '2016,'] into "P11-A16"

    Parameters
    ----------
    periode : list
        List containing the type and the year for beginning and end of
        schooling

    Returns
    -------
    result = str
        Acronym which indicates the beginning and end of schooling

    """
    result = '{0}{1}-{2}{3}'.format(acronyme_semestre(periode[0]),\
                                    periode[1][-2:],\
                                    acronyme_semestre(periode[2]),\
                                    periode[3][-2:],)
    return result

def identification_periode_double(test_periode,liste_scanned):
    #Fetching the 2 word from test_periode
    candidat = [test_periode[0][0],test_periode[1][0]]
    #2 cases : both word in liste_candidat are the same or are slightly different
    if candidat[0] == candidat[1]:
        #2 index to be found with only one word
        indice = [i for i, x in enumerate(liste_scanned) if x == candidat[0]]
    else:
        indice = [liste_scanned.index(candidat[0]),liste_scanned.index(candidat[1])]
    periode = determination_periode_with_indice(indice,liste_scanned)
    return periode
    
def identification_data_resume(page):
    #Permet d'identifier la période des études, le nombre de crédits total 
    #validés et la spécialité de l'étudiant
    pattern_periode = "a obtenu, dans le cadre de son inscription à l'UTC"
    ratio_id = 80
    split_text = page.split('\n')
    line_intro = process.extractOne(pattern_periode, split_text,\
                                    scorer = fuzz.token_sort_ratio)
    index_line_intro = split_text.index(line_intro[0])
    full_line = split_text[index_line_intro]+ ' ' + split_text[index_line_intro+1]
    split_line = full_line.split()
    #1er bloc: identification de la période d'étude
    #Extraction des candidats pour chaque mot clés
    semestre_prin = process.extract("printemps",split_line)
    semestre_auto = process.extract("l'automne",split_line)
    #Conservation des 2 meilleurs candidats pour début et fin 
    liste_candidat = semestre_auto + semestre_prin
    liste_candidat.sort(key=lambda x: x[1], reverse=True)
    match = [liste_candidat[0][0],liste_candidat[1][0]]
    #Determination si match contient 2 printemps, 2 automne ou les deux
    test_prin = process.extract("printemps",match)
    test_auto = process.extract("l'automne",match)
    if test_auto [0][1] and test_auto[1][1] > ratio_id:
        #Cas de deux automnes
        periode = identification_periode_double(test_auto, split_line)
    elif test_prin [0][1] and test_prin[1][1] > ratio_id:   
        #Cas de deux printemps
        periode = identification_periode_double(test_prin, split_line)
    else:
        #Cas d'un printemps et d'un automne
        indice = [
            split_line.index(test_prin[0][0]),
            split_line.index(test_auto[0][0]),
            ]
        periode = determination_periode_with_indice(indice,split_line)
    #Convert periode in liste format to str format with acronym
    periode = convert_period_to_acronym(periode)
    #2eme bloc : Identification du nombre de crédits
    #Identification de l'index du mot crédit
    word_credit = process.extractOne("crédits", split_line)[0]
    index = split_line.index(word_credit)
    try:
        credit = int(split_line[index-1])
        """
        Attention à refaire
        """
        if credit < 130:
            #Etudiant fini en tronc commun
            specialite = 'TC'
            return credit,periode,specialite
    except:
        credit = "ERR"
    
    """
    if credit < 130:
        #Etudiant fini en tronc commun
        specialite = 'TC'
        return credit,periode,specialite
    """
    #3ème bloc : Identification de la spécialité de l'étudiant
    pattern_spec = "étudiant en spécialité"
    line_spec = process.extractOne(pattern_spec, split_text,\
                                scorer = fuzz.token_sort_ratio)[0]
    #Determination de l'indice du chiffre dans la phrase
    indice_number = line_spec.find(re.findall('[0-9]',line_spec)[-1])
    #Suppresion de la phrase du chiffre et de ce qui suit
    line_spec = line_spec[0:indice_number]
    #Séparation ligne par des espaces
    split_line = line_spec.split()
    #Identification du mot spécialité dans la liste
    spe_word = process.extractOne('spécialité',split_line)[0]
    index = split_line.index(spe_word)
    liste_filtered = split_line[index+1:]
    spe = ' '.join(liste_filtered)
    specialite = convert_speciality_to_acronym(spe)
    return credit,periode,specialite

def convert_speciality_to_acronym(speciality):
    """
    Convert the name of engineer speciality identify on raw text data into 
    the acronym used at UTC and defined in dict_speciality
    
    Parameters
    ----------
    speciality : str
        Word in the raw text data identify as engineer speciality

    Returns
    -------
    acronym : str
        Acronym corresponding to speciality based on variable dict_speciality
        which link each speciality to his acronym.

    """
    dict_speciality = {
        'biologie':'GB',
        'informatique':'GI',
        'mécanique':'GM',
        'procédés':'GP',
        'urbain':'GSU',
        }
    identified_speciality = process.extractOne(speciality,\
                                               dict_speciality.keys())
    acronym = dict_speciality[identified_speciality[0]]
    return acronym
    
def clean_data(i,page):
    #Definition des bornes de séaration du texte
    if i == 0:
        bornes = ["Enseignements UTC Note ECTS Crédits","page 1 de 2"]
    elif i == 1:
        borne_fin, credits_etranger = identification_semestre_etranger(page)
        bornes = ['Enseignements UTC (suite ...) Note ECTS  Crédits'\
                  ,"page 2 sur 2"]
    else:
        print("Erreur 3 pages détectées")
        sys.exit()
    #Identification de la ligne correspndante with fuzzy wuzzy
    split_text = page.split('\n')
    """
    indice = []
    for b in bornes:
        line_borne = process.extractOne(b, split_text,\
                                        scorer=fuzz.token_sort_ratio)
        indice.append(split_text.index(line_borne[0]))
    #Filtre de la liste en fonction des indices identifiés
    split_text = split_text[indice[0]+1:indice[1]]
    """
    return split_text

def detect_first_letter(line):
    """
    Detect the index of the first letter in the line given as argument

    Parameters
    ----------
    line : str
        Line extracted from raw text data and separated by \n.

    Returns
    -------
    int
        Index of the first letter in the line.

    """
    for i,character in enumerate(line):
        if character.isalpha() or character in ['1','!'] :
            return i
    return 0

def clean_line(line):
    #Definition des cas de remplacement de lettre et chiffre séparement
    clean_letter = {
        '0':'O',
        'o':'O',
        '!':'I',
        '1':'I',
        'i':'I',
        'l':'I',
        }
    clean_number = {
        'O':'0',
        'o':'0',
        'I':'1',
        'i':'1',
        'l':'1',
        '!':'1',
        }
    #Reconnaissance du premier caractère qui est une lettre
    index_letter = detect_first_letter(line)
    base = line[index_letter:index_letter+4]
    #Boucle sur les 4 caracteres:
    for i,character in enumerate(base):
        if i in [0,1] and character in clean_letter.keys():
            base = base.replace(character,clean_letter[character])
        elif i in [2,3] and character in clean_number.keys():
            base = base.replace(character,clean_number[character])
    #Remplacement dans la ligne du sigle
    line = base + line[index_letter+4:]
    return line

def identification_uv(motif_regex,line,interrupteur = 0):
    uv = re.findall(motif_regex,line)
    if len(uv) == 1:
        #Test sur existance du pattern
        uv = uv[0]          
    elif interrupteur == 0:
        #Test of change pattern
        line = clean_line(line)
        interrupteur = 1
        uv = identification_uv(motif_regex,line,interrupteur=1)
    else:
        #Else after change pattern : xxxx
        uv= "xxxx"
    return uv
    
def extract_line_data(line):
    #Suppression des accents dans chaque string
    line = unidecode.unidecode(line)
    #Identification UV Regex : 2 majuscules suivis de 2 chiffres
    motif_regex_uv = r"[A-Z]{2}[0-9]{2}"
    uv = identification_uv(motif_regex_uv,line,interrupteur = 0)
    #Identification nombre de credit
    if line.split()[-1].isdigit(): #Test sur l'identification d'un chiffre
        credit = int(line.split()[-1])
    else: #Si pas identifié alors on mets 99
        credit = 99
    #Identification de la note
    motif_regex_note = r" [a-eA-EG] "
    #Extract only last maj letter of the line
    liste_maj = re.findall(motif_regex_note,line)
    #Test on the existence of one element in the list
    if liste_maj:
        note = liste_maj[-1][1]
        #Change for C analysed as c
    else:
        note = 'Z'
    return uv,credit,note
    
def extract_data(list_pages):
    df = pd.DataFrame(columns=('UV','NOTE'))
    for i,page in enumerate(list_pages):
        if i == 0:
            credit_total,periode,specialite = identification_data_resume(page)
        split_text = clean_data(i,page)
        for line in split_text:
            #Test qu'au moins un élément soit présent dans la liste sinon next
            if not line.split():
                continue
            #On recupere le nombre de crédit mais non valorisé car aléatoire
            uv,credit,note = extract_line_data(line)
            df.loc[df.shape[0]] = [uv,note]
    #Ajout des lignes de fin du dataframe
    df.loc[df.shape[0]] = ['TOTA','{0};{1}'.format(periode,str(credit_total))]
    df.loc[df.shape[0]] = ['SPEC',f'{specialite}']
    return df
        
if __name__ == '__main__':
    file = '/Users/nicollemathieu/Desktop/Halouf_documents/Clean_transcript/Nicolle.pdf'
    name_outfile = '/Users/nicollemathieu/Desktop/coucou.csv'
    list_pages = convert_pdf(file,600)
    df = extract_data(list_pages)
    """
    print(df)
    df.to_csv(name_outfile,index = False)
    """
