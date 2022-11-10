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
    Transform pdf file into a list of str in which each element
    contains the raw text data obtained with OCR tool

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
            student_name = identify_student_name(txt)
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

def identify_student_name(raw_text_data):
    """
    Identify student's name from raw text data extracted with tesseract OCR

    Parameters
    ----------
    raw_text_data : str
        Raw text data extracted with tesseract OCR

    Returns
    -------
    student_name : str
        Student's name extracted from the first page of pdf file
    """
    #Split raw text based on line break
    split_text = raw_text_data.split('\n')
    #Filter empty string from list
    split_text = list(filter(None,split_text))
    #Identify sentence before student's name
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
    
def identify_erasmus_semester(page):
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

def acronym_semester(word):
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
    
def identify_education_periode(list_index,list_text_extract):
    """
    Identify education period of the student

    Parameters
    ----------
    list_index : list
        Contains the word's index corresponding to a semester in
        list_text_extract
    list_text_extract : list
        Contains words of the introduction which describes the education period
        of the student
    Returns
    -------
    education_period : list
        List containing 4 elements described as follows :
        ['type_of_begin_semester','year_of_begin_semester',
        'type_of_end_semester','year_of_end_semester']
    """
    #Sorting of the list index by ascending order
    list_index.sort()
    #Fetching 4 word corresponding to semester and year of beginning and end of
    #study
    education_period = [
        list_text_extract[list_index[0]],
        list_text_extract[list_index[0]+1],
        list_text_extract[list_index[1]],
        list_text_extract[list_index[1]+1],
        ]
    return education_period
    
def convert_education_period_to_acronym(education_period):
    """
    Convert a list of four elements corresponding to semester's types and
    associated years in four digit format into an acronym with format XYY-XYY
    (X the letter corresponding to semester type and Y the year expressed in
    two digits format). Example convert the list :
    ["le printemps", '2011', "l'automne", '2016,'] into "P11-A16"

    Parameters
    ----------
    education_period : list
        List containing the type and the year for beginning and end of
        education period

    Returns
    -------
    result = str
        Acronym which indicates the beginning and end of education_period

    """
    result = '{0}{1}-{2}{3}'.format(acronym_semester(education_period[0]),\
                                    education_period[1][-2:],\
                                    acronym_semester(education_period[2]),\
                                    education_period[3][-2:],)
    return result

def identify_same_semester_type(semester_type_candidate,list_text_extract):
    """
    Identify education period of the student in case of first and last semester
    of the same type ("printemps" or "automne"). This function is essential
    because tesseract OCR can slightly misspell semester type hence the test of
    equality between the two candidates identified by fuzzywuzzy.

    Parameters
    ----------
    semester_type_candidate : list
        List containing the two best candidates identified by fuzzywuzzy package
        for semester type
    list_text_extract : list
        Contains words of the introduction which describes the education period
        of the student

    Returns
    -------
    education_period : list
        List containing the type and the year for beginning and end of
        schooling
    """
    #Fetching the two candidates words for semester type
    candidates = [semester_type_candidate[0][0],semester_type_candidate[1][0]]
    #2 cases: Both word in candidates are the same or are slightly different
    if candidates[0] == candidates[1]:
        #2 index to be found with only one word
        index = [i for i, x in enumerate(list_text_extract) \
                 if x == candidates[0]]
    else:
        index = [
            list_text_extract.index(candidates[0]),
            list_text_extract.index(candidates[1])
        ]
    education_period = identify_education_periode(index,list_text_extract)
    return education_period
    
def identify_student_information(first_page):
    """
    Identify the basic information of the student (credits, period of education
    and engineer speciality) which are given on first page of the transcript of
    records.

    Parameters
    ----------
    first_page : str
        Raw text data extracted with tesseract OCR from first pdf page

    Returns
    -------
    credit_number : int
        Number of credits obtained by the student during his education.
    period : str
        Acronym which indicates the beginning and end of education period
    speciality : str
        Acronym corresponding to speciality chosen by the student during his
        studies
    """
    #Split raw text data of first page based on line break
    split_text = first_page.split('\n')
    #Define the pattern on first page which is before education period
    pattern_education_period = "a obtenu, dans le cadre de son inscription à "\
                               "l'UTC"
    #Extract line which is the closest to variable pattern_education_period
    line_intro = process.extractOne(pattern_education_period, split_text,\
                                    scorer = fuzz.token_sort_ratio)
    #Fetch the index corresponding to the identified line
    index_line_intro = split_text.index(line_intro[0])
    full_line = split_text[index_line_intro]+ ' ' + \
                split_text[index_line_intro+1]
    #Split the intro line based on space
    split_line = full_line.split()
    ### First block : Identification of the education period
    #Extract candidates for each type of semester
    semester_prin = process.extract("printemps",split_line)
    semester_auto = process.extract("l'automne",split_line)
    #Merge of all the fuzzywuzzy evaluation
    liste_candidat = semester_auto + semester_prin
    #Sorting of candidate based on their fuzzywuzzy score
    liste_candidat.sort(key=lambda x: x[1], reverse=True)
    #Selection of the two best matches
    match = [liste_candidat[0][0],liste_candidat[1][0]]
    #Identify if first and last semester are the same type.
    test_prin = process.extract("printemps",match)
    test_auto = process.extract("l'automne",match)
    #Word are considered similar if fuzz.ratio > 80.
    ratio_id = 80
    if test_auto [0][1] and test_auto[1][1] > ratio_id:
        #Case of first and last semesters which are "automne"
        education_period = identify_same_semester_type(test_auto, split_line)
    elif test_prin [0][1] and test_prin[1][1] > ratio_id:   
        #Case of first and last semesters which are "printemps"
        education_period = identify_same_semester_type(test_prin, split_line)
    else:
        #Case with differents semester types
        index = [
            split_line.index(test_prin[0][0]),
            split_line.index(test_auto[0][0]),
            ]
        education_period = identify_education_periode(index,split_line)
    #Convert education period into acronym
    period = convert_education_period_to_acronym(education_period)
    ### Second block : Identification of credits number and student speciality
    #Look for credits word index
    word_credit = process.extractOne("crédits", split_line)[0]
    index_word_credit = split_line.index(word_credit)
    try:
        credit_number = int(split_line[index_word_credit-1])
    except:
        credit = "ERR"
    #Case student have been expelled
    if credit_number < 130:
        speciality = 'TC'
    #Case student have complete entirely his diploma
    else:
        speciality = identify_student_speciality(split_text)
    return credit_number, period, speciality

def identify_student_speciality(split_text):
    """
    Identify student_speciality

    Parameters
    ----------
    split_text : list
        Split raw text data of lines containing student's speciality

    Returns
    -------
    speciality : str
        Acronym corresponding to speciality
    """
    pattern_spec = "étudiant en spécialité"
    line_spec = process.extractOne(pattern_spec, split_text,\
                                scorer = fuzz.token_sort_ratio)[0]
    #Text treatements
    indice_number = line_spec.find(re.findall('[0-9]',line_spec)[-1])
    line_spec = line_spec[0:indice_number]
    split_line = line_spec.split()
    #Identification of the word "spécialité"
    spe_word = process.extractOne('spécialité',split_line)[0]
    index = split_line.index(spe_word)
    liste_filtered = split_line[index+1:]
    speciality = ' '.join(liste_filtered)
    #Converting speciality into acronym
    speciality = acronym_speciality(speciality)
    return speciality

def acronym_speciality(speciality):
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

def clean_subject_code(line):
    """
    Clean line extracted from transcript of records spreadsheet in order to
    obtain the code associated to the university subject which is always
    defined by 2 uppercase letters and 2 digits. This function is
    mandatory in order to clean slight errors from OCR recognition.

    Parameters
    ----------
    line : str
        Line extracted from transcript of records spreadsheet which contains
        two key information : the subject code and the grade the student receive

    Returns
    -------
    cleaned_line : str
        Line in which characters corresponding subject code has been cleaned
        from common OCR recognition errors
    """
    #Dictionnary giving common corrections for the 2 first letters of subject
    #code
    clean_letter = {
        '0':'O',
        'o':'O',
        '!':'I',
        '1':'I',
        'i':'I',
        'l':'I',
        }
    #Dictionnary giving common corrections for the 2 last digits of subject code
    clean_number = {
        'O':'0',
        'o':'0',
        'I':'1',
        'i':'1',
        'l':'1',
        '!':'1',
        }
    #Identify the first letter of the line in order to extract subject code
    index_letter = detect_first_letter(line)
    subject_code = line[index_letter:index_letter+4]
    #Trying to correct each character of subject_code
    for i,char in enumerate(subject_code):
        if i in [0,1] and char in clean_letter.keys():
            subject_code = subject_code.replace(char,clean_letter[char])
        elif i in [2,3] and char in clean_number.keys():
            subject_code = subject_code.replace(char,clean_number[char])
    #Replacement in the original line of corrected subject code
    cleaned_line = subject_code + line[index_letter+4:]
    return cleaned_line

def identify_subject_code(subject_pattern,line,switch = 0):
    """
    Identify subject code contains in the input line extracted from transcript
    of records spreadsheet.

    Parameters
    ----------
    subject_pattern : str
        Regular expression representing subject pattern
    line :
        Line extracted from transcript of records spreadsheet
    switch : bool, optional
        Variable indicating if function clean_line has already been used.
        The default is False.

    Returns
    -------
    subject_code : str
        Subject code extracted from input line
    """
    subject_code = re.findall(subject_pattern,line)
    #Testing if subject_pattern has been detected directly without cleaning
    if len(subject_code) == 1:
        subject_code = subject_code[0]
    #If no pattern detected, new try with a cleaned line
    elif switch == 0:
        #Cleaning line from common OCR recognition errors
        line = clean_subject_code(line)
        switch = 1
        subject_code = identify_subject_code(subject_pattern,line,switch = 1)
    #If no pattern detected after cleaning, considering subject code "xxxx"
    #corresponding to an error.
    else:
        subject_code= "xxxx"
    return subject_code
    
def extract_line_data(line):
    """
    Extract key information (subject code and associated grade) from input
    line extracted from transcript of records spreadsheet.

    Parameters
    ----------
    line : str
        Line extracted from transcript of records spreadsheet

    Returns
    -------
    subject_code : str
        Subject code extracted from input line
    grade : str
        Grade obtained by the student for this subject code
    """
    #Cleaning all accents in the line
    line = unidecode.unidecode(line)
    #Defining which pattern correspond to subject code : two uppercase letters
    #followed by two digits
    subject_pattern = r"[A-Z]{2}[0-9]{2}"
    subject_code = identify_subject_code(subject_pattern,line)
    #Defining which pattern correspond to a grade (letter in lowercase or
    #uppercase between A and E. G in uppercase consider as "C"
    grade_pattern = r" [a-eA-EG] "
    #Extract all letters corresponding to grade_pattern
    grade_candidates = re.findall(grade_pattern,line)
    #Testing if at least one element is in grade_candidates
    if grade_candidates:
        #Fetching last element of grade_candidates
        grade = grade_candidates[-1]
        #Cleaning unwanted spaces and capitalize if possible
        grade = grade.replace(" ","").upper()
        #Change "G" for "C" if possible because classic OCR recognition errors
        if grade == "G" : grade = "C"
    else:
        #'Z' corresponds to an error in grade recognition
        grade = 'Z'
    return subject_code,grade
    
def extract_data(file_path):
    """
    Convert the data contained in a transcript of records in pdf-format into
    a pandas dataframe

    Parameters
    ----------
    file_path : str
        Absolute path of transcript of records in pdf-format

    Returns
    -------
    df : pandas.DataFrame
        Dataframe containing key information in the transcript of records
        (subject code and associated grade). It also contains information on
        erasmus semester, total student's credits, period of education and
        engineer speciality
    """
    #Reading transcript of record with OCR tool and convert it into
    #raw text data
    list_pages = convert_pdf(file_path,600)
    # Initialize dataframe
    df = pd.DataFrame(columns=('Subject_code', 'Grade'))
    # Iterating through each page contained in pdf file
    for i,page in enumerate(list_pages):
        # For first page identify key information about the student
        if i == 0:
            credit_number,period,speciality = identify_student_information(page)
        # For second page identify whether a semester was spent abroad
        elif i == 1:
            erasmus_credits,erasmus_country = identify_erasmus_semester(page)
        # No transcript of records can have 3 pages
        else:
            print("42 You're a God")
            sys.exit()
        split_text = page.split('\n')
        # Iterating through spreadsheets lines to extract subject code and
        # associated grade
        for line in split_text:
            subject_code,grade = extract_line_data(line)
            df.loc[df.shape[0]] = [subject_code,grade]
    # Adding resume information on the student at the end of dataframe
    df.loc[df.shape[0]] = ['PERIOD',f'{period}']
    df.loc[df.shape[0]] = ['TOTAL_CREDITS', f'{credit_number}']
    df.loc[df.shape[0]] = ['SPECIALITY',f'{speciality}']
    df.loc[df.shape[0]] = ['ERASMUS_CREDITS', f'{erasmus_credits}']
    df.loc[df.shape[0]] = ['ERASMUS_COUNTRY', f'{erasmus_country}']
    return df
        
if __name__ == '__main__':
    file = '/Users/nicollemathieu/Desktop/Halouf_documents/Clean_transcript/'\
           'Nicolle.pdf'
    df = extract_data(file)
    print(df)

