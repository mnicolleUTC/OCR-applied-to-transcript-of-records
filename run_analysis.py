#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Sep 12 19:19:23 2022

@author: Nicolle Mathieu
"""
import os
import pandas as pd
from OCR import extract_data


def concatenate_transcript_of_records(folder_pdf_file):
    """
    Concatenate all transcript of records in pdf format into a dataframe
    containing information for every student.

    Parameters
    ----------
    folder_pdf_file : str
        Absolute path to the folder containing all transcript of record in
        pdf-format

    Returns
    -------
    df_all : pandas.DataFrame
        Dataframe containing all subject codes and associated grades for
        every student.
    """
    # Initialize dataframe
    df_columns = ['Subject_code', 'Grade' 'Student']
    df_all = pd.DataFrame(columns=df_columns)
    # Loop over pdf file contained in folder_pdf_file
    for file in os.listdir(folder_pdf_file):
        if file[-4:] == '.pdf':
            path_file = os.path.join(folder_pdf_file, file)
            df_student = extract_data(path_file)
            # Adding column with student name for each intermediate dataframe
            # student_name. Student's name is contained in pdf file name
            df_student["Student"] = file[:-4]
            # Concatenate df_student to a df_all which contains all the data
            df_all = pd.concat([df_all, df_student])
    # Saving df_all to csv file. This step is not mandatory but for this project
    # it allows us to anonymize the data for further analysis
    df_all.to_csv("database_result.csv", index=False)
    return df_all

def basics_analysis(path_csv_file):
    """
    Basic analysis realized on the dataframe in order to answer simple question.
    Each code bloc answer a question which is marked as comment
    In order to not disclose all transcript of records in GitHub repository,
    this function directly load the csv file with anonymized data on 100
    students. This csv file has been obtained with the function
    concatenate_transcript_of_records with a folder containing 100 transcript
    of records. An anonymization has been carried out on the file given in
    GitHub repository.
    At the end of function write a csv file whose line correspond to a question
    and the top 5 student corresponding to the answer.

    Parameters
    ----------
    path_csv_file : str
        Absolute path to the csv file concatenating all student's data.
    """
    # Loading dataframe
    df = pd.read_csv(path_csv_file, sep=",")
    # Nettoyage espace + ; mais chelou à refaire ou réecrire si possible
    # Qualité des datas à chier
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    df["WHO;"] = df["WHO;"].str.replace(';', '')
    df.columns = ['UV', 'NOTE', 'WHO']

    # Filtrage des 3 dernieres lignes pour chaque étudiant pour éviter les confusions
    # Creation df_resume pour récupérer ces datas
    df_resume = df.query('UV == ["TOTA","SPEC","ETRA","ETRB"]')
    df = df.query('UV != ["TOTA","SPEC","ETRA","ETRB"]')
    # Comptage du nombre de note par personne
    for lettre in ["A", "B", "C", "D", "E", "F", "FX"]:
        table = df[df["NOTE"] == lettre].groupby(["WHO"])["UV"].count()
        table = table.sort_values(ascending=False)
        print(table)

    # Top5 des uv les plus ratées
    df_f = df.query('NOTE == ["F","FX"]')
    top_f = df_f.groupby(["UV"])["WHO"].count()
    top_f = top_f.sort_values(ascending=False)
    # Calcul du GPA pour tous
    df["POINT_GPA"] = 0
    for pt, lettre in enumerate(["F", "E", "D", "C", "B", "A"]):
        df.loc[df["NOTE"] == lettre, "POINT_GPA"] = pt

    df_gpa = pd.DataFrame(df.groupby("WHO")["POINT_GPA"].sum())
    df_gpa["NB_UV"] = df.groupby("WHO")["UV"].count()
    df_gpa["GPA"] = df_gpa["POINT_GPA"] / df_gpa["NB_UV"]
    df_gpa["GPA"] = df_gpa["GPA"].round(2)
    df_gpa = df_gpa.sort_values(by="GPA", ascending=False)

    # Filiere la plus représenté
    df_spec = df_resume[df_resume["UV"] == "SPEC"]
    df_spec_count = df_spec.groupby(["NOTE"])["UV"].count()
    df_spec_count = df_spec_count.sort_values(ascending=False)

    # Nombre total de crédit
    df_cred = df_resume[df_resume["UV"] == "TOTA"]
    df_cred['periode'] = df_cred['NOTE'].apply(lambda row: row.split(';')[0])
    df_cred['tot_cred'] = df_cred['NOTE'].apply(lambda row: row.split(';')[-1])
    df_cred['tot_cred'] = df_cred['tot_cred'].astype(int)
    # Le plus rousto et le moin
    # df_cred = df_cred.sort_values(by = "tot_cred",ascending = False)
    df_cred["nb_sem"] = df_cred.apply(lambda row: analyse_sem(row["periode"]),
                                      axis=1)
    df_cred["renta"] = df_cred["tot_cred"] / df_cred["nb_sem"]
    df_cred["renta"] = df_cred["renta"].round(2)
    df_cred = df_cred.sort_values(by="renta", ascending=False)

    # Extract ETRA et ETRB only
    df_out = df_resume.query('UV == ["ETRA","ETRB"]')
    df_out.to_csv("etranger.csv", sep=',')

def analyse_semester(semester_acronym):
    """
    Analyse semester acronym and deduce the number of semester followed by the
    student.

    Parameters
    ----------
    semester_acronym : str
        Acronym of the beginning and end semester of schooling

    Returns
    -------
    number_of_semester : int
        Number of semester followed by the student at university
    """
    begin_semester, end_semester = semester_acronym.split('-')
    type_begin_sem, year_begin_sem = begin_semester[0], begin_semester[1:]
    type_end_sem, year_end_sem = end_semester[0], end_semester[1:]
    number_of_years = int(year_end_sem) - int(year_begin_sem)
    if type_end_sem == "P":
        number_of_semester = number_of_years * 2
    else:
        number_of_semester = number_of_years * 2 + 1
    return number_of_semester

if __name__ == '__main__':
    # Defining folder containing transcript of records
    folder_pdf = os.path.join(os.getcwd(), "Transcript_of_records")
    # Concatenate in one dataframe all data extracted from transcript of records
    df_all_students = concatenate_transcript_of_records(folder_pdf)
    # Run basics analysis on the dataframe
    basics_analysis(os.path.join(os.getcwd(), "Student_data_anonymized"))
