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
    df : pandas.DataFrame
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


if __name__ == '__main__':
    folder_pdf = os.path.join(os.getcwd(), "Transcript_of_records")
    df = concatenate_transcript_of_records(folder_pdf)
