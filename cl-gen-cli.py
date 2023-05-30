# -*- coding: utf-8 -*-

import os
import importlib.util
import sys
import csv
import pandas as pd

import random
import datetime
import argparse
from pathlib import Path

import openpyxl

from docxtpl import DocxTemplate
from docx2pdf import convert

def parse_args():
    '''
    Argument parser function from CLI to obtain:
        @opt arg [--app_list]: A ".xlsx" or "csv" of job applications in format "company", "role", (and optional) "event"
        @opt arg [--company]: The name of a company (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        @opt arg [--role]: The name of the desired role within the company (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        @opt arg [--event]: The name of any applicable events attended by the user within the target company/associated institutions (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        @arg [name]: The name of the applicant, can be applicable to either the inout template but mostly used for file naming purposes
        @arg [template]: The complete path (including file name) from the working directory (location of Python file) to the location of the template to be filled in
        @opt arg [--folder][--no_folder]: To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder] (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        @opt arg [--pdf][--no_pdf]: Whether to save generated ".docx" files as a ".pdf" file, toggles between boolean true for [--pdf] and false for [--no_pdf]
        @return: argparse.ArgumentParser() object
    '''
    
    parser = argparse.ArgumentParser()

    ## Optional PATH to list of application
    parser.add_argument('--app_list', type=str, default=None, help='A ".xlsx" or "csv" of job applications in format "company", "role", (and optional) "event"')

    ## Arguments to be used to fill in context for template and saving name
    parser.add_argument('--company', type=str, default=None, help='The name of a company (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--role', type=str, default=None, help='The name of the desired role within the company (in the case of generating a single applications''s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--event', type=str, default=None, help='The name of any applicable events attended by the user within the target company/associated institutions (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    
    parser.add_argument('name', type=str, default=None, help='The name of the applicant, can be applicable to either the inout template but mostly used for file naming purposes')

    ## Arguments to be used to specify template to be generated from
    parser.add_argument('template', type=str, default='cover-letter-template.docx', help='The complete path (including file name) from the working directory (location of Python file) to the location of the template to be filled in')
    
    ## Whether to have folders generated for output (if '--multiple' this defaults to true)
    parser.add_argument('--folder', action='store_true', help='To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder] (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)') ## Defaults folder name to company name
    parser.add_argument('--no_folder', dest='folder', action='store_false', help='To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder] (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    
    # parser.set_defaults(pdf=False, folder=None) # Set default pdf action to return true for return type as pdf
    parser.add_argument('--pdf', action='store_true', help='Whether to save generated ".docx" files as a ".pdf" file, toggles between boolean true for [--pdf] and false for [--no_pdf]')
    parser.add_argument('--no_pdf', dest='pdf', action='store_false', help='Whether to save generated ".docx" files as a ".pdf" file, toggles between boolean true for [--pdf] and false for [--no_pdf]')

    parser.print_help()

    return parser.parse_args()

def get_today():
    '''
    Obtains today's daty in the format
        Month dd, YYYY:
        e.g. May 28, 2023
    '''
    today = datetime.date.today()
    return today.strftime('%B %d, %Y')

def get_out_dir(*app):
    '''
    Gets (or creates if doesn't exist) the folder associated with the company the cover letter is for, programatically determining whether determing company from args.company or the specific row in *app
        @param *app: *args row of applications in order of "company", "role", "event"
        @return: String version of the relative path of the folder assocaited to the comapny, or created otherwise
    '''
    ## Both availabilities below default the out path to the name of the entered company as a subfolder
    out_path = Path(f'./{args.company}/') if args.app_list is not None else Path(f'./{app[0]}/')
    out_path.mkdir(parents=True, exist_ok=True) ## Returns None due to command query separation, needs separate line
    
    out_dir = out_path.as_posix() if args.folder else ''
    return str(out_dir)

def get_file_name(*app):
    '''
    Obtains the raw file name of the saved cover letter in the format of "First Last-Company-Role"-Cover-Letter
        @param *app: *args row of applications in order of "company", "role", "event"
        @return: Name of the file without suffix for file type (e.g. ".pdf" or ".docx")
    '''
    file_name = f'{args.name}-{app[0]}-{app[1]}-Cover-Letter' if \
        args.app_list is not None else f'{args.name}-{args.company}-{args.role}-Cover-Letter' 
    return file_name

def get_complete_path(out_dir, file_name, type='docx'):
    '''
    Appends the file_name to out_dir, as well as the suffix file type depending on entered string
        @param out_dir: The directory of the FOLDER (company name) to output the generated cover letter
        @param file_name: The name ONLY (no file type suffix) for the cover letter to be generated
        @param type: The type of file to output, accepting only "docx" and "pdf" types
        @return: the fully appended path of "out_dir/file_name.[type]"
    '''
    out_path = out_dir + '/' + file_name + ('.docx' if type == 'docx' else '.pdf' if type == 'pdf' else None) ## Potential error
    if out_path is None:
         raise ValueError('type must be "value" or "pdf"')
    return out_path
    
def save_cl(template, *app):
    '''
    Programatically saves the generated cover letter based on whether to use a subfolder and whether to generate a pdf file based on generation from one company's information or a list of companies
        @param template: The completed template to be saved
        @param *app: *args row of applications in order of "company", "role", "event"
        @output: The saved files (docxs and potentially pdfs) deposited in respective folders (or current workign directory) as need be
    '''
    out_dir = get_out_dir(app)
    file_name = get_file_name(app)
    
    out_docx = get_complete_path(out_dir, file_name, type='docx')
    out_pdf = get_complete_path(out_dir, file_name, type='pdf') if args.pdf else None
    
    if args.app_list is not None:
        template.save(out_docx)
        if args.pdf:
            convert(out_docx, out_pdf)
    
    else:
        if not args.folder:
            template.save(file_name + '.docx')
        else:
            template.save(out_docx)

        if args.pdf:
            if not args.folder:
                convert(file_name + '.docx', file_name + '.pdf')
            else:
                convert(out_docx, out_pdf)

def render_cl(*app):
    '''
    Main function to create DoxcTemplate object, programatically determining for singular or multiple cover letters to be generated, creating a template based on contexts given, 
    and calling necessary function to save generated cover letter
        @param *app: *args row of applications in order of "company", "role", "event"
    '''
    template = DocxTemplate(args.template)
    
    context
    if args.app_list is None:
        context = {
            'COMPANY': args.company,
            'ROLE': args.role,
            'EVENT': args.event,
            'DATE': date
        }
        
    else:
        context = {
            'COMPANY': app[0],
            'ROLE': app[1],
            'EVENT': app[2],
            'DATE': date
        }
        
    template.render(context)
    save_cl(template, app)

if __name__ == '__main__':
    args = parse_args()
    date = get_today()

    if args.app_list is not None:
        global app_list
        try:
            app_df = pd.read_csv(args.app_list)
        except:
            app_df = pd.read_excel(args.app_list)
        else:
            print('No ".xlsx" or ".csv" file found at entered file location')
        
        for row in app_df[['company', 'role', 'event']].to_numpy():
            render_cl(row)
    
    else:   
        render_cl()