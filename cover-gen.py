# -*- coding: utf-8 -*-

import os
import sys
import csv
import errno
import pandas as pd
import importlib.util

import random
import datetime
import argparse
from pathlib import Path

import openpyxl
from docx2pdf import convert
from docxtpl import DocxTemplate

def parse_args():
    '''
    Argument parser function from CLI to obtain:
        @arg [-name]: The name of the applicant, can be applicable to either the inout template but mostly used for file naming purposes
        @opt arg [--template]: The complete path (including file name) from the working directory (location of Python file) to the location of the template to be filled in
        @opt arg [--app_list]: A ".xlsx" or ".csv" of job applications in format "company", "role", (and optional) "event"
        
        @opt arg [--company]: The name of a company (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        @opt arg [--role]: The name of the desired role within the company (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        @opt arg [--event]: The name of any applicable events attended by the user within the target company/associated institutions (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        @opt arg [--other]: Any "other" content related to the application (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        
        @opt arg [--folder][--no_folder]: To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder] (in the case of generating a single applications's cover letter), 
            and potentially overrided by the [--app_list] argument (if provided will not be used)
        @opt arg [--pdf][--no_pdf]: Whether to save generated ".docx" files as a ".pdf" file, toggles between boolean true for [--pdf] and false for [--no_pdf]
        @return: argparse.ArgumentParser() object
    '''
    
    parser = argparse.ArgumentParser()
    
    parser.add_argument('-name', type=str, default=None, help='The name of the applicant, can be applicable to either the inout template but mostly used for file naming purposes')

    ## Arguments to be used to specify template to be generated from
    parser.add_argument('--template', type=str, default='cover-letter-template.docx', help='The complete path (including file name) from the working directory (location of Python file) to the location of the template to be filled in')

    ## Optional PATH to list of application
    parser.add_argument('--app_list', type=str, default=None, help='A ".xlsx" or ".csv" of job applications in format "company", "role", (and optional) "event"')

    ## Arguments to be used to fill in context for template and saving name
    parser.add_argument('--company', type=str, default=None, help='The name of a company (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--role', type=str, default=None, help='The name of the desired role within the company (in the case of generating a single applications''s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--event', type=str, default=None, help='The name of any applicable events attended by the user within the target company/associated institutions (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--other', type=str, default=None, help='Any "other" content related to the application (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')

    
    ## Whether to have folders generated for output (if '--multiple' this defaults to true)
    parser.add_argument('--folder', action='store_true', help='To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder] (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used) (default True)') ## Defaults folder name to company name
    parser.add_argument('--no_folder', dest='folder', action='store_false', help='To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder] (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used) (default True)')
    
    # parser.set_defaults(pdf=False, folder=None) # Set default pdf action to return true for return type as pdf
    parser.add_argument('--pdf', action='store_true', help='Whether to save generated ".docx" files as a ".pdf" file, toggles between boolean True for [--pdf] and False for [--no_pdf] (default True)')
    parser.add_argument('--no_pdf', dest='pdf', action='store_false', help='Whether to save generated ".docx" files as a ".pdf" file, toggles between boolean True for [--pdf] and False for [--no_pdf] (default True)')

    parser.set_defaults(folder=True, pdf=True)

    parser.print_usage()

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
    out_path = Path(f'./{args.company}/') if args.app_list is None else Path(f'./{app[0][0]}/')
    out_path.mkdir(parents=True, exist_ok=True) ## Returns None due to command query separation, needs separate line
    
    out_dir = out_path.as_posix() if args.folder or args.app_list is not None else ''
    return str(out_dir)

def get_file_name(*app):
    '''
    Obtains the raw file name of the saved cover letter in the format of "First Last-Company-Role"-Cover-Letter
        @param *app: *args row of applications in order of "company", "role", "event"
        @return: Name of the file without suffix for file type (e.g. ".pdf" or ".docx")
    '''
    file_name = f'{args.name}-{app[0][0]}-{app[0][1]}-Cover-Letter' if \
        args.app_list is not None else f'{args.name}-{args.company}-{args.role}-Cover-Letter' 
    return file_name

def get_complete_path(out_dir, file_name, file_type='docx'):
    '''
    Appends the file_name to out_dir, as well as the suffix file type depending on entered string
        @param out_dir: The directory of the FOLDER (company name) to output the generated cover letter
        @param file_name: The name ONLY (no file type suffix) for the cover letter to be generated
        @param file_type: The type of file to output, accepting only "docx" and "pdf" types
        @return: the fully appended path of "out_dir/file_name.[file_type]"
    '''
    file_name_type = f'{file_name}.{file_type}'

    if out_dir is None:
        return file_name_type
    
    out_path = os.path.join(out_dir, file_name_type)
        
    return out_path
    
def save_cl(template, *app):
    '''
    Programatically saves the generated cover letter based on whether to use a subfolder and whether to generate a pdf file based on generation from one company's information or a list of companies
        @param template: The completed template to be saved
        @param *app: *args row of applications in order of "company", "role", "event"
        @output: The saved files (docxs and potentially pdfs) deposited in respective folders (or current workign directory) as need be
    '''
    
    out_dir = get_out_dir(app[0]) if args.app_list is not None else get_out_dir() if args.folder else None ## Gets the FOLDER of the output as needed    
    file_name = get_file_name() if args.app_list is None else get_file_name(app[0]) ## Gets the NAME of the file
    
    out_docx = get_complete_path(out_dir, file_name, file_type='docx')
    out_pdf = get_complete_path(out_dir, file_name, file_type='pdf') if args.pdf else None
    
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
    try:
        template = DocxTemplate(args.template)
    except:
        raise FileNotFoundError(
            errno.ENOENT, os.strerror(errno.ENOENT, args.template)
        )
    
    context = str()
    if args.app_list is None:
        context = {
            'COMPANY': args.company,
            'ROLE': args.role,
            'EVENT': args.event,
            'OTHER': args.other,
            'DATE': date
        }
        
        template.render(context)
        save_cl(template)
        
    else:
        context = {
            'COMPANY': app[0][0],
            'ROLE': app[0][1],
            'EVENT': app[0][2],
            'OTHER': app[0][3],
            'DATE': date
        }
        
        template.render(context)
        save_cl(template, app[0])

def print_logo():
    print('='*74)
    print(r'''
 ██████  ██████  ██    ██ ███████ ██████         ██████  ███████ ███    ██ 
██      ██    ██ ██    ██ ██      ██   ██       ██       ██      ████   ██ 
██      ██    ██ ██    ██ █████   ██████  █████ ██   ███ █████   ██ ██  ██ 
██      ██    ██  ██  ██  ██      ██   ██       ██    ██ ██      ██  ██ ██ 
 ██████  ██████    ████   ███████ ██   ██        ██████  ███████ ██   ████ 
                                                                           
                                                                           
██    ██ ██████      ██████                                                
██    ██      ██    ██  ████                                               
██    ██  █████     ██ ██ ██                                               
 ██  ██  ██         ████  ██                                               
  ████   ███████ ██  ██████                                                
                                                                             ''')
    print('='*74)

if __name__ == '__main__':

    args = parse_args()    
    print_logo()
 
    if (args.role == None or args.company == None) and args.app_list == None:
        raise argparse.ArgumentTypeError('Must enter either both company and role or a ".csv"/".xlsx" file containing a list of companies and roles (row indexed)')
    
    date = get_today()

    count_gen = 0
    if args.app_list is not None:
        global app_list
        try:
            app_df = pd.read_csv(args.app_list)
        except Exception:
            app_df = pd.read_excel(args.app_list)
        except:
            print('No ".xlsx" or ".csv" file found at entered file location')
        
        for app in app_df[['company', 'role', 'event', 'other']].to_numpy():
            render_cl(app)
            count_gen += 1
    
    else:   
        render_cl()
        count_gen += 1
        
    PDF_num = count_gen if args.pdf else 0
    print(f'Generated {count_gen} cover letters and {PDF_num} PDFs')
    print('='*74)
