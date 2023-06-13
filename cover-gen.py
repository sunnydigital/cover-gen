# -*- coding: utf-8 -*-
'''
@version: 3.0.0
@author: Sunny Son
    @github: sunnydigital
    
@about: Python based CLI script to automate cover letters in job applications
@license: MIT License
ASCII art font: ANSI Regular from TextKool
'''

## Standard packages
import re
import os
import sys
import errno
import datetime
import argparse
from pathlib import Path
from collections import defaultdict

## Packages requiring installation
import pandas as pd
import openpyxl
from dateutil import parser
from docx2pdf import convert
from docxtpl import DocxTemplate

def parse_args():
    '''
    Argument parser function from CLI to obtain:
        @arg [-name]: The name of the applicant, can be applicable to either the inout template but mostly used for file naming purposes
        @opt arg [--template]: The complete path (including file name) from the working directory (location of Python file) to the location of the template to be filled in
        @opt arg [--app_list]: A ".xlsx" or ".csv" of job applications in format "company", "role", (and optional) "event"
        
        @opt arg [--date]: A datetime readable string, if none exist defaults to today's date and outputs error message in console
        @opt arg [--company]: The name of the company being applied to
        @opt arg [--address]: The address of the company being applied to
        @opt arg [--role]: The name of the desired role within the company being applied to
        @opt arg [--event]: The name of any applicable events attended by the user within the target company/associated institutions
        @opt arg [--contact]: The name of any applicable contacts to the company being applied to (networking, social events, etc.)
        @opt arg [--referral]: The name of the person giving the applicant (current user) a referral to the company
        @opt arg [--hmanager]: The name of the hiring manager cover letter is to be sent to
        @opt arg [--convo1]: A first blurb of meaningful conversation to be included in the cover letter
        @opt arg [--convo2]: A second blurb of meaningful conversation to be included in the cover letter
        @opt arg [--other1]: A first "other" content related to the application
        @opt arg [--other2]: A second "other" content related to the application
        
        @opt arg [--folder][--no_folder]: To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder]
            All above two blocks to be used in generating the cover letter for a single application), and potentially overrided by the [--app_list] argument (and if still provided will not be used)
        
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
    parser.add_argument('--date', type=str, default=None, help='A datetime readable string, if none exist defaults to today\'s date and outputs error message in console')
    parser.add_argument('--company', type=str, default=None, help='The name of a company (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--address', type=str, default=None, help='The address of the company being applied to')
    parser.add_argument('--role', type=str, default=None, help='The name of the desired role within the company (in the case of generating a single applications''s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--event', type=str, default=None, help='The name of any applicable events attended by the user within the target company/associated institutions (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--contact', type=str, default=None, help='The name of any applicable contacts to the company being applied to (networking, social events, etc.) (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--referral', type=str, default=None, help='The name of the person giving the applicant (current user) a referral to the company (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--hmanager', type=str, default=None, help='The name of the hiring manager cover letter is to be sent to (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--convo1', type=str, default=None, help='A first blurb of meaningful conversation to be included in the cover letter (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--convo2', type=str, default=None, help='A second blurb of meaningful conversation to be included in the cover letter (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--other1', type=str, default=None, help='A first "other" content related to the application (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')
    parser.add_argument('--other2', type=str, default=None, help='A second "other" content related to the application (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used)')

    ## Whether to have folders generated for output (if '--multiple' this defaults to true)
    parser.add_argument('--folder', action='store_true', help='To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder] (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used) (default True)') ## Defaults folder name to company name
    parser.add_argument('--no_folder', dest='folder', action='store_false', help='To determine whether or not to save generated cover letters in a subfolder saved as a boolean true in the case of [--folder] and false [--no_folder] (in the case of generating a single applications\'s cover letter), and potentially overrided by the [--app_list] argument (if provided will not be used) (default True)')
    
    # parser.set_defaults(pdf=False, folder=None) # Set default pdf action to return true for return type as pdf
    parser.add_argument('--pdf', action='store_true', help='Whether to save generated ".docx" files as a ".pdf" file, toggles between boolean True for [--pdf] and False for [--no_pdf] (default True)')
    parser.add_argument('--no_pdf', dest='pdf', action='store_false', help='Whether to save generated ".docx" files as a ".pdf" file, toggles between boolean True for [--pdf] and False for [--no_pdf] (default True)')

    parser.set_defaults(folder=True, pdf=True)

    parser.print_usage()

    return parser.parse_args()

def get_date(use_today, *app):
    '''
    Obtains today's date in the format
        Month dd, YYYY:
        e.g. May 28, 2023
    And automatically returns today's date if the date is not provided or illlegible
        @param use_today: Automatically returns today\'s date if flagged true
        @param *app: *args row of applications in order of "company", "role", "event"
        @return: datetime.date string of format Month dd, YYYY
    '''
    
    today = datetime.date.today()
  
    def parse_date(date_str):
        '''
        General function to take entered date, determine whether valid datetime format, and throws exception otherwise
            @param date_str: the date string to be evaluated
            @return: datetime.date object of the parsed input string with 
        '''
        
        try:
            date = parser.parse(date_str)
            return datetime.date(date.year, date.month, date.day)
        except Exception as e:
            errors['date'] += 1
            print(f'Could not parse date string: {date_str}. Error: {e}')
            return today
  
    if use_today == True:
        return today.strftime('%B %d, %Y')
    
    if app is None:
        if args.date is not None:
            try:
                date = parse_date(args.date)
            except:
                date = today
                errors['date'] += 1
                print(f'The date of the application for {args.company} has no (legible) date, defaulting to today\'s date')
        else:
            date = today
            print('='*74)
            print(f'The application for {args.company} has no associated date, defaulting to today\'s date')
            print('='*74)
            
    else:
        try:
            date = parse_date(app[0][rm['date']])
        except:
            rm_date = rm['date'] ## Not sure what version of Python allows f-strings to have so little utility but this ain't chief
            date = today
            errors['date'] += 1
            print('='*74)
            print(f'The application for {app[0][rm_date]} has no (legible) date, defaulting to today\'s date')
            print('='*74)
    
    return date.strftime('%B %d, %Y')

def get_address(*app):
    '''
    Obtains the correct address format and returns it
        e.g.:
            1234 Sciendenfield Lane
            Los Angeles, CA 90001
        @param *app: *args row of applications in order of "company", "role", "event", ..., "other"
    '''
    
    def parse_address(address_str):
        '''
        General function to take entered address, determine whether valid, and output address in valid format
        The function excepts an error if the address string cannot be parsed
            @param address_str: The string of an address to be evaluated
            @return: Correct string format if no error, else blank string
        '''
        
        rm_company = rm['company']

        try:
            elements = [element.strip() for element in address_str.split(',')]
            if len(elements) < 3:
                print('Address string for does not contain enough elements, defaulting to no address for application for company {app[0][rm_company]}')
                return ''

            street = elements[0]
            city = elements[1]
            state = elements[2].split()[0]
            postal_code = elements[2].split()[1]

            return f'{street}\n{city}, {state} {postal_code}'

        except Exception as e:
            errors['address'] += 1
            print('='*74)
            print(f'Could not parse address string: {address_str}. Error: {e} for comapny {app[0][rm_company]}')
            print('='*74)
            return ''

    if app is None:
        if args.address is None:
            return ''
        else:
            return parse_address(args.address)
    else:
        return parse_address(app[0][rm['address']])

def get_out_dir(*app):
    '''
    Gets (or creates if doesn't exist) the folder associated with the company the cover letter is for, programatically determining whether determing company from args.company or the specific row in *app
        @param *app: *args row of applications in order of "company", "role", "event"
        @return: String version of the relative path of the folder assocaited to the comapny, or created otherwise
    '''
    
    ## Both availabilities below default the out path to the name of the entered company as a subfolder
    rm_company = rm["company"]
    out_path = Path(f'./{args.company}/') if args.app_list is None else Path(f'./{app[0][rm_company]}/')
    out_path.mkdir(parents=True, exist_ok=True) ## Returns None due to command query separation, needs separate line
    
    out_dir = out_path.as_posix() if args.folder or args.app_list is not None else ''
    return str(out_dir)

def get_file_name(*app):
    '''
    Obtains the raw file name of the saved cover letter in the format of "First Last-Company-Role"-Cover-Letter
        @param *app: *args row of applications in order of "company", "role", "event"
        @return: Name of the file without suffix for file type (e.g. ".pdf" or ".docx")
    '''
    
    rm_company = rm['company']
    rm_role = rm['role']
    file_name = f'{args.name}-{app[0][rm_company]}-{app[0][rm_role]}-Cover-Letter' if \
        args.app_list is not None else f'{args.name}-{args.company}-{args.role}-Cover-Letter' 
    return file_name

def get_complete_path(out_dir, file_name, file_type='docx'):
    '''
    Appends the file_name to out_dir, as well as the suffix file type depending on entered string
        @param out_dir: The directory of the FOLDER (company name) to output the generated cover letter
        @param file_name: The name ONLY (no file type suffix) for the cover letter to be generated
        @param file_type: The type of file to output, accepting only "docx" and "pdf" types
        @return: The fully appended path of "out_dir/file_name.[file_type]"
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

def get_df_hash(df, ret_idx=True):
    '''
    Function to obtain a hashing between the index of the generated numpy array in our case and the columns of a given dataframe
        @param df: The dataframe in question
        @param ret_idx: A boolean value to determine whether to return a hashing with index as the key or value of column as the key
            ret_idx=True: [col value] -> [list idx]
            ret_idx=False: [list idx] -> [col value]
        @return: dictionary with return values
    '''
    df_hash = {idx: value for idx, value in enumerate(union_list)} if not ret_idx else \
        {value: idx for idx, value in enumerate(union_list)}
    return df_hash

def get_union_list(df):
    '''
    Function to return all available columns to be used in processing, and throws error if "company" or "role" doesn't exist in the columns
        Does so by first converting all the column names to lowercase and then for each potentially different spelling of:
            "Hiring Manager", 
            "Conversation 1"/"Conversation 2"
            "Other 1"/"Other 2"
        @param df: A pandas.DataFrame object representing the ".xlsx" or ".csv" for a given applicaiton tracker
        @return: a list of unions between columns acceptable by this script and columns entered by the user
    '''
    
    ## Changes all columns to lower case
    df_cols = df.columns.str.lower()
    
    ## Changes phonetically correct spelling of "Hiring Manager" to "hmanager":
    df_cols = [re.sub('hiring manager', 'hmanager', col) for col in df_cols]
    
    ## Changes implementation of "Conversation 1" to "convo1" and the like
    df_cols = [re.sub('conversation (\d+)', r'convo\1', col) for col in df_cols]
    
    ## Changes Implementation of "Other 1" to "other1" and the like
    df_cols = [re.sub('other (\d+)', r'convo\1', col) for col in df_cols]    

    df_cols = set(df_cols)
    
    union_list = list(allowed_cols.intersection(df_cols))

    if 'company' not in union_list or 'role' not in union_list:
        raise ValueError('In input ".xlsx" or ".csv" file a company and role column must exist')
    
    return union_list

def render_cl(*app):
    '''
    Main function to create DocxTemplate object, programatically determining for singular or multiple cover letters to be generated, creating a template based on contexts given, 
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
        context = { ## Defaults to None otherwise
            'DATE': get_date(use_today=False) if args.date else get_date(use_today=True),
            'COMPANY': args.company,
            'ADDRESS': get_address(args.address) if args.address else '',
            'ROLE': args.role,
            'EVENT': args.event,
            'CONTACT': args.contact,
            'REFERRAL': args.referral,
            'HMANAGER': 'Dear' + args.hmanager if args.hmanager else 'To Whom it May Concern',
            'CONVO1': args.convo1,
            'CONVO2': args.convo2,
            'OTHER1': args.other1,
            'OTHER2': args.other2,
        }
        
        template.render(context)
        save_cl(template)
        
    else:
        context = {
            'DATE': get_date(True, app[0]) if 'date' in union_list else get_date(False),
            'COMPANY': app[0][rm['company']] if 'company' in union_list else None,
            'ADDRESS': get_address(app[0]) if 'address' in union_list else None,
            'ROLE': app[0][rm['role']] if 'role' in union_list else None,
            'EVENT': app[0][rm['event']] if 'event' in union_list else None,
            'CONTACT': app[0][rm['contact']] if 'contact' in union_list else None,
            'REFERRAL': app[0][rm['referral']] if 'referral' in union_list else None,
            'HMANAGER': 'Dear' + app[0][rm['hmanager']] if 'hmanager' in union_list else 'To Whom it May Concern,',
            'CONVO1': app[0][rm['convo1']] if 'convo1' in union_list else None,
            'CONVO2': app[0][rm['convo2']] if 'convo2' in union_list else None,
            'OTHER1': app[0][rm['other1']] if 'other1' in union_list else None,
            'OTHER2': app[0][rm['other2']] if 'other2' in union_list else None,
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
                                                                           
                                                                           
██    ██ ██████     ██████      ██████                                     
██    ██      ██         ██    ██  ████                                    
██    ██  █████      █████     ██ ██ ██                                    
 ██  ██       ██         ██    ████  ██                                    
  ████   ██████  ██ ██████  ██  ██████                                     
                                                                            ''')
    print('='*74)

if __name__ == '__main__':

    allowed_cols = set(['name', 'date', 'company', 'address', 'role', 'applied', 'event', 'contact', 'referral', 'hmanager', 'convo1', 'convo2', 'other1', 'other2'])
    
    args = parse_args()
    
    print_logo()
 
    if (args.role == None or args.company == None) and args.app_list == None:
        raise argparse.ArgumentTypeError('Must enter either both "company" and "role" or a ".csv"/".xlsx" file containing a list of "companies" and "roles" (row indexed)')
    
    count_gen = 0
    global errors
    if args.app_list is not None:
        global app_list
        try:
            app_df = pd.read_csv(args.app_list)
        except Exception:
            app_df = pd.read_excel(args.app_list)
        except:
            print('No ".xlsx" or ".csv" file found at entered file location')

        global union_list ## Is this needed??
        union_list = get_union_list(app_df)
        
        global rm ## And this??
        rm = get_df_hash(app_df)

        errors = defaultdict(lambda: 'N/A')
        
        for item in union_list:
            errors[item] = 0

        ## Replaces all instances of potential words indicating the user to have applied with "yes" and the user 
        ## not having applied with the blank entry ""
        if 'applied' in union_list:
            app_df['applied'].replace({r'([Aa]pplied)|([Ss]ent)|(Yes)|[Xx]': 'yes',
                                          r'([Nn]ot [Aa]pplied)|([Nn]ot [Ss]ent)|([Nn]o)': ''}, 
                                          regex=True,
                                          inplace=True
            )
        ## Above code might be deprecated, but is not in conflict at the moment, so will be left in

        for app in app_df[union_list].to_numpy():
            if 'applied' in union_list and app[rm['applied']] != '': # Skips current row based on whether or not the "applied" column exists and whether or not already applied
                continue

            render_cl(app)
            count_gen += 1
    
    else:
        arg_names = {k: v for k, v in vars(args).items() if v is not None and k in allowed_cols}

        errors = defaultdict(lambda: 'N/A')
        for key, _ in arg_names.items():
            errors[key] = 0
        
        render_cl()
        count_gen += 1
        
    PDF_num = count_gen if args.pdf else 0
    print('='*74)
    print(f'Generated {count_gen} cover letters and {PDF_num} PDFs')
    print('='*74)
    