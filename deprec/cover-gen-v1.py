# -*- coding: utf-8 -*-
import os
import importlib.util
import sys

import random
import datetime
import matplotlib.pyplot as plt
import argparse
from pathlib import Path

from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
from docx2pdf import convert

def parse_args():
    parser = argparse.ArgumentParser()

    ## Arguments to be used to fill in context for template
    parser.add_argument('--company', type=str, default='Apple')
    parser.add_argument('--role', type=str, default='Data Scientist')

    ## Arguments to be used for generating template
    parser.add_argument('--name', type=str, default='Sunny Son')
    parser.add_argument('--template', type=str, default='Cover-Letter-Template.docx')
    parser.add_argument('--folder', type=str, default=None)
    
    parser.set_defaults(pdf=False) # Set default pdf action to return true for return type as pdf
    parser.add_argument('--pdf', action='store_true')
    parser.add_argument('--no-pdf', dest='pdf', action='store_false')

    return parser.parse_args()

def get_today():
    today = datetime.date.today()
    return today.strftime('%B %d, %Y')

if __name__ == '__main__':
    args = parse_args()
    date = get_today()

    template = DocxTemplate(args.template)

    context = {
        'COMPANY': args.company,
        'ROLE': args.role,
        'DATE': date
    }

    template.render(context)

    out_str = f'{args.name}-{args.company}-{args.role}-Cover-Letter'
    
    out_path = Path(f'./{args.folder}/')
    out_path.mkdir(parents=True, exist_ok=True) ## Returns None due to command query separation
    
    out_folder = out_path.as_posix() if args.folder else ''
    
    out_docx = str(out_folder) + '/' + out_str + '.docx'
    out_pdf = str(out_folder) + '/' + out_str + '.pdf'

    if not args.folder:
        template.save(out_str + '.docx')
    else:
        template.save(out_docx)

    if args.pdf:
        if not args.folder:
            convert(out_str + '.docx', out_str + '.pdf')
        else:
            convert(out_docx, out_pdf)
