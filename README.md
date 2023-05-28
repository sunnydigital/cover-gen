# Cover Letter Generator

Repository for a simple Python-baed command line cover letter generator using a base `.docx` template

## Dependencies

This script requires three packages that are not automatically installed with `Python 3.9`, and as such please run `pip install` on:

- `docx`
- `docxtpl`
- `docx2pdf`

## Usage

 - Change current directory to the location of the template (by default named `"Cover-Letter-Template.docx"`
 - Use `python cv-convert.py [--company] [COMPANY] [--role] [ROLE] [--name] [NAME] [--template] [TEMPLATE] [--folder] [FOLDER] [--pdf] [PDF]`

With specifications of the arguments as follows:

1. `--company` the name of the company applying for
2. `--role` the role applying for
3. `--template` the name of the template to be modified
4. `--folder` the name of the subfolder for the outputted `.pdf` or `.docx` file to be placed in
5. `--pdf` whether or not to output a `.pdf` or `.docx` file

### Template

Within the template (a `.docx` document), the script effectively replaces all dates, roles, and companies with the given format:

- role: `{{ROLE}}`
- company: `{{COMAPNY}}`
- date: `{{DATE}}`

As such, in the Word `.docx` document, change each mention of a role, company, and date accordingly.
