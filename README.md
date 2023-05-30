# Command Line Interface Cover Letter Generator

Repository for a simple Python-based command line cover letter generator script using a user-provided `.docx` template running on CLI

## Dependencies

This script requires three packages that are not automatically installed with `Python 3.9`, and as such please run `pip install` on:

- `docx`
- `docxtpl`
- `docx2pdf`

## Usage

Download and place the script `cl-gen.py` file to the same location as the desired cover letter template `.docx` file

 - Change current directory to the location of the template (by default named `"Cover-Letter-Template.docx"`)
 - Use `python cv-convert.py [--company COMPANY] [--role ROLE] [--name NAME] [--template TEMPLATE] [--folder FOLDER] [--pdf PDF]`

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

## Future Features

At the moment, attempting to implement two features and one potential API integration:

- [ ] Automatically downloading dependencies - oops, no idea how to really work this HAHAHahaaha... :(
- [x] The generation of multiple cover letters at once through reading in a `.xslx` or `.csv` file containing company and roles
- [x] The generation of cover letters for an `{{EVENT}}` flag, indicating any events attended by the user
- [x] The generation of cover letters for an `{{OTHER}}` flag, indicating other, wildcard options the user would like to fill
- [ ] The integration of Open AI GPT API to customize sections of cover letters
