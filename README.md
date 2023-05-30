# Command Line Interface Cover Letter Generator

Repository for a simple Python-based command line cover letter generator using a user-provided `.docx` template

<img src="https://github.com/sunnydigital/cover-gen/blob/main/assets/demo_console_img.png" alt="Command Line Example" width="100%">

## Dependencies

Install all dependencies (all four) through the provided `requirements.txt` using

```Python
!pip install -r requirements.txt
```

in any `Python` or `Jupyter` file, or

```unix
python pip install -r requirements.txt
```

## Usage

Download and place the script `cl-gen.py` file to the same location as the desired cover letter template `.docx` file

 - Change current directory to the location of the template (by default named `cover-letter-template.docx` but can be changed later)

### Single Application Generation

For generating a single application (e.g. a single role for a single comapny)

- Use `python cover-gen.py [-name NAME] [--company COMPANY] [--role ROLE] [--name NAME] [--template TEMPLATE] [--folder FOLDER] [--pdf PDF]`

With arguments:

1. `-name` the name of the user, you
2. `--company` the name of the company applying for
3. `--role` the role applying for
4. `--event` (Optional) event to mention in the application, e.g. networking event, company social
5. `--other` (Optional) other information as found pertinent to the application
6. `--template` the name of the template to be modified (defaults to `cover-letter-template.docx`)
7. `--folder` Whether or not for the outputted `.pdf` or `.docx` file to be placed in a subfolder with the name of the associated company
8. `--pdf` whether or not to output a `.pdf` or `.docx` file

### Multi Application Generation

For generating multiple applications (i.e. multiple roles from multiple companies)

- Use `python cover-gen.py [-name NAME] [--template TEMPLATE] [--app_list APP_LIST] [--pdf PDF]`

With arguments:

1. `-name` the name of the user, you
2. `--template` the name of the template to be modified (defaults to `cover-letter-template.docx`)
3. `--app_list` a `.xlsx` or `.csv` file in the format of having columns of `role` and `company`, with optional columns of `event` and `other` (as specified above)

### Template

Within the template (a `.docx` document), the script effectively replaces all dates, companies, roles, events, and "other" item found with the given format change:

- `{{DATE}}` -> `datetime.date` in the format `BB dd, YYYY` (e.g. May 28, 2023)
- `{{COMAPNY}}` -> `--company` if a singular entry or the row number for a given `company`
- `{{ROLE}}` -> `--role` if a singular entry or the row associated with a given `role`
- `{{EVENT}}` -> `--event` if a singular entry or the row associated with a given `event`
- `{{OTHER}}` -> `--other` if a singular entry or the row associated with a given `other`

As such, in the Word `.docx` document, change each mention of a date, company, role, event, and "other" item accordingly, please take a peek at the given sample cover letter `cover-letter-template.docx`, but an example would be "May 28, 2023" -> "{{DATE}}" in the `.docx` (Microsoft Word) document

## Download and Implementation

Simply either download the repository as a `.zip` file or clone it to GitHub Desktop, follow the installation instructions above, and to test change the directory to the folder and run the commands:

```unix
python cover-gen.py -name "First Last" --template cover-letter-template.docx --app_list test_file.csv
```

If this works, replace the template with your own cover letter and list of companies to apply to with your custom list as well and all should run smoothly

Good luck applying :)

## Future Features

At the moment, attempting to implement two features and one potential API integration:

- [ ] Automatically downloading dependencies - oops, no idea how to really work this HAHaHAhahaHaha... :( pssst for now, please just use the `requirements.txt` and maybe learn how to read?
- [x] The generation of multiple cover letters at once through reading in a `.xslx` or `.csv` file containing company and roles
- [x] The generation of cover letters for an `{{EVENT}}` flag, indicating any events attended by the user
- [x] The generation of cover letters for an `{{OTHER}}` flag, indicating other, wildcard options the user would like to fill
- [ ] The generation of cover letters with dates other than the day the script was run
- [ ] The integration of Open AI GPT API to customize sections of cover letters
