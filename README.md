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

(Psst if you don't have Python installed, download it [here](https://www.python.org/downloads/) and might I recommend an editor, mine of choice is [Visual Studio Code](https://code.visualstudio.com/))

## Usage

Download and place the script `cl-gen.py` file to the same location as the desired cover letter template `.docx` file

 - Change current directory to the location of the template (by default named `cover-letter-template.docx` but can be changed later)

### Single Application Generation

For generating a single application (e.g. a single role for a single comapny)

- Use `python cover-gen.py [-name NAME] [--company COMPANY] [--role ROLE] [--name NAME] [--template TEMPLATE] [--folder FOLDER] [--pdf PDF]`

With arguments:

1. `-name` the name of the user, you
2. `--date` the date of the application, if different from today
3. `--company` the name of the company applying for
4. `--address` the address of the company you are applying for
5. `--role` the role applying for
6. `--event` event to mention in the application, e.g. networking event, company social
7. `--contact` the name of a contact within the comapny or otherwise
8. `--referral` the name of the person referring the user, you, to the company
9. `--hmanager` the name of the hiring manager (if known)
10. `--convo1`/`--convo2` the contexts for conversations applicable to the application
11. `--other1`/`--other2` other information as found pertinent to the application
12. `--template` the name of the template to be modified (defaults to `cover-letter-template.docx`)
13. `--folder` Whether or not for the outputted `.pdf` or `.docx` file to be placed in a subfolder with the name of the associated company
14. `--pdf` whether or not to output a `.pdf` or `.docx` file

### Multi Application Generation

For generating multiple applications (i.e. multiple roles from multiple companies)

- Use `python cover-gen.py [-name NAME] [--template TEMPLATE] [--app_list APP_LIST] [--pdf PDF]`

With arguments:

1. `-name` the name of the user, you
2. `--template` the name of the template to be modified (defaults to `cover-letter-template.docx`)
3. `--app_list` a `.xlsx` or `.csv` file in the format of having columns of `role` and `company`, with optional columns of `event` and `other` (as specified above)

### Template

Within the template (a `.docx` document), the script effectively replaces all dates, companies, roles, events, contacts, referrers, hiring managers, conversations and "other" items found with the given format change:

- `{{NAME}}` -> `--name` in the format of `First Last` name
- `{{DATE}}` -> `--date` in a generally accepted date format (e.g. `BB dd, YYYY`; `May 28, 2023`) if a singular entry or the row number for a given `date` column if from a `.csv` or `.xlsx` or today's date (as provided by `datetime.date.today()`) if none provided
- `{{COMPANY}}` -> `--company` if a singular entry or the row number for a given `company` column if from a `.csv` or `.xlsx`
- `{{ADDRESS}}` -> `--address` if a singular entry or the row associated with a given `address` column if from a `.csv` or `.xlsx`
    - This is separated by at least 2 commas, e.g. "1234 Meridian Lane, New York, NY 10004"
- `{{ROLE}}` -> `--role` if a singular entry or the row associated with a given `role` column if from a `.csv` or `.xlsx`
- `{{EVENT}}` -> `--event` if a singular entry or the row associated with a given `event` column if from a `.csv` or `.xlsx`
- `{{CONTACT}}` -> `--contact` if a singular entry or the row associated with the given `contact` column if from a `.csv` or `.xlsx`
- `{{REFERRAL}}` -->
- `{{HMANAGER}}`
- `{{CONVO1}}`/`{{CONVO2}}` -> `--convo1`/`--convo2`
- `{{OTHER1}}`/`{{OTHER2}}` -> `--other1`/`--other2` if a singular entry or the row associated with a given `other`

As such, in the Word `.docx` document, change each mention of a date, company, role, event, and "other" item accordingly, please take a peek at the given sample cover letter `cover-letter-template.docx` (courtesy of ChatGPT), but an example would be "May 28, 2023" -> "{{DATE}}" in the `.docx` (Microsoft Word) document

## Download and Implementation

Simply either download the repository as a `.zip` file or clone it to GitHub Desktop, follow the installation instructions above, and to test change the directory to the folder and run the commands:

```unix
python cover-gen.py -name "First Last" --template cover-letter-template.docx --app_list test_file.csv
```

If this works, replace the template with your own cover letter and list of companies to apply to with your custom list as well and all should run smoothly

Good luck applying :)

## Future Updates

At the moment, attempting to implement two features and one potential API integration:

- [ ] Automatically downloading dependencies - oops, no idea how to really work this HAHaHAhahaHaha... :( pssst for now, please just use the `requirements.txt`?
- [x] The generation of multiple cover letters at once through reading in a `.xslx` or `.csv` file containing company and roles (v2.0.0)
- [x] An `{{EVENT}}` flag, indicating any events attended by the user (v2.0.0)
- [x] An `{{OTHER}}` flag, indicating other, wildcard options the user would like to fill (v2.0.0)
- [x] A `{{HMANAGER}}` flag, indicating references to specific hiring managers (v3.0.0)
- [x] A `{{CONTACT}}` flag, indicatinga specific contact's name (v3.0.0)
- [x] An `{{ADDRESS}}` flag, with individual company addresses (v3.0.0)
- [x] A `{{DATE}}` flag, with specific dates that do not have to be today (v3.0.0)
- [x] The addition of one more `{{OTHER}}` flag (v3.0.0)
- [x] The addition of two `{{CONVO}}` flags, indicating interesting pieces of conversation to include in the cover letter (v3.0.0)
- [ ] The integration of a feature to output the number of errors for each type (e.g. 3 address errors/4 date format errors)
- [ ] The integration of a feature to output the companies/roles/applications associated with each error
- [ ] The integration of Open AI Chat/GPT API to customize sections of cover letters

## Shoutout and Thanks

Thanks to [TextKool](https://www.textkool.com/en) for its [ASCII Art Generator](https://textkool.com/en/ascii-art-generator?hl=default&vl=default&font=ANSI%20Regular&text=cover-gen%0Av3.0.0)
