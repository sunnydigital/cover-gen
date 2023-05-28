# Cover Letter Generator

---

Repository for a simple Python-baed command line cover letter generator using a base docx template

## Usage

---

 - Change current directory to the location of the template (by default named `"Cover-Letter-Template.docx"`
 - Use `python cv-convert.py --company [COMPANY] --role [ROLE] --name [NAME] --template [TEMPLATE] --folder [FOLDER] --pdf [PDF]`

With specifications of the arguments as follows:

1. `--company` the name of the company applying for
2. `--role` the role applying for
3. `--template` the name of the template to be modified
4. `--folder` the name of the subfolder for the outputted `.pdf` or `.docx` file to be placed in
5. `--pdf` whether or not to output a `.pdf` or `.docx` file

In terms, this script takes a basic template and simply outputs a specific cover letter for each company, for the addressed date, for a given role 
