# jpm - The Job Package Manager

Create 100's of application packages in seconds. 

## What you need
* Your template cover letter - docx file
This word file must have  placeholder variables like: `{{ company_name }}` or `{{ job_title }}

* Data you want to fill your cover letter with - CSV file
You need to provide a CSV containing columns associated with each placeholder variable in the template and each row associated with a new job package as such:
| company_name | job_title         | ... |
|--------------|-------------------|-----|
| Facebook     | Software Engineer | ... |
| Google       | Data Science      | ... |
| ...          | ...               | ... |

* Extra files to append to package - PDF files
This could be files like your resume, your transcript, etc. Static files that are to be added to the end of the package in alphabetical order. (So make sure to name them appropriately i.e. prepend 'a_', 'b_', etc.)

* Python
* Windows Machine 
Currently, jpm only supports Microsoft Word templates so may only run on windows machines. Linux version with support for LibreOffice (ODT format) coming soon.

## Quick start

First let's get the repo onto your machine.
```bash
$ cd ~
$ git clone 
$ cd jpm
$ pip install -r requirements.txt
```

Now, move all following files to `./config`:
* Your cover letter 
* CSV of jobs data
* Extra PDF files to append (optional)

Finally, let's make all the packages.
```bash
$ cd ~/jpm
$ python main.py
```

You can now find all final PDF packages in `./packages`. Along the way, cover letters in docx format are stored in `./docs` and pdf version of the same are stored in `./pdfs`.

To reset the repo of all the files just created (this will not touch your `./config`), run the following from the jpm folder:
```bash
$ python reset.py
```
