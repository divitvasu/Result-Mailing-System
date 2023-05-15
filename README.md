# Result-Mailing-System (2018)

> Tools - Python, HTML

## About

The project focused on formulating a functional "Result Mailing System", and aimed to provide a convenient and hassle-free process for preparing the PDF files of the students' results from Excel files and mailing them to the respective students/guardians. This approach reduced the need for manual creation of PDFs and individual mailing. This project prepares a PDF file for individual students from these excel files, fetches the email ids of the respective students and mails the PDF files to them. Thus, this system aims at offering ease of access and convenience.

## Libraries Used

-	*XLRD*- This module is used to read data from an excel workbook.
-	*PDFKIT*- This module is used to convert an html code into its pdf counterpart.
-	*WKHTMLTOPDF*- This is a free-to-download software that is required by the module PDFKIT to work and generate the desired PDF.
-	*SMTPLIB*- This module is required for facilitating the mail services.
-	*EASYGUI*- This module is required for facilitating the GUI.

## Execution

- Install the required dependencies: 
`pip install xlrd`  `pip install pdfkit` `pip install easygui` <br>
*wkhtmltopdf* (https://wkhtmltopdf.org/)

- Include a `logo.png` and a `title.png` for the header of the pdf. These have been omitted from the upload, as these files are specific to the end-user/institute. *The code would not work as-is, without the presence of these two files in the same directory. Else, slight modifications would be needed to the code.*

- The *XLRD* module reads the data from the given excel sheet. The excel workbooks have to be in a specific format. Prepare the required excel workbooks *Marksheet.xls* (*Sheet1*, *Sheet2*) and *EmailList.xls* (*Sheet1*). Do note, *XLRD* might not work with newer .xlsx format, use .xls instead to avoid issues.

- *Marksheet.xls* *Sheet1* template
<p align="center">
<img src="https://github.com/divitvasu/Result-Mailing-System/assets/30820920/ca671516-ca71-450c-8a8e-ff3d4f44c0c7" alt="Image" width="700" height="250">
</p>

- *Marksheet.xls* *Sheet2* template
<p align="center">
<img src="https://github.com/divitvasu/Result-Mailing-System/assets/30820920/aee90d5f-f861-439f-a03f-e386193641e1" alt="Image" width="500" height="100">
</p>

- *EmailList.xls* *Sheet1* template
<p align="center">
<img src="https://github.com/divitvasu/Result-Mailing-System/assets/30820920/43611a6c-29cc-43b2-a133-60186c655792" alt="Image" width="400" height="100">
</p>

- Once the python code is run, the *PDFKIT* module converts the embedded html code template (defining mark sheet structure) into a pdf template. Once the PDF template is created, data is written into it according to the formats specified in *Marksheet.xls* *Sheet1*. The location of the resultant PDF is same as the location of the executing python code.
<p align="center">
<img src="https://github.com/divitvasu/Result-Mailing-System/assets/30820920/844b5c28-e53f-4f83-bd17-6726b0e3f14d" alt="Image" width="700" height="500">
</p>

- Once the pdf is ready it can be emailed to the recepients in *EmailList.xls* *Sheet1* using the *SendEmail.py* file as shown below. The created PDF file, is then attached into a mail with the desired Subject and message lines. This mail is sent using the *SMTPLIB* module, which uses the local SMTP server to send the mail. The recipients of the mail are fetched from *EmailList.xls* *Sheet1*.
<p align="center">
<img src="https://github.com/divitvasu/Result-Mailing-System/assets/30820920/e6ae1592-c2d1-49f0-80d4-f85019ed82ea" alt="Image" width="400" height="300">
</p>
 
## References

* [Generate PDF from HTML template](https://stackoverflow.com/questions/12194467/python-wkhtmltopdf-to-generate-pdf)
*	[Smtplib for sending emails](https://docs.python.org/2/library/smtplib.html)
*	[Generate PDF files](https://micropyramid.com/blog/how-to-create-pdf-files-in-python-using-pdfkit/)

### Contributors
@jainkrunal
