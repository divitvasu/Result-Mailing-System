######################################
##########IMPORTING MODULES###########
######################################

import xlrd
import pdfkit
import sys,os
import easygui


######################################
###########DISABLE OUTPUT#############
######################################

sys_original=sys.stdout
def blockPrint():
    sys.stdout = open(os.devnull, 'w')

def enablePrint():
    sys.stdout = sys_original



######################################
######### READING EXCEL FILE##########
######################################

data_file=easygui.fileopenbox(msg="Select Marksheet file: ")

try:
    b=xlrd.open_workbook(data_file)
    b1=xlrd.open_workbook('EmailList.xls')
except FileNotFoundError:
    print("EmailList.xls not found!!")
    import PROJECT
except TypeError as te:
    print("Error: ",te)
    import PROJECT

    

######################################
############OPENING SHEETS############
######################################

sheet1 = b.sheet_by_name('Sheet1')
sheet2 = b.sheet_by_name('Sheet2')
sheet3 = b1.sheet_by_name('Sheet1')


######################################
#####VARIABLES DECLARED GLOBALLY######
######################################

counter=0
d1={}
subname=[]
sb=[]
code=[]
columns={}
i=0
headers=[]



######################################
###### ABSOLUTE PATH FOR IMAGES#######
######################################

image1=os.path.abspath("logo.jpg")
image2=os.path.abspath("title.jpg")


######################################
###########FETCHING HEADERS###########
######################################

for col in range(sheet2.ncols):
    data=sheet2.cell(0,col).value
    headers.append(data)

for key in headers:
    columns[key]=i
    i+=1

######################################
######FETCHING BASIC INFORMATION######
######################################

def getDetails():
    global sem,branch,exam,programme,institutename,totalsub
    sem=sheet1.cell(3,1).value
    branch=sheet1.cell(2,1).value
    exam=sheet1.cell(4,1).value
    programme=sheet1.cell(1,1).value
    institutename=sheet1.cell(0,1).value
    totalsub=int(sheet1.cell(5,1).value)
    
    for sn in d1:
        if((sn!=list(columns.keys())[0]) and (sn!=list(columns.keys())[1])):
            subname.append(sn)

    for k in range(len(subname)):
        for i in range(6,6+totalsub):
            if(sheet1.cell(i,0).value==subname[k]):
                sb.append(sheet1.cell(i,1).value)
                code.append(sheet1.cell(i,2).value)

######################################
##########PDF Generation##############
######################################
def printToPDF(eno,html_str):
    blockPrint()
    output_filename = eno+'.pdf'
    path_wkthmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
    try:
        pdfkit.from_string(html_str,output_filename,configuration=config)
    except OSError:
        enablePrint()
        print("Please close "+output_filename+" file!")
        input("Press any key to continue!")
        printToPDF(eno,html_str)
    enablePrint()




######################################
###########FETCHING RESULT############
######################################

def generateResult(item):
    global counter
    counter+=1
    #print(columns)
    n=0
    
    for cell in sheet2.col(0,start_rowx=0,end_rowx=sheet2.nrows):
        if (cell.value==item):
            #print('val'+str(item))
            break
        else:
            n=n+1

    
    for row in range(1,sheet2.nrows):
        
        for col in range(0,len(headers)):
            if(n==row):
                d1[headers[col]]=sheet2.cell(row,col).value


    #print(d1)
    name=d1[list(columns.keys())[1]]
    eno=str(int(d1[list(columns.keys())[0]]))
    if counter==1:
        getDetails()
#code=['0','1','2','3','4','5','6','&nbsp;']
#sb=['sb0','sb1','sb2','sb3','sb4','sb5','sb6','&nbsp;']

    marks=[]
    marksobtained=0
    totalsubjects=0
    for i in range(8):
        
        if((i+1)>len(subname)):
            marks.append('&nbsp;')
            sb.append('&nbsp;')
            code.append('&nbsp;')
        elif(d1[subname[i]]=='ABSENT'):
            marks.append('ABSENT')
            totalsubjects+=1
        else:
            marks.append(str(int(d1[subname[i]])))
            marksobtained+=int(d1[subname[i]])
            totalsubjects+=1

    totalmarks=totalsubjects*(30.00)
    percentage=(marksobtained/totalmarks)*100
    if(percentage>=40.00):
        result='PASS'
    else:
        result='FAIL'
    marksobtained=str(marksobtained)
    percentage='{:.2f}'.format((percentage))


######################################
########RESULT IN HTML FORMAT#########
######################################
            
    html_str=r"""
<html>
<head>
<meta name="pdfkit-orientation" content="Landscape"/>
    <style type="text/css">
        .ace-icon{text-align:center}
        .text-primary{color:#337ab7}.bigger-120{font-size:120%!important}

         
        center {
            vertical-align:center;
            horizontal-align:center;
            margin: auto;
            width: 70%;
            padding: 10px;
        }
 
        .mainTbl
        {
            margin: 5px 5px 5px 5px;
            border: Solid 2px #ba373d;
        }
        td
        {
            font-family: Verdana, Arial, Helvetica, sans-serif;
            font-size: 11pt;
            font-weight: bold;
            text-align: left;
            border: Solid 1px #ba373d;
        }
        .style2
        {
            font-size: 19pt;
            color: #BA393E;
            text-align: center;
            width: auto;
        }
        .HeaderDiv
        {
            
            margin-left: 5px;
        }
        .HdrCell
        {
            border: Solid 0px;
            text-align: center;
        }
        .style16
        {
            color: #000000;
            font-weight: bold;
            font-size: 11pt;
        }
        .style18
        {
            color: #000033;
            font-weight: bold;
            font-size: 11pt;
        }
        .style19
        {
            font-size: 11pt;
            font-weight: bold;
        }
        .style21
        {
            color: #003399;
            font-weight: bold;
            text-align: center;
        }
        .lbl
        {
            text-align: center;
            float: left;
            width: 100%;
        }
        .grd
        {
            color: #003399;
            font-weight: bold;
            font-family: Verdana, Arial, Helvetica, sans-serif;
            font-size: 9.5pt;
            border: Solid 2px #19AA8A;
            text-align: center;
        }
        .GrdCell
        {
            height: 17px;
            text-align: center;
        }
         .footer .footer-inner .footer-content
        {
            position: absolute;
            left: 0px;
            right: 0px;
            bottom: 0px;
            padding: 8px;
            border-top: 3px double #E5E5E5;
            height:30px;
        }
        .footer-inner
        {
            text-align: center;
            bottom: 0;
        }
        .footer-content
        {
            background-color: #ba373d;
            color: White;
        }
    </style>
    
</head>
<body>
<center>
<br>
<br>

<form id="form1" action="./" method="post">
<div align="center">
<div id="dvgrdgrid" style="text-align: left; display: block;">
<div style="width: 100%;">
<table width="100%">
<tbody>
<tr>
<td style="border: None 0px; width: 100%;"><img id="uclGrd_ImgLogo1" style="height: 107px; width: 100px;" src='"""+image1+"""' alt="" /> <img id="uclGrd_Image1" style="height: 107px; width: 86%;" src='"""+image2+"""' alt="" /></td>
</tr>
</tbody>
</table>
<br />
<br>
<div style="text-align: center; width: 100%; font-size: 30px; color: #1b6aaa;">"""+exam+""" Result</div>
<br />
<br>
<table border="1" width="100%" cellspacing="0" cellpadding="3">
<tbody>
<tr>
<td class="GrdCell" bgcolor="#CCCCCC" width="15%"><span class="style21">Programme</span></td>
<td colspan="4"><span id="uclGrd_lblProgramme">"""+programme+"""</span></td>
</tr>
<tr>
<td class="GrdCell" bgcolor="#CCCCCC" width="15%"><span class="style21">Name of Institute </span></td>
<td colspan="4"><span id="uclGrd_lblInstituteName">"""+institutename+"""</span></td>
</tr>
<tr bgcolor="#CCCCCC">
<td class="GrdCell" width="15%"><span class="style21">Enrollment No. </span></td>
<td class="GrdCell" width="34%"><span class="style21">Name</span></td>
<td class="GrdCell" width="21%"><span class="style21">Branch</span></td>
<td class="GrdCell" width="12%"><span class="style21">Semester</span></td>
<td class="GrdCell" width="18%"><span class="style21">Mid Sem</span></td>
</tr>
<tr>
<td width="15%"><span id="uclGrd_lblExamNo" class="lbl">"""+eno+"""</span></td>
<td width="34%"><span id="uclGrd_lblStudentName" class="lbl">"""+name+"""</span></td>
<td width="21%"><span id="uclGrd_lblDegreeName" class="lbl">"""+branch+"""</span></td>
<td width="12%"><span id="uclGrd_lblSemester" class="lbl">"""+sem+"""</span></td>
<td width="18%"><span id="uclGrd_lblMnthYr" class="lbl">"""+exam+"""</span></td>
</tr>
</tbody>
</table>
<br />
<br>
<br>
<div>
<table id="uclGrd_grdResult" class="grd" style="border-color: #BA373D; border-collapse: collapse;" border="1" rules="all" cellspacing="0">
<tbody>
<tr>
<th style="background-color: #cccccc; height: 30px; width: 15%; white-space: nowrap;" scope="col" align="center">Course Code</th>
<th style="background-color: #cccccc; width: 45%; white-space: nowrap;" scope="col" align="center">Course Title</th>
<th style="background-color: #cccccc; height: 30px; width: 10%;" scope="col" align="center">Marks Obtained (Max 30)</th>
</tr>
<tr align="center">
<td class="GrdCell" align="center" valign="top">"""+code[0]+"""</td>
<td valign="top">"""+sb[0]+"""</td>
<td class="GrdCell" align="center" valign="top">"""+marks[0]+"""</td>
</tr>
<tr align="center">
<td class="GrdCell" align="center" valign="top">"""+code[1]+"""</td>
<td valign="top">"""+sb[1]+"""</td>
<td class="GrdCell" align="center" valign="top">"""+marks[1]+"""</td>
</tr>
<tr align="center">
<td class="GrdCell" align="center" valign="top">"""+code[2]+"""</td>
<td valign="top">"""+sb[2]+"""</td>
<td class="GrdCell" align="center" valign="top">"""+marks[2]+"""</td>
</tr>
<tr align="center">
<td class="GrdCell" align="center" valign="top">"""+code[3]+"""</td>
<td valign="top">"""+sb[3]+"""</td>
<td class="GrdCell" align="center" valign="top">"""+marks[3]+"""</td>
</tr>
<tr align="center">
<td class="GrdCell" align="center" valign="top">"""+code[4]+"""</td>
<td valign="top">"""+sb[4]+"""</td>
<td class="GrdCell" align="center" valign="top">"""+marks[4]+"""</td>
</tr>
<tr align="center">
<td class="GrdCell" align="center" valign="top">"""+code[5]+"""</td>
<td valign="top">"""+sb[5]+"""</td>
<td class="GrdCell" align="center" valign="top">"""+marks[5]+"""</td>
</tr>
<tr align="center">
<td class="GrdCell" align="center" valign="top">"""+code[6]+"""</td>
<td valign="top">"""+sb[6]+"""</td>
<td class="GrdCell" align="center" valign="top">"""+marks[6]+"""</td>
</tr>
<tr align="center">
<td class="GrdCell" align="center" valign="top">"""+code[7]+"""</td>
<td valign="top">"""+sb[7]+"""</td>
<td class="GrdCell" align="center" valign="top">"""+marks[7]+"""</td>
</tr>
</tbody>
</table>
</div>
<br>
<br />
<table border="0" width="100%">
<tbody>
<tr align="center">
<td style="border: none 0px;">
<table style="height: 43px;" border="1" width="75%" cellspacing="0" cellpadding="3" align="center">
<tbody>
<tr bgcolor="#CCCCCC">
<td class="GrdCell" style="width: 33%;">Total Marks&nbsp;</td>
<td class="GrdCell" style="width: 33%;">Percentage</td>
<td class="GrdCell" style="width: 33%;">Result</td>
</tr>
<tr>
<td style="width: 33%;"><span id="uclGrd_lblCrdOffered" class="lbl">"""+marksobtained+"""</span></td>
<td style="width: 33%;"><span id="uclGrd_lblGrdPtEarned" class="lbl">"""+percentage+"""%</span></td>
<td style="width: 33%;"><span id="uclGrd_lblGrdPtEarned" class="lbl">"""+result+"""</span></td>
</tr>
</tbody>
</table>
</td>
<td style="border: none 0px;" rowspan="3" valign="top">&nbsp;</td>
</tr>
<tr>
<td style="border: none 0px;" rowspan="3" align="center">&nbsp;</td>
</tr>
</tbody>
</table>
<span id="uclGrd_lblTransferRem" style="font-weight: bold; display: none;"></span>
<p style="text-align: left;">&nbsp;</p>
</div>
</div>
</div>
</form>

</center>
</body>
</html>
"""

    printToPDF(eno,html_str)
    print("Successfully created "+eno+".pdf")


######################################
##########GENERATING RESULT###########
######################################

for r in range(sheet3.nrows-1):
    generateResult(sheet3.cell(r+1,0).value)
    

b.release_resources()
del b
b1.release_resources()
del b1
