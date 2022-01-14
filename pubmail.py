import win32com.client as win32
import mammoth
import os
import pandas as pd
import numpy as np
from more_itertools import sort_together

def readwos(infile,cols):
    '''
    Read input excel files exported from Web of Science, and return relevant
    concatenated columns (cols) in a dictionary of numpy arrays
    '''
    wos = {}
    for i in range(len(cols)):
        col = np.array([])
        for j in range(len(infile)):
            df = pd.read_excel(infile[j])
            col = np.append(col,np.array(df[cols[i]]))
        wos[cols[i]] = np.array(col)
    return wos

def readcr(infile,cols):
    '''
    Read input excel files exported from Web of Science Citation Report, and return relevant
    concatenated columns (cols) in a dictionary of numpy arrays; Also return a separate dictionary
    with number of citations for each available year
    '''
    cr = {}
    for i in range(len(cols)):
        col = np.array([])
        for j in range(len(infile)):
            df = pd.read_excel(infile[j],skiprows=10)
            col = np.append(col,np.array(df[cols[i]]))
        cr[cols[i]] = np.array(col)

    cryear = {}
    for year in list(range(1980,2030)):
        colyear = np.array([])
        for j in range(len(infile)):
            df = pd.read_excel(infile[j],skiprows=10)
            if year in df.columns:
                colyear = np.append(colyear, np.array(df[year]))
        if len(colyear) > 0:
            cryear[year] = np.array(colyear)

    return cr,cryear

def adddoctype(cr,wos,cols):
    '''
    Add columns to the cr dictionary by matching with DOIs of wos dictionary
    '''
    for col in cols:
        dctypes = np.empty(len(cr['DOI']),dtype=object)
        for i in range(len(cr['DOI'])):
            index = np.where(wos['DOI'] == str(cr['DOI'][i]))
            dctypes[i] = str(wos[col][index][0])
        cr[col] = dctypes
    return cr

def sortdict(dic,refkey,reverse=True):
    '''
    Sort all values of a dictionary based on a reference key
    '''
    newdict = {}
    for keys,value in dic.items():
        if reverse:
            newdict[keys] = sort_together([dic[refkey],dic[keys]],reverse=True)[1]
        else:
            newdict[keys] = sort_together([dic[refkey],dic[keys]])[1]
    return newdict

def strsplit(string,key):
    list = str(string).split(key)
    return list,len(list)

def autname(name,num):
    '''
    Format author names for email intro
    '''
    if num == 1:
        auth = 'Dr. %s'%(name[0].split(',')[0].strip())
    elif num == 2:
        auth = 'Dr. %s and Dr. %s'%(name[0].split(',')[0].strip(),name[1].split(',')[0].strip())
    elif num == 3:
        auth = 'Dr. %s, Dr. %s, and Dr. %s'%(name[0].split(',')[0].strip(),name[1].split(',')[0].strip(),name[2].split(',')[0].strip())
    elif num > 3:
        auth = 'Dr %s, Dr. %s, Dr. %s, and co-authors'%(name[0].split(',')[0].strip(),name[1].split(',')[0].strip(),name[2].split(',')[0].strip())
    return auth

def email(template_file,fromaddress,toaddress,subject,placeholders,values,
          autosend=False,saveemail=True,fileout='email.msg',savehtml=True):
    '''
    Script to read a template email, replace placeholder expressions for their values, and save .msg and/or send
    '''
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = fromaddress
    mail.To = toaddress
    mail.Subject = subject
    mail.Body = 'Message body'
    f = open(template_file, 'rb')
    document = mammoth.convert_to_html(f)
    if savehtml:
        b = open('mail_in_HTML.txt', 'wb')
        b.write(document.value.encode('utf8'))
        b.close()
    f.close()

    docstring = document.value

    # Replacing placeholders by assigned values
    for i in range(len(placeholders)):
        docstring = docstring.replace(placeholders[i],values[i])

    mail.HTMLBody = docstring
    if saveemail:
        mail.saveas(os.path.dirname(__file__)+'\\'+fileout) # .msg file to be saved in the same directory of .py
    if autosend:
        mail.Send() # Use this to send email directly - CAREFUL!

# Read Web of Science record of publications
tab = ['wos.xls']
exfile_cols = ['Authors','Article Title','Document Type','Author Keywords','Keywords Plus',
               'Reprint Addresses','Email Addresses','Times Cited, All Databases',
               'Publication Year','Number of Pages','Open Access Designations','DOI']
wos = readwos(tab,exfile_cols)

# Read citation report file(s) - publications must correspond to the Web of Science files
crfile = ['wos_citrep.xls']
# Columns of crfile(s) to read, in addition to all year columns
crfile_cols = ['DOI','Total Citations','Average per Year']
cr,cryear = readcr(crfile,crfile_cols)
# Add columns to the cr dictionary by matching with DOIs of wos dictionary
cr = adddoctype(cr,wos,['Authors','Article Title','Email Addresses','Publication Year'])
# Sort cr list to show highest "Average per year" articles first, in order decreasing
cr = sortdict(cr,'Average per Year',reverse=True)

# Maximum number of emails
maxmail = 50
subject = 'End of year message to Ap&SS esteemed authors'
template_email= 'mail_template.docx'
placeholders = ["!AUTHOR_NAMES!","!PAPER_YEAR!","!PAPER_TITLE!","!CITPERYEAR!","!TOTCIT!"]

# Create one email for each person on the list
for i in range(maxmail):
    # Selection author's names
    namelist,lenaut = strsplit(cr['Authors'][i],';')
    auth = autname(namelist,lenaut)
    # Prepare email
    placeh_replace = [auth,str(int(float(cr['Publication Year'][i]))),cr['Article Title'][i],
                      str(cr['Average per Year'][i]),str(int(cr['Total Citations'][i]))]
    print(placeh_replace)
    print(cr['Email Addresses'][i])
    email(template_email,'fabio.p.santos@springernature.com',cr['Email Addresses'][i],subject,
        placeholders,placeh_replace,
        autosend=False,saveemail=True,
        fileout='EMAIL_%i_%s_%s.msg'%(i+1,cr['Authors'][i][0:20],str(cr['Publication Year'][i])))

