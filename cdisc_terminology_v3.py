__name__        = 'cdisc_terminology_v4.py'
__location__    = '\\samfsvr01\samshared\CDISC\terminology\prog'
__purpose__     = '''Download and update CDISC terminology files from: 
                     https://evs.nci.nih.gov/ftp1/CDISC/'''
__developer__   = 'Melanie Logan'
__version__     = '1.0'
__notes__       = '''if files require additional time for download, time.sleep(n)
                     where n = seconds'''
#=========================================================================
# {modules}
#=========================================================================
import logging, os, shutil, itertools, time, pyxlsb, subprocess, smtplib
from datetime import date
from email.message import EmailMessage

from selenium import webdriver
browser = webdriver.Chrome(r'C:\Users\mlogan\Downloads\chromedriver.exe')
browser.get('https://evs.nci.nih.gov/ftp1/CDISC/')
browser.implicitly_wait(2)


#=========================================================================
# {directories}
#=========================================================================
# set terminology path
#termdir = '\\\\samfsvr01\\SAMSHARED\\CDISC\\terminology'

# set archive path
#archdir = '\\\\samfsvr01\\SAMSHARED\CDISC\\terminology\\retire'

# set changes path
#chgdir = '\\\\samfsvr01\\SAMSHARED\\CDISC\\terminology\\changes'

# set downloads path
downdir = 'C:\\Users\\mlogan\\Downloads'



# set terminology path
termdir = '\\\\samfsvr01\\samshared\\CDISC\\terminology\\testing\\test - melanie'

# set archive path
archdir = '\\\\samfsvr01\\samshared\\CDISC\\terminology\\testing\\retire - melanie'

# set changes path
chgdir = '\\\\samfsvr01\\samshared\\CDISC\\terminology\\testing\\changes - melanie'

#=========================================================================
# {lists}
#=========================================================================
# elements
elmADaM     =   ('ADaM/','ADaM Terminology.xls','ADaM Terminology Changes.xls')
elmSEND     =   ('SEND/','SEND Terminology.xls','SEND Terminology Changes.xls')
elmPTL      =   ('Protocol/','Protocol Terminology.xls','Protocol Terminology Changes.xls')
elmXML      =   ('Define-XML/','Define-XML Terminology.xls','Define-XML Terminology Changes.xls')
elmSDTM     =   ('SDTM/','SDTM Terminology.xls','CDASH Terminology.xls','SDTM Terminology Changes.xls','CDASH Terminology Changes.xls')

FileList    =   ('ADaM Terminology','CDASH Terminology','Define-XML Terminology', 
                 'Protocol Terminology','SDTM Terminology','SEND Terminology',
                 'ADaM Terminology Changes','CDASH Terminology Changes','Define-XML Terminology Changes', 
                 'Protocol Terminology Changes','SDTM Terminology Changes','SEND Terminology Changes')

#=========================================================================
# {doawnload .xls files}
#=========================================================================   
# define download function
def download(List):
  for i in List:
    try:
        elem = browser.find_element_by_link_text(i)
        print('Found <%s> element with that class name!' % (elem.text))
    except:
        print('Was not able to find an element with that name.')
 
    # opens link stored in elem
    elem.click()
    
# pass elements
download(elmSDTM)

browser.back()

download(elmADaM)

browser.back()

download(elmSEND)

browser.back()

download(elmPTL)

browser.back()

download(elmXML)

# + 80s delay
time.sleep(80)

# close chrome
browser.quit()

#=========================================================================
# {rename existing + archive}
#=========================================================================
# change working directory
os.getcwd()

# terminology dir
os.chdir(termdir)

ext = '.xlsb'

# rename files, append date
for i in itertools.islice(FileList, 0, 6, None):
    
    # open file
    wb = pyxlsb.open_workbook(i + ext)
    
    # get sheet name (i.e. second element of tuple index=1)
    Sheet = wb.sheets[1]
    print(Sheet)
    
    # split string
    WordList = Sheet.split()
    print(WordList)
    
    # get date
    Lstdate = WordList[-1]
    print(Lstdate)
    
    # close file
    wb.close()
    
    # store current file name into variable
    fname = "".join(i + ext)
    print(fname)
    
    # rename .xlsb file
    os.rename(fname, i + ' ' + Lstdate + ext)
    
    # store new file name into variable
    NewFile = i + ' ' + Lstdate + ext
    print(NewFile)

    # if exisiting, delete from archive and replace 
    try:
        os.remove(archdir + f'\{NewFile}')
    except WindowsError:
        pass

    # move file from terminology folder to archive
    shutil.move(NewFile, archdir)
    
#=========================================================================
# {convert to binary / copy files}
#=========================================================================

# change working directory
os.chdir(downdir)

# define extension types
ext = '.xls'
extSave = '.xlsb'
              
# convert .xls to binary / transfer new terminology files
for i in itertools.islice(FileList, 0, 6, None):
    #global psPath
    
    psPath = f'{downdir}' + f'\{i}'
    print (psPath)
    
    # define ps scripts
    def run(cmd):
        completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
        #return error code
        return completed
    
    
    # ps script: open workbook
    cmd1 = "$xlExcel12 = 50"
    cmd2 = "$Excel = New-Object -Com Excel.Application"
    cmd3 = f"$WorkBook = $Excel.Workbooks.Open('{psPath}{ext}')"

    # ps script: save as .xlsb
    cmd4 = "$Excel.Application.DisplayAlerts=$False"
    cmd5 = f"$WorkBook.SaveAs('{psPath}{extSave}',$xlExcel12,[Type]::Missing,[Type]::Missing,$false,$false,2)"
    cmd6 = "$Excel.Quit()"
    
    
    # run PowersShell scripts
    run(cmd1 + ';' + cmd2 + ';' + cmd3 + ';' + cmd4 + ';' + cmd5 + ';' +cmd6)
 
    
    # copy .xlsb files to terminology folder
    shutil.copy(i + extSave, termdir)


# delete existing 'changes' .xls files + replace
for i in itertools.islice(FileList, 6, None, None):
    try:
        os.remove(chgdir + i + ext)
    except WindowsError:
        pass
    
    shutil.copy(i + ext, chgdir)
    
    
# purge downloads
dFiles = os.listdir(downdir)
print(dFiles)

for item in dFiles:
    if item.endswith(('.xls', '.xlsb')):
        os.remove(os.path.join(downdir, item))

#=========================================================================
# {lEmail}
#=========================================================================

# prep email
sender_email = 'melanie@samumed.com'
receiver_email = ['melanie@samumed.com']
#receiver_email = ['billk@samumed.com']
#receiver_email = ['programming@samumed.com', 'stats@samumed.com']
today = date.today()

msg = EmailMessage()
msg['Subject'] = f'<Automated Message> CDISC Controlled Terminology Download: {today} [Completed]'
msg["From"] = sender_email
msg["To"] = receiver_email
#msg.set_content('Dear all,\n\nCDISC Controlled Terminology Download has been completed.')
msg.add_alternative("""\
<html>
  <body>
  <h1 style="background-color:cornsilk;font-family:helvetica;color:midnightblue;font-size:20px">CDISC Controlled Terminology Download</h1>
    <p>Dear all,<br><br>
        *** TEST EMAIL***<br><br>
       A routine CDISC controlled terminology download has been completed.<br><br>
       <b>The following files have been downloaded:</b>
       <ul >
           <li>ADaM Terminology</li>
           <li>CDASH Terminology</li>
           <li>Define-XML Terminology</li>
           <li>Protocol Terminology</li>
           <li>SDTM Terminology</li>
           <li>SEND Terminology</li>
       </ul>
       The downloaded files can be accessed <a href="\\\\samfsvr01\\samshared\\CDISC\\terminology">HERE</a><br><br>
       Thank you!
    </p>
  </body>
</html>
""", subtype='html')
  
# send email
try:
    smtpObj = smtplib.SMTP('10.20.10.136')
    smtpObj.send_message(msg)
    logging.info('Email sent successfully.')
    #disconnect from SMTP
    smtpObj.quit()
except Exception as e:
    pass
    logging.info(e)
    logging.info('ERROR: Unable to send email.')
                                    
#=========================================================================
# {logger}
#=========================================================================
#TODO Update log to send email once updaate has been performed/flag for .xlsb conversion
# Also, email programming and stats when terminology  have been updated.

# change working directory
#os.chdir('\\\\samfsvr01\\samshared\\CDISC\\terminology\\py_log')

#logging.basicConfig(filename='cdisc_terminology_v2.log', level=logging.CRITICAL, 
                    #format= '%(asctime)s:%(levelname)s:%(message)s')

## End of program ##









