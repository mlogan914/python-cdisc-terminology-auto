__name__        = 'cdisc_terminology.py'
__location__    = r'C:\Users\Melanie_logan\OneDrive - Edwards Lifesciences\Documents\_my dir\projects\CDISC Terminology Downloads'
__purpose__     = '''Download and update CDISC terminology files from: https://evs.nci.nih.gov/ftp1/CDISC/'''
__developer__   = 'Melanie Logan'
__version__     = '1.0'
__notes__       = '''if files require additional time for download, time.sleep(n) where n = seconds'''

#=========================================================================
# {modules}
#=========================================================================
import logging, os, shutil, itertools, time, pyxlsb, subprocess, smtplib

#from pathlib import Path
from datetime import date
from datetime import datetime
logfile = datetime.now().strftime('py_logfile_%Y_%m_%d_%H_%M_%S.log')
DownldDt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# config logger
logging.basicConfig(level=logging.INFO, 
                    filename= rf'N:\\admin\\(SPT) Statistical Programming Team\\Standards\\CDISC\\CDISC Controlled Terminology\\py_log\\{logfile}', 
                    #filemode='w', 
                    format='%(asctime)s [LINE:%(lineno)d] %(levelname)-8s %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    )



from email.message import EmailMessage

from selenium import webdriver
browser = webdriver.Chrome(r'N:\admin\(SPT) Statistical Programming Team\Tools\Python\chromedriver\95_0_4368_54\chromedriver.exe')
browser.get('https://evs.nci.nih.gov/ftp1/CDISC/')
browser.implicitly_wait(2)

#=========================================================================
# {directories}
#=========================================================================
# set downloads path
downdir = 'C:\\Users\\Melanie_Logan\\Downloads'

# set terminology path
termdir = 'N:\\admin\\(SPT) Statistical Programming Team\\Standards\\CDISC\\CDISC Controlled Terminology'

# set archive path
archdir = 'N:\\admin\\(SPT) Statistical Programming Team\\Standards\\CDISC\\CDISC Controlled Terminology\\Archive'

# set changes path
chgdir = 'N:\\admin\\(SPT) Statistical Programming Team\\Standards\\CDISC\\CDISC Controlled Terminology\\Changes'

# set log path
logdir = 'N:\\admin\\(SPT) Statistical Programming Team\\Standards\\CDISC\\CDISC Controlled Terminology\\py_log'

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
        logging.info('Sucsessfullly found element with class name: <%s>' % (elem.text))
    except:
        print('Was not able to find an element with that name.')
        logging.error('Was not able to find an element with class name <%s>.' % (elem.text))
 
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
#TODO: add powershell command to set .xlsb files to READ-ONLY
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
# {readme.txt}
#=========================================================================
outFileName = f'{termdir}' + '\\readme.txt'
outFile = open(outFileName, "w")

outFile.write(rf""" *** NOTE: CDISC Controlled Terminology files were automatically downloaded on {DownldDt}
              
__name__        = {__name__}
__location__    = {__location__}
__purpose__     = {__purpose__}
__developer__   = {__developer__}
__version__     = {__version__}
__notes__       = {__notes__} """)

outFile.close()

#=========================================================================
# {lEmail}
#=========================================================================
# prep email
sender_email = 'Melanie_Logan@edwards.com'
#receiver_email = ['Melanie_Logan@edwards.com']
receiver_email = ['THV_Stats_Programming@edwards.com']
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
    <p>Dear All,<br><br>
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
       The downloaded files can be accessed <a href="N:\\admin\\(SPT) Statistical Programming Team\\Standards\\CDISC\\CDISC Controlled Terminology">HERE</a><br><br>
       The log can be accessed <a href="N:\\admin\\(SPT) Statistical Programming Team\\Standards\CDISC\CDISC Controlled Terminology\\py_log">HERE</a><br><br>
       Thank you
    </p>
  </body>
</html>
""", subtype='html')
  
# send email
try:
    smtpObj = smtplib.SMTP('smtpirv.edwards.lcl') #update
    smtpObj.send_message(msg)
    logging.info('Email sent successfully.')
    #disconnect from SMTP
    smtpObj.quit()
except Exception as e:
    pass
    logging.error(e)
    logging.error('Unable to send email.')
    

### - End of Program Code - ###
