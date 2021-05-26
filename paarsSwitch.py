import wx
import wx.html as wxhtml
import os
import pandas as pd
import numpy as np
import boto3
import botocore
from botocore.exceptions import ClientError
import paramiko
import threading
import time
import datetime
import shutil
from openpyxl import load_workbook
import getpass
import json
import glob
from collections import OrderedDict
import pickle
from wx.lib import masked
from scipy import stats
from random import random
import HTML
# NEWPAARS.py
wildcard = "Excel sheets (*.xls;*.xlsx)|*.xls;*.xlsxi"     
#           "All files (*.*)|*.*"
IPADDRESS = "ec2-3-22-226-74.us-east-2.compute.amazonaws.com"
HOMEDIR = '/home/test/working/'
RESULTSDIR = HOMEDIR + 'results/'
INSTANCE_ID = 'i-03889c587823b7ec5' 
class constants(object):
    pass
K = constants()

if getpass.getuser() == 'phjl': # then we are on amazon
    PCDIRECTORY = 'd:/users/phjl/workdocs/CDALive/'
    ROLLFORWARD = 'd:/users/phjl/workdocs/CDALive/DummyOutputSS.xlsx'
    KEYFILE = 'd:/paars/CDALive/PJEAmazonJupyter.pem'
    K.INDIR = 'd:/Users/Phjl/workdocs/CDALive/'
else:
    PCDIRECTORY = 'w:/My Documents/CDALive/'
    ROLLFORWARD = 'w:/My Documents/CDALive/DummyOutputSS.xlsx'
    KEYFILE = 'w:/My Documents/PJEAmazonJupyter.pem'
    K.INDIR = 'w:/My Documents/CDALive/'

key = paramiko.RSAKey.from_private_key_file(KEYFILE)
client = paramiko.SSHClient()
client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ec2 = boto3.client('ec2')
AppBaseClass = wx.App

aboutText = """<p>Sorry, there is no information about this program. It is
running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>""" 

class MyHtmlFrame(wx.Dialog):
    # Simple HTML message
    def __init__(self,parent,title,content):
        wx.Frame.__init__(self,parent,-1,title)
        htm = wxhtml.HtmlWindow(self) 
        if "gtk2" in wx.PlatformInfo: htm.SetStandardFonts()
        htm.SetPage(content)
        self.SetSize((500,500))
        self.Fit()

def htmlmsg(title,content):
    frm = MyHtmlFrame(None, title,content)
    frm.ShowModal()
    frm.Destroy()

class AboutBox(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, None, -1, "About <<project>>",
            style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|
                wx.TAB_TRAVERSAL)
        hwin = HtmlWindow(self, -1, size=(400,200))
        vers = {}
        vers["python"] = sys.version.split()[0]
        vers["wxpy"] = wx.VERSION_STRING
        hwin.SetPage(aboutText % vers)
        btn = hwin.FindWindowById(wx.ID_OK)
        irep = hwin.GetInternalRepresentation()
        hwin.SetSize((irep.GetWidth()+25, irep.GetHeight()+10))
        self.SetClientSize(hwin.GetSize())
        self.CentreOnParent(wx.BOTH)
        self.SetFocus()



def grow(df): # extract growth rates from cumulative values
    # NB the first column is a zero - this is compliant with the non-cumulative Conning files
    #  And the way the algorithm works, this will give you a 1 as the first year value
    #  when turned back to cumulative.
    num = df.iloc[:,1:]
    den = df.iloc[:,:-1]
    den.columns = num.columns
    growth = num/den - 1
    growth.insert(0,0,0) # insert column zero at position zero with zeros in it!
    return growth

def undermusig(ez, sdz): # gets the underlying mu and sigma for the lognormal
    ez2 = ez * ez
    varz = sdz * sdz
    ex = np.log(ez2/np.sqrt(ez2 + varz))
    sdx = np.sqrt(np.log(1+varz/ez2))
    return ex,sdx

def logn(p,mu,sigma):
    return np.log(stats.lognorm.ppf(p, sigma, loc=0, scale=np.exp(mu))) + 1

def loader(fn, indir,filetype):
    #print(fn)
    dfin = pd.read_csv(indir + fn + '.csv',header=None)
    if filetype == 'Conning':
        return dfin
    else:
        return grow(dfin) # gets the growth rates each period from the cumulative AAA file

# End of ln6
class SummaryResult(wx.Dialog): 
    def __init__(self, parent, htstring): 
        super(SummaryResult, self).__init__(parent, title = 'Summary Results', size = (1100,300)) 
        panel = wx.Panel(self) 
        self.btn = wx.Button(panel, wx.ID_OK, label = "ok") #, size = (50,50)) #, pos = (75,50))
        txt_style = wx.VSCROLL|wx.HSCROLL|wx.TE_READONLY|wx.BORDER_SIMPLE
        htwin = wxhtml.HtmlWindow(panel, -1, size=(1200, 1200), style=txt_style)
        htwin.SetPage(htstring)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(htwin, 0, wx.ALL, 10)
        sizer.Add(self.btn, 0, wx.ALL, 10)
        panel.SetSizer(sizer)
        panel.Layout()

def getamastate():
    alpha = 255
    response = ec2.describe_instances(
        Filters=[],InstanceIds=[INSTANCE_ID],)
    state = response['Reservations'][0]['Instances'][0]['State']['Name'] 
    if state == 'running':
        color = (0,255,0,alpha)
    elif state == 'stopped':
        color = (255,0,0,alpha)
    else:
        color = (255,153,51)
    return state,color

def startamazon():
    # Do a dryrun first to verify permissions
    try:
        ec2.start_instances(InstanceIds=[INSTANCE_ID], DryRun=True)
    except ClientError as e:
        if 'DryRunOperation' not in str(e):
            raise
    # Dry run succeeded, run start_instances without dryrun
    try:
        response = ec2.start_instances(InstanceIds=[INSTANCE_ID], DryRun=False)
        #print(response)
    except ClientError as e:
        print(e)
    time.sleep(1)

def stopamazon():
    # Do a dryrun first to verify permissions
    try:
        ec2.stop_instances(InstanceIds=[INSTANCE_ID], DryRun=True)
    except ClientError as e:
        if 'DryRunOperation' not in str(e):
            raise
    # Dry run succeeded, call stop_instances without dryrun
    try:
        response = ec2.stop_instances(InstanceIds=[INSTANCE_ID], DryRun=False)
        #print(response)
    except ClientError as e:
        print(e)

class Paars(wx.Panel):
    def __init__(self, parent, frame):
        wx.Panel.__init__(self, parent)
        self.parent = parent
        # Create the menubar and menu
        self.frame = frame 
        #empty data
        self.fn, self.outdir = None, None
        self.status  = 'ok'
        # Now create the Panel to put the other controls on.
        panel = wx.Panel(self)
        lin = wx.StaticLine(panel,-1,style=wx.LI_HORIZONTAL)
        txt = wx.StaticText(panel, -1, " ")
        txta = wx.StaticText(panel, -1, " ")
        text = wx.StaticText(panel, -1, "Merit Insurance - Monte Carlo portfolio simulation")
        text.SetFont(wx.Font(14, wx.SWISS, wx.NORMAL, wx.BOLD))
        text.SetSize(text.GetBestSize())
        self.state,self.color = getamastate()
        self.statebtn = wx.ToggleButton(panel,-1,self.state)
        self.statebtn.SetBackgroundColour(self.color)       
        self.Bind(wx.EVT_TOGGLEBUTTON,self.switchamazon,self.statebtn)
        self.statetext= wx.StaticText(panel, -1, "Click to switch Amazon on or off.  Currently...")
        #Now the filename and output directory
        self.rftext = wx.StaticText(panel, -1, "RollForward Spreadsheet "+ROLLFORWARD)
        self.rftext.SetFont(wx.Font(8, wx.SWISS, wx.NORMAL, wx.BOLD))
        self.sstext = wx.StaticText(panel, -1, "Input Spreadsheet")
        self.sstext.SetFont(wx.Font(8, wx.SWISS, wx.NORMAL, wx.BOLD))
        self.dirtext = wx.StaticText(panel, -1, "Output Directory")
        self.dirtext.SetFont(wx.Font(8, wx.SWISS, wx.NORMAL, wx.BOLD))
        # The buttons
        infil = wx.Button(panel, -1, "1. Choose the spreadsheet", (50,50))
        self.Bind(wx.EVT_BUTTON, self.choosefile, infil)
        outdir = wx.Button(panel, -1, "2. Choose the output directory", (50,50))
        self.Bind(wx.EVT_BUTTON, self.choosedirectory, outdir)
        self.processbtn = wx.Button(panel, -1, "3. Process...")
        self.Bind(wx.EVT_BUTTON, self.process, self.processbtn)
        rslt = wx.Button(panel, -1, "4. Show Summary Results...")
        self.Bind(wx.EVT_BUTTON, self.xldisplay, rslt)
        # Use a sizer to layout the controls, stacked vertically 10 pix border
        sizer = wx.BoxSizer(wx.VERTICAL)
        #statesizer = wx.BoxSizer(wx.HORIZONTAL)
        #statesizer.Add(self.statetext,0,wx.ALL,10)
        #statesizer.Add(self.statebtn,0,wx.ALL,10)
        sizer.Add(text, 0, wx.ALL, 10)
        #sizer.Add(text2, 0, wx.ALL, 10)
        #sizer.Add(iptext, 0, wx.ALL, 10)
        sizer.Add(self.rftext, 0, wx.ALL, 10)
        sizer.Add(self.sstext, 0, wx.ALL, 10)
        sizer.Add(self.dirtext, 0, wx.ALL, 10)
        sizer.Add(txt)
        sizer.Add(self.statetext, 0, wx.LEFT, 10)
        sizer.Add(self.statebtn, 0, wx.ALL, 10)
        sizer.Add(txta)
        sizer.Add(infil, 0, wx.ALL, 10)
        sizer.Add(outdir, 0, wx.ALL, 10)
        sizer.Add(self.processbtn, 0, wx.ALL, 10)
        sizer.Add(rslt, 0, wx.ALL, 10)
        panel.SetSizer(sizer)
        panel.Layout()

        # And also use a sizer to manage the size of the panel such
        # that it fills the frame
        self.frame.SetStatusText('Remote process waiting')
        sizer = wx.BoxSizer()
        sizer.Add(panel, 1, wx.EXPAND)
        self.SetSizer(sizer)
        self.Fit()

    def setstate(self):
        self.state,self.color = getamastate()
        #print(self.state)
        self.statebtn.SetBackgroundColour(self.color)
        self.statebtn.SetLabel(self.state)

    def switchamazon(self,evt):
        if self.state == 'running':
            stopamazon()
            self.state = 'stopping'
        if self.state == 'stopped':
            startamazon()
        time.sleep(1)
        if self.state != 'stopping':
            self.state,self.color = getamastate()
        while self.state not in ['running','stopped','stopping']:
            self.setstate()
            time.sleep(5)
        self.setstate()

    def choosefile(self, evt):
        dlg = wx.FileDialog(
            self, message="Choose a file",
            defaultDir=os.getcwd(),
            defaultFile="",
            wildcard=wildcard,
            style=wx.FD_OPEN | 
                  wx.FD_CHANGE_DIR | wx.FD_FILE_MUST_EXIST |
                  wx.FD_PREVIEW
            )
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.fn = path
            self.justthename = self.fn.split('\\')[-1]
            self.sstext.SetLabel("Input Spreadsheet: " + self.fn)
        dlg.Destroy()
        self.frame.SetStatusText('Input spreadsheet chosen.  Step 1 complete.')

    def OnAbout(self, event):
        dlg = AboutBox()
        dlg.ShowModal()
        dlg.Destroy() 

    def sendfile(self):
        self.msg = '' # this is the message to go in the message line if anything fails
        try: # step 1 check the connection
            client.connect(hostname=IPADDRESS, username="ubuntu", pkey=key)
            client.exec_command('rm /home/test/working/results/*')
            ftp_client=client.open_sftp()
        except:
            self.msg += 'Failed to find server.\n'
        try: # step two send the json files over
            ftp_client.put(self.outdir + '\\' + self.census, RESULTSDIR + 'census.json')
            ftp_client.put(self.outdir + '\\' + self.scenario, RESULTSDIR + 'scenario.json')
        except:
            if not self.msg: self.msg += 'No json files.\n'
        try: # step 3 put a results dummy output file to the server for each scenario
            wb = pd.read_excel(self.fn,index_col=0)
            outputfns = [i + '.xlsx' for i in wb.columns] # col B and on contain the scenario names
            for fn in outputfns:
                ftp_client.put(ROLLFORWARD, RESULTSDIR + fn)
        except:
            if not self.msg: self.msg += 'No dummy output file.\n'
        try: # send the csv files
            for f in self.dfsc.loc['CsvFile'].values:
                fsplit = f.split('.')
                fsplit[0] = fsplit[0].upper() # ubuntu looks for capitalized file names. Could take this out
                fname = '.'.join(fsplit)
                ftp_client.put(self.outdir + '\\' + f, RESULTSDIR + fname)
        except:
            if not self.msg: 
                self.msg = 'No CSVFile in results directory. '+str(f)
        for f in self.dfsc.loc['CensusFile'].values:
            if not pd.isna(f): #then it is a file name not an empty cell
                fsplit = f.split('.')
                fsplit[1] = 'json'
                fname = '.'.join(fsplit)
                self.census_json(self.outdir + '\\' + f,fname)
                try:
                    ftp_client.put(self.outdir + '\\' + fname, RESULTSDIR + fname)
                except:
                    self.msg += 'Error processing ' + str(f) +'\n'
        # done processing the files.  Close the connection.
        ftp_client.close()
        client.close()
        destfn = self.outdir + '\\' + self.justthename 
        if destfn != self.fn:
            try:
                shutil.copy(self.fn,destfn) # only need to copy it if it is not there
            except:
                self.msg += 'Error moving ' + self.fn + ' to results directory.  Is it already there?'
        if self.msg == '':
            self.frame.SetStatusText('Files sent')
        else:
            self.frame.SetStatusText(self.msg)
        #self.msg += 'Done for now - please stop'

    def getfile(self):
        try:
            client.connect(hostname=IPADDRESS, username="ubuntu", pkey=key)
            ftp_client=client.open_sftp()
            fs = ftp_client.listdir(RESULTSDIR)
            for f in fs:
                ftp_client.get(RESULTSDIR + f, self.outdir + '\\' + f)
            ftp_client.close()
            client.close()
            self.frame.SetStatusText('Files retrieved')
        except:
            self.frame.msg('Failed at get file')

    def census_json(self,inf,outf):
        try:
            dfcs = pd.read_excel(inf,'Census',skiprows=1,index_col=0) # get the Census tab if it exists 
        except:
            dfcs = pd.read_excel(inf,skiprows=1,index_col=0) # if not get the first tab
        dfcs2 = dfcs.drop('Total',axis=0) # assumes the total row is named 'Total'
        dfcs2.T.to_json(self.outdir + '\\' + outf)

    def do_spreadsheet(self):
        if self.fn and self.outdir:
            self.census_json(self.fn,self.census)
            self.dfsc = pd.read_excel(self.fn,'Scenarios',index_col=0)
            def_f = open(PCDIRECTORY + 'defaults.json')
            defaults = json.load(def_f)
            for k,v in defaults.items():
                if k not in self.dfsc.index: # if default value is already there, do not override
                    self.dfsc.loc[k] = v
            self.dfsc.to_json(self.outdir + '\\' + self.scenario)
            self.frame.SetStatusText('Spreadsheet processed')
        else:
            self.frame.msg('You have not selected a spreadsheet and output directory yet')

    def getresults(self):
        scenarios = json.load(open(self.outdir + '\\' + 'scenario.json'))
        names = scenarios.keys()
        descs = [v['Description'] for v in scenarios.values()]
        resultsdata = OrderedDict()
        for n,d in zip(names,descs):
            fn = self.outdir + '/table_' + n + '.P'
            results = (pickle.load(open(fn,'rb')))
            results['Description'] = d
            resultsdata[n] = results
        resultsdf = pd.DataFrame(resultsdata)
        ix = list(resultsdf.index)
        # remove Description from wherever it is
        descrix = ix.index('Description')
        ix.pop(descrix)
        ix.insert(0,'Description') # is there a better way???
        resultsdf_ = resultsdf.reindex(ix) # now description is at top
        dft = resultsdf_.T
        dft.to_excel(self.outdir + '/ResultsTable.xlsx')
        return dft

    def xldisplay(self, evt):
        rslt = self.getresults()
        for k in rslt.columns:
            if '%' in k:
                fstr = "{:.2f}%"
            elif '$' in k:
                fstr = "${:,.0f}"
            elif 'year' in k.lower():
                fstr = "{:.0f}"
            elif 'number' in k.lower():
                fstr = "{:,.0f}"
            else:
                fstr = "{0}"
            rslt[k] = [fstr.format(i) for i in rslt[k]]
        a = SummaryResult(self, rslt.to_html()).ShowModal() 

    def choosedirectory(self, evt):
        dlg = wx.DirDialog(self, "Choose a directory:",
                          style=wx.DD_DEFAULT_STYLE
                           )
        if dlg.ShowModal() == wx.ID_OK:
            self.outdir = dlg.GetPath()
            self.census =  'census.json'
            self.scenario =  'scenario.json'
            self.dirtext.SetLabel("Output Directory: " + self.outdir)
        dlg.Destroy()
        self.frame.SetStatusText('Directory chosen,  Step 2 complete.')

    def worker(self):
        try:
            client.connect(hostname=IPADDRESS, username="ubuntu", pkey=key)
            stdin, stdout, stderr = client.exec_command('./mp9starter.sh')
            client.close()
            self.frame.SetStatusText('Process started on Amazon')
        except :
            1/0
            self.frame.msg('Worker failed')

    def getupdate(self):
        client.connect(hostname=IPADDRESS, username="ubuntu", pkey=key)
        stdin, stdout, stderr = client.exec_command('tail ' + RESULTSDIR + 'output.txt')
        output = stdout.readlines()
        client.close()
        if len(output) > 0:
            lastline = output[-1]
        else:
            lastline = 'Running...'
        return lastline

    def checker(self):
        start = True
        while start:
            time.sleep(5)
            status = self.getupdate()
            self.frame.SetStatusText(status)
            if status == 'FINISHED\n':
                start = False
        self.getfile()
        self.frame.SetStatusText('Files retrieved and in output directory')

    def process(self, evt):
        self.processbtn.Disable()
        self.do_spreadsheet()
        self.sendfile()
        if self.msg == '':
            t = threading.Thread(target=self.worker)
            u = threading.Thread(target=self.checker)
            t.start()
            u.start()
        else:
            self.frame.msg(self.msg)

class Blendo(wx.Panel):
    def __init__(self,parent,frame):
        wx.Panel.__init__(self,parent)
        self.parent = parent # refers back to the notebook
        self.K = K
        self.frame = frame
        self.outdir = '...'
        self.setuppath = '...' 
        self.vs = wx.StaticText(self, -1,'',(20,170))
        self.makesttext()
        # radio buttons to do Conning style or AAA style (annual return or cumulative value)
        lblList = ['Conning','AAA']
        self.filetype = lblList[0]
        self.rbox = wx.RadioBox(self, label = 'Who supplied the source files?', pos = (80,10), choices = lblList,
           majorDimension = 1, style = wx.RA_SPECIFY_ROWS) 
        self.rbox.Bind(wx.EVT_RADIOBOX,self.onRadioBox) 
        # end radio buttons 
        self.dirin = wx.Button(self, -1, "1. Choose the portfolio setup spreadsheet", (50,50))
        self.Bind(wx.EVT_BUTTON, self.choosesetupfile, self.dirin)
        self.dirout = wx.Button(self, -1, "2. Choose the output directory", (50,50))
        self.Bind(wx.EVT_BUTTON, self.chooseoutdirectory, self.dirout)
        self.Process = wx.Button(self, -1, "3. Process...")
        self.Bind(wx.EVT_BUTTON, self.process, self.Process)
        # Use a sizer to layout the controls, stacked vertically 10 pix border
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.vs, 0, wx.ALL, 10)
        self.sizer.Add(self.dirin, 0, wx.ALL, 10)
        self.sizer.Add(self.dirout, 0, wx.ALL, 10)
        self.sizer.Add(self.rbox, 0, wx.ALL, 10)
        self.sizer.Add(self.Process, 0,wx.ALL, 10)
        self.SetSizer(self.sizer)
        self.Layout()

    def makesttext(self):
        self.status = 'Portfolio setup files in ' + self.setuppath +'\nOutput csv file in ' + self.outdir
        self.vs.SetLabel(self.status)
   
    def onRadioBox(self,e): 
        self.filetype = self.rbox.GetStringSelection()

    def process(self,evt):
        df = pd.read_excel(self.setuppath,index_col=0,skiprows=1)
        self.filetype = self.rbox.GetStringSelection()
        if self.filetype == 'Conning':
            self.K.INFDIR = K.INDIR + 'ConningDec2019/'
        else:
            self.K.INFDIR = K.INDIR +  'AAACSV/'
        fail=False
        try:
            if self.filetype == 'Conning':
                trial = [['MoneyMarket',df.loc['MoneyMarket'].Weight],
                         ['BondsCorpHighYield',df.loc['BondsCorpHighYield'].Weight],
                         ['BondsCorpIntermediateInv',df.loc['BondsCorpIntermediateInv'].Weight],
                         ['BondsCorpLongInv',df.loc['BondsCorpLongInv'].Weight],
                         ['BondsCorpShortInv',df.loc['BondsCorpShortInv'].Weight],
                         ['BondsGovtIntermediate',df.loc['BondsGovtIntermediate'].Weight],
                         ['BondsGovtLong',df.loc['BondsGovtLong'].Weight],
                         ['BondsGovtShort',df.loc['BondsGovtShort'].Weight],
                         ['EquityInternationalAggressive',df.loc['EquityInternationalAggressive'].Weight],
                         ['EquityInternationalDiversified',df.loc['EquityInternationalDiversified'].Weight],
                         ['EquityUSAggressive',df.loc['EquityUSAggressive'].Weight],
                         ['EquityUSLargeCap',df.loc['EquityUSLargeCap'].Weight],
                         ['EquityUSMidcap',df.loc['EquityUSMidcap'].Weight],
                         ['EquityUSSmallCap',df.loc['EquityUSSmallCap'].Weight]]
            else:
                trial = [['US',df.loc['US'].Weight],
                         ['INT',df.loc['INT'].Weight],
                         ['SMALL',df.loc['SMALL'].Weight],
                         ['AGGR',df.loc['AGGR'].Weight],
                         ['MONEY',df.loc['MONEY'].Weight],
                         ['INTGOV',df.loc['INTGOV'].Weight],
                         ['LTCORP',df.loc['LTCORP'].Weight]]
        except:
            msg = 'The portfolio setup file is not correctly formatted.  Please review.'
            fail=True
        else:
            test = round(sum([i[1] for i in trial]),5)
            error = df.loc['St Dev Error'].Weight
            if test == 1.0:
                outputtab = [[i[0],"{:.2f}%".format(100*i[1])] for i in trial]
                msg= '<html><body><h1>Blending File Data</h1>'
                msg += HTML.table(outputtab,header_row=['Asset','File','Weight %'])
                msg += '<p>Sum of weights = ' + "{:.2f}%".format(100*test) + '</p>'
                msg += '<p>Error = ' + "{:.2f}%".format(100*error) + '</p>'
                outfilenm = self.setuppath.split('\\')[-1].replace(' ','')
                outfilenm = outfilenm.replace('.xlsx','')
                outfilenm = outfilenm.replace('.xls','')
                outfilenm = outfilenm.upper()
                outfilenm += 'ESG' # PJE note sure if this will break something!
                msg += '<p>Output file is ' + outfilenm + '.csv</p>'
                msg += '<p>Close this window to process the file.</p>'
                msg +='</body></html>'
                htmlmsg("Blending file is ok...",msg)
                self.blendo(trial, self.K.INFDIR, self.outdir, outfilenm, error)
                self.frame.msg(outfilenm+'.csv'+' created')
            else:
                msg = 'Portfolio totals do not equal 100% in weights.  Please fix and try again.'
                fail = True
        if fail:
            self.frame.msg(msg)

    def choosesetupfile(self, evt):
        dlg = wx.FileDialog(self, "Choose the input spreadsheet with the blends and error calc:",
                          style=wx.DD_DEFAULT_STYLE,
                          defaultDir=os.getcwd(),
                           )
        if dlg.ShowModal() == wx.ID_OK:
            self.setuppath = dlg.GetPath()
            self.makesttext()
        dlg.Destroy()

    def chooseoutdirectory(self, evt):
        dlg = wx.DirDialog(self, "Choose a directory for the output files:",
                          style=wx.DD_DEFAULT_STYLE
                           )
        if dlg.ShowModal() == wx.ID_OK:
            self.outdir = dlg.GetPath()
            self.makesttext()
        dlg.Destroy()

    def blendo(self,files, indir, outdir, fname, error=0.05): # default error is 5%
        # files is a dictionary.  Key = asset name, Value=[filename,fraction] fraction is decimal 0..1
        #   see trial in __main__ below
        # error is the error term as a %, derived from te 
        self.frame.SetStatusText('Blending files - please wait.')
        filenms = [i[0] for i in files]
        pcts = [i[1] for i in files]
        returns = [loader(fn,indir,self.filetype) for fn in filenms]
        balret = [ret * pct for ret,pct in zip(returns,pcts)]
        sumbalret = sum(balret)
        sds = error * abs(sumbalret)
        e1 = np.random.normal(sumbalret,sds)
        e2 = pd.DataFrame(e1)
        errs = round(e2,6)
        errs.columns = sumbalret.columns
        balgro = 1 + sumbalret
        errbalgro = 1 + errs
        balfundval = np.cumproduct(balgro,axis=1)
        errbalfundval = np.cumproduct(errbalgro,axis=1)
        cols_needed = [i for i in balfundval.columns if i % 12 == 0]
        outputbalfund=round(balfundval[cols_needed],6)
        #outputbalfund.insert(0,-1,1)
        #outputbalfund.to_csv(outdir + '//' + fname +'raw.csv', header=False, index=False)
        outputerrbal = round(errbalfundval[cols_needed],6)
        #outputerrbal.insert(0,-1,1)
        outputerrbal.to_csv(outdir + '//' + fname + '.csv',header = False, index = False)
        #round(sumbalret,6).to_csv(outdir + '//' + fname + 'returns.csv', header = False, index = False)
        #round(errs,6).to_csv(outdir + '//' + fname + 'returnserr.csv', header = False, index = False)

class MainFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,title="MeritInsurance Monte Carlo Simulator",size=(1000,700))
        p = wx.Panel(self)
        self.Bind(wx.EVT_CLOSE, self.OnClose) # event close
        nb = wx.Notebook(p)
        self.CreateStatusBar()
        paarstab = Paars(nb,self)
        blendtab = Blendo(nb,self)
        menuBar = wx.MenuBar()
        menu = wx.Menu()
        menu.Append(101, "&Help\tAlt-H","This will show the instructions")
        menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Exit")
        self.Bind(wx.EVT_MENU, self.Help, id=101)
        self.Bind(wx.EVT_MENU, self.Exit, id=wx.ID_EXIT)
        menuBar.Append(menu,"&Help")
        self.SetMenuBar(menuBar)
        nb.AddPage(paarstab,"Run Paars")
        nb.AddPage(blendtab,"Create Scenarios from blended asset factors")
        sizer = wx.BoxSizer()
        sizer.Add(nb,1,wx.EXPAND)
        p.SetSizer(sizer)
        p.Layout()
        p.Fit()

    def OnClose(self, event):
        self.state,self.color = getamastate()
        #if event.CanVeto() and self.state in ['Running','Pending']:
        if self.state in ['running','pending']:
            if wx.MessageBox("The environment has not been stopped... continue closing?",
                "Please confirm",
                wx.ICON_QUESTION | wx.YES_NO) != wx.YES:
                event.Veto()
                return
        self.Destroy()  # you may also do:  event.Skip()
			# since the default event handler does call Destroy(), too


    def msg(self, msg):
        dlg = wx.MessageDialog(self, msg,
                                     'Message',
                                      wx.OK | wx.ICON_INFORMATION
                                      )
        dlg.ShowModal()
        dlg.Destroy()

    def Help(self, evt):
        text = '''There are four steps...\n
First, set up the spreadsheet with the scenarios and census.  
\tYou can use the spreadsheet Standard.xlsx as a template
\tThen choose the file you have set up.\n
Second, choose an output directory.\n
Third, the Process button will check the spreadsheet,
\tsend it to Amazon for processing, and
\tsend any output back to this computer.\n
Fourth you can review the graphs and spreadsheet
and run the dashboard with the files in the Output directory.
        '''
        self.msg(text)
        #dlg = wx.MessageDialog(self, text,
        #                             'Message',
        #                              wx.OK | wx.ICON_INFORMATION
        #                              )
        #dlg.ShowModal()
        #dlg.Destroy()

    def Exit(self, evt):
        """Event handler for the button click."""
        self.Close()

if __name__ == "__main__":
    app = wx.App()
    MainFrame().Show()
    app.MainLoop()

