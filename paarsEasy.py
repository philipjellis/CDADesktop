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
from collections import OrderedDict, namedtuple
import pickle
from wx.lib import masked
from scipy import stats
from random import random
import HTML
from corr import standarderr

WILDCARD = "Excel sheets (*.xls;*.xlsx)|*.xls;*.xlsxi"     
IPADDRESS = "ec2-3-22-226-74.us-east-2.compute.amazonaws.com"
HOMEDIR = '/home/test/working/' # NB home directory on Linux 
RESULTSDIR = HOMEDIR + 'results/' # NB home directory on Linux - these 2 not the Windows directories
INSTANCE_ID = 'i-03889c587823b7ec5' 

class constants(object):
    pass
K = constants()
K.csvfiles=[]

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

KEY = paramiko.RSAKey.from_private_key_file(KEYFILE)
CLIENT = paramiko.SSHClient()
CLIENT.set_missing_host_key_policy(paramiko.AutoAddPolicy())
EC2 = boto3.client('ec2')
AppBaseClass = wx.App

def flt(s):
    try:
        result = float(s)
    except:
        result = s
    return result 

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

def blendo(parent, data, indir, outdir, fname, error, pftype): # default error is 5%
    # files is a list.  [[asset name, weight]]
    # error is the error term as a fraction, derived from
    parent.frame.SetStatusText('Blending files - please wait.')
    filenms = [i[0] for i in data[:-2]] # last two items are total and r-squared
    pcts = [i[1] for i in data[:-2]]
    returns = [loader(fn,indir,pftype) for fn in filenms]
    balret = [ret * pct for ret,pct in zip(returns,pcts)]
    sumbalret = sum(balret)
    sds = error * abs(sumbalret)
    e1 = np.random.normal(sumbalret,sds)
    e2 = pd.DataFrame(e1)
    #1/0
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
    outputerrbal.to_csv(outdir + '//' + fname, header = False, index = False)
    #round(sumbalret,6).to_csv(outdir + '//' + fname + 'returns.csv', header = False, index = False)
    #round(errs,6).to_csv(outdir + '//' + fname + 'returnserr.csv', header = False, index = False)

class SummaryResult(wx.Dialog): 
    def __init__(self, parent, htstring): 
        super(SummaryResult, self).__init__(parent, title = 'Summary Results', size = (1100,300)) 
        panel = wx.Panel(self) 
		# get and format the results
        self.rslt = self.getresults()
        self.btn = wx.Button(panel, wx.ID_OK, label = "ok") #, size = (50,50)) #, pos = (75,50))
        txt_style = wx.VSCROLL|wx.HSCROLL|wx.TE_READONLY|wx.BORDER_SIMPLE
        htwin = wxhtml.HtmlWindow(panel, -1, size=(1200, 1200), style=txt_style)
        htwin.SetPage(self.rslt.to_html())
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(htwin, 0, wx.ALL, 10)
        sizer.Add(self.btn, 0, wx.ALL, 10)
        panel.SetSizer(sizer)
        panel.Layout()
		
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

    def formatresults(self):
        for k in self.rslt.columns:
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
            self.rslt[k] = [fstr.format(i) for i in self.rslt[k]]

def getamastate():
    alpha = 255
    response = EC2.describe_instances(
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
        EC2.start_instances(InstanceIds=[INSTANCE_ID], DryRun=True)
    except ClientError as e:
        if 'DryRunOperation' not in str(e):
            raise
    # Dry run succeeded, run start_instances without dryrun
    try:
        response = EC2.start_instances(InstanceIds=[INSTANCE_ID], DryRun=False)
        #print(response)
    except ClientError as e:
        print(e)
    time.sleep(1)

def stopamazon():
    # Do a dryrun first to verify permissions
    try:
        EC2.stop_instances(InstanceIds=[INSTANCE_ID], DryRun=True)
    except ClientError as e:
        if 'DryRunOperation' not in str(e):
            raise
    # Dry run succeeded, call stop_instances without dryrun
    try:
        response = EC2.stop_instances(InstanceIds=[INSTANCE_ID], DryRun=False)
        #print(response)
    except ClientError as e:
        print(e)

class Monty(object): # Start the Monte Carlo process

    def __init__(self,parent,scenarios,K):
        self.parent = parent
        self.scenarios = scenarios # scenarios object
        self.K = K
        self.sendfile()
        if self.msg == '':
            t = threading.Thread(target=self.worker)
            u = threading.Thread(target=self.checker)
            t.start()
            u.start()
        else:
            self.parent.frame.msg(self.msg)
	
    def sendfile(self):
        self.msg = '' # this is the message to go in the message line if anything fails
        scenario_names = self.scenarios.keys()
        if self.K.outdir[-1] not in ['\\','/']: self.K.outdir += '/'
        try: # step 1 check the connection
            CLIENT.connect(hostname=IPADDRESS, username="ubuntu", pkey=KEY)
            CLIENT.exec_command('rm /home/test/working/results/*')
            ftp_client=CLIENT.open_sftp()
        except:
            self.msg += 'Failed to find server.\n'
        filestogo = glob.glob(self.K.outdir + '*')
        for f in filestogo:
            filenm = f.split('\\')[-1]
            try:
                ftp_client.put(self.K.outdir + filenm, RESULTSDIR + filenm)
            except:
                self.msg += 'Failed to send' + filenm + '\n'
        # done processing the files.  Close the connection.
        ftp_client.close()
        CLIENT.close()
        if self.msg == '':
            self.parent.frame.SetStatusText('Files sent')
        else:
            self.parent.frame.SetStatusText(self.msg)

    def getfile(self):
        try:
            CLIENT.connect(hostname=IPADDRESS, username="ubuntu", pkey=KEY)
            ftp_client=CLIENT.open_sftp()
            fs = ftp_client.listdir(RESULTSDIR)
            for f in fs:
                ftp_client.get(RESULTSDIR + f, K.outdir + '\\' + f)
            ftp_client.close()
            CLIENT.close()
            self.parent.frame.SetStatusText('Files retrieved')
        except:
            self.parent.frame.msg('Failed at get file')

    def worker(self):
        try:
            CLIENT.connect(hostname=IPADDRESS, username="ubuntu", pkey=KEY)
            stdin, stdout, stderr = CLIENT.exec_command('./mp9starter.sh')
            CLIENT.close()
            self.parent.frame.SetStatusText('Process started on Amazon')
        except :
            self.parent.frame.msg('Worker failed')

    def getupdate(self):
        CLIENT.connect(hostname=IPADDRESS, username="ubuntu", pkey=KEY)
        stdin, stdout, stderr = CLIENT.exec_command('tail ' + RESULTSDIR + 'output.txt')
        output = stdout.readlines()
        CLIENT.close()
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
            self.parent.frame.SetStatusText(status)
            if status == 'FINISHED\n':
                start = False
        self.getfile()
        self.parent.frame.SetStatusText('Files retrieved and in output directory')


class Paars(wx.Panel):
    def __init__(self, parent, frame):
        wx.Panel.__init__(self, parent)
        self.parent = parent
        # Create the menubar and menu
        self.frame = frame 
        #empty data
        self.fn, self.outdir = None, None
        self.K = K
        self.status  = 'ok'
        self.census =  'census.json'
        self.scenario =  'scenario.json'
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
        sizer.Add(text, 0, wx.ALL, 10)
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
            wildcard=WILDCARD,
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

    def do_spreadsheet(self):
        if self.fn and self.outdir:
            self.dfsc = pd.read_excel(self.fn,'Scenarios',index_col=0)
            self.dfsc.columns = [i.replace(' ','') for i in self.dfsc.columns]
            for col in self.dfsc.columns: # copy the rollforward spreadsheet for each run
                rollforward = col + '.xlsx'
                shutil.copy(ROLLFORWARD, self.K.outdir + fn)
            def_f = open(PCDIRECTORY + 'defaults.json')
            defaults = json.load(def_f)
            for k,v in defaults.items():
                if k not in self.dfsc.index: # if default value is already there, do not override
                    self.dfsc.loc[k] = v
            self.scenario_json = self.dfsc.to_json()
            self.frame.SetStatusText('Spreadsheet processed')
        else:
            self.frame.msg('You have not selected a spreadsheet and output directory yet')

    def xldisplay(self, evt):
        a = SummaryResult(self).ShowModal()

    def choosedirectory(self, evt):
        dlg = wx.DirDialog(self, "Choose a directory:",
                          style=wx.DD_DEFAULT_STYLE
                           )
        if dlg.ShowModal() == wx.ID_OK:
            self.K.outdir = dlg.GetPath()
            self.dirtext.SetLabel("Output Directory: " + self.outdir)
        dlg.Destroy()
        self.frame.SetStatusText('Directory chosen,  Step 2 complete.')

    def process(self, evt):
        self.processbtn.Disable()
        self.do_spreadsheet()
        self.moverollforward()
        montecarlo = Monty(self,self.scenario_json,self.K)


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
            self.K.blendir = self.K.INDIR + 'ConningDec2019/'
        else:
            self.K.blendir = self.K.INDIR +  'AAACSV/'
        fail=False
        try:
            if self.filetype == 'Conning':
                trial = [
                         ['BondsGovtIntermediate',df.loc['BondsGovtIntermediate'].Weight],
                         ['BondsCorpLongInv',df.loc['BondsCorpLongInv'].Weight],
                         ['MoneyMarket',df.loc['MoneyMarket'].Weight],
                         ['BondsGovtShort',df.loc['BondsGovtShort'].Weight],
                         ['BondsGovtLong',df.loc['BondsGovtLong'].Weight],
                         ['BondsCorpShortInv',df.loc['BondsCorpShortInv'].Weight],
                         ['BondsCorpHighYield',df.loc['BondsCorpHighYield'].Weight],
                         ['BondsCorpIntermediateInv',df.loc['BondsCorpIntermediateInv'].Weight],
                         ['EquityUSLargeCap',df.loc['EquityUSLargeCap'].Weight],
                         ['EquityUSSmallCap',df.loc['EquityUSSmallCap'].Weight],
                         ['EquityUSAggressive',df.loc['EquityUSAggressive'].Weight],
                         ['EquityUSMidcap',df.loc['EquityUSMidcap'].Weight],
                         ['EquityInternationalAggressive',df.loc['EquityInternationalAggressive'].Weight],
                         ['EquityInternationalDiversified',df.loc['EquityInternationalDiversified'].Weight]]
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
            test = round(sum([i[1] for i in trial]),6)
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
                outfilenm = outfilenm.upper().replace(' ','')
                outfilenm += 'ESG.csv' # PJE note sure if this will break something!
                self.K.csvfiles.append(outfilenm)
                msg += '<p>Output file is ' + outfilenm + '.csv</p>'
                msg += '<p>Close this window to process the file.</p>'
                msg +='</body></html>'
                htmlmsg("Blending file is ok...",msg)
                blendo(self, trial, self.K.blendir, self.outdir, outfilenm, error, self.filetype)
                self.frame.msg(outfilenm + ' created')
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

class Easy(wx.Panel):
    def __init__(self,parent,frame):
        wx.Panel.__init__(self,parent)
        self.parent = parent # refers back to the notebook
        self.aaakeys = set(['Portfolio', 'date_range', 'Cash', 'SandP_500_Index', 'MSCI_EAFE_USD', 
                'BBgBarc_US_Treasury_3_5_Yr_TR_USD', 'Bbgbarc_US_Corporate_7_10_years_TR_USD', 
                'Russell_2000_TR_USD', 'AAAAggr', 'Style_R_Squared', 'Predicted_R_Squared'])
        self.connkeys = set(['Portfolio', 'date_range', 'Cash', 'ICE_BofA_5_10_Year_US_Treasury_Index', 
                'ICE_BofA_10plus_Year_US_Corporate_Index', 'ICE_BofA_1_5_Year_US_Treasury_Index', 
                'ICE_BofA_10plus_Year_US_Treasury_Index', 'ICE_BofA_1_5_Year_US_Corporate_Index', 
                'ICE_BofA_US_High_Yield_Index', 'ICE_BofA_5_10_Year_US_Corporate_Index', 
                'SandP_500_Index', 'R2000', 'NASDAQ_Composite_TR_USD', 'Russell_Mid_Cap_Index', 
                'MSCI_EM_Emerging_Markets_USD', 'MSCI_EAFE', 'Style_R_Squared', 'Predicted_R_Squared'])
        self.K = K
        self.frame = frame
        self.K.outdir = 'w:/My Documents/CDALive/test/'
        self.vs = wx.StaticText(self, -1,'',(20,170))
        self.makesttext()
        self.dirout = wx.Button(self, -1, "1. Choose the output directory", (50,50))
        self.Bind(wx.EVT_BUTTON, self.chooseoutdirectory, self.dirout)
        self.Paste = wx.Button(self, -1, "2. Paste MPI Data")
        self.Bind(wx.EVT_BUTTON, self.paste, self.Paste)
        # Use a sizer to layout the controls, stacked vertically 10 pix border
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.vs, 0, wx.ALL, 10)
        self.sizer.Add(self.dirout, 0, wx.ALL, 10)
        self.sizer.Add(self.Paste, 0,wx.ALL, 10)
        self.SetSizer(self.sizer)
        self.Layout()

    def makesttext(self):
        self.status = 'Output files in ' + self.K.outdir
        self.vs.SetLabel(self.status)

    def paste(self, evt):
        success = False
        data = wx.TextDataObject()

        if wx.TheClipboard.Open():
            success = wx.TheClipboard.GetData(data)
            wx.TheClipboard.Close()
        if not success:
            error_msg = 'No data in clipboard'
        
        state,color = getamastate()
        if state != 'running':
            success = False
            error_msg = 'Amazon server is OFF.'

        if success:
            error_msg, ass_weights, pftype = self.processClipboard(data)
            json_run = {} # this becomes scenario.json
            run_summary = {} # contains the weights and R_squared
            if not error_msg:
                for key,val in ass_weights.items():
                    #print (key,val)
                    error_msg, fname, data = self.prepare_blend(error_msg, key, val, pftype)
                    self.stderr = standarderr(data, pftype)
                    run_summary[key] = data
                    if not error_msg: blendo(self, data, self.K.blendir, self.K.outdir, fname, self.stderr, pftype) # puts ESG file in outdir
                    if not error_msg: json_run[key] = self.json_details(key,fname)

            self.frame.SetStatusText('Files blended ok.')
            if not error_msg: # we have all the individual runs prepared
                with open(self.K.outdir + 'scenario.json','w') as jsfile:
                    json.dump(json_run,jsfile)
                    #self.scenario_json = json.dumps(json_run)
                with open(self.K.outdir + 'summary.json','w') as summaryfile:
                    json.dump(run_summary,summaryfile) # may need this for html display...
                shutil.copy(PCDIRECTORY+'census.json',self.K.outdir+'census.json')

        if error_msg:
            wx.MessageBox(error_msg,"Error")
        else:
            montecarlo = Monty(self,json_run,self.K) # start the Monte Carlo engine

    def processClipboard(self,data):
        failmsg = ''
        try:
            lines = data.GetText().split('\n')
        except:
            failmsg = 'Clipboard does not contain data'
        if not failmsg:
            try:
                heads = lines[1].split('\t')
                heads[1] = 'date_range'
                heads[0] = 'Portfolio'
            except:
                failmsg = 'Invalid clipboard, no MPI heading data'
        if not failmsg:
            delete_dict = {' ':'_','|':'_','-':'_',')':'','(':'','+':'plus','&':'and'}
            deletetable = str.maketrans(delete_dict)
            heads = [i.translate(deletetable) for i in heads]
            pftuple = namedtuple('portfolio',heads)
            clipdata = OrderedDict()
            for line in lines[2:]: # may need a try/except here, see if it ever fails...
                linedata = line.split('\t')
                linedata = [flt(i) for i in linedata]
                try:
                    pfdata = pftuple(*linedata)
                    clipdata[pfdata.Portfolio] = pfdata
                except:
                    pass # there is a harmless blank line at the end, which will cause an error
            if set(heads) == self.aaakeys:
                pftype = 'AAA'
                retkeys = self.aaakeys
            elif set(heads) == self.connkeys:
                pftype = 'Conning'
                retkeys = self.connkeys
            else:
                pftype = 'Unknown'
                retkeys = None
                failmsg = 'Unknown Portfolio keys ' + ' '.join(heads)
            lst = list(clipdata.keys())
            defaults = range(len(lst)) # default choice = everything
            dlg = wx.MultiChoiceDialog( self, # now set up a chooser, so you do not have to do it all
                               "Choose the portfolios",
                               pftype + "Choices...", lst)
            dlg.SetSelections(defaults)
            if (dlg.ShowModal() == wx.ID_OK):
                selections = dlg.GetSelections()
                keysneeded = [lst[x] for x in selections]
            else:
                keysneeded = None
            dlg.Destroy()
            result = dict((k, clipdata[k]) for k in keysneeded)
        else:
            result = None
            pftype = 'Not known'
        return failmsg, result, pftype

    def prepare_blend(self, err_msg, name, data, pftype):
        msg = ''
        self.filetype = pftype
        if self.filetype == 'Conning':
            self.K.blendir = self.K.INDIR + 'ConningDec2019/'
        else:
            self.K.blendir = self.K.INDIR +  'AAACSV/'
        fail=False
        try:
            if self.filetype == 'Conning':
                output = [
                         ('MoneyMarket',data.Cash),
                         ('BondsCorpHighYield',data.ICE_BofA_US_High_Yield_Index),
                         ('BondsCorpIntermediateInv',data.ICE_BofA_5_10_Year_US_Corporate_Index),
                         ('BondsCorpLongInv',data.ICE_BofA_10plus_Year_US_Corporate_Index),
                         ('BondsCorpShortInv',data.ICE_BofA_1_5_Year_US_Corporate_Index),
                         ('BondsGovtIntermediate',data.ICE_BofA_5_10_Year_US_Treasury_Index),
                         ('BondsGovtLong',data.ICE_BofA_10plus_Year_US_Treasury_Index),
                         ('BondsGovtShort',data.ICE_BofA_1_5_Year_US_Treasury_Index),
                         ('EquityInternationalAggressive',data.MSCI_EM_Emerging_Markets_USD),
                         ('EquityInternationalDiversified',data.MSCI_EAFE),
                         ('EquityUSAggressive',data.NASDAQ_Composite_TR),
                         ('EquityUSLargeCap',data.SandP_500_Index),
                         ('EquityUSMidcap',data.Russell_Mid_Cap_Index),
                         ('EquityUSSmallCap',data.R2000)] #)
            else:
                output = [ 
                         ('US',data.SandP_500_Index),
                         ('INT',data.MSCI_EAFE_USD),
                         ('SMALL',data.Russell_2000_TR_USD),
                         ('AGGR',data.AAAAggr),
                         ('MONEY',data.Cash),
                         ('INTGOV',data.BBgBarc_US_Treasury_3_5_Yr_TR_USD),
                         ('LTCORP',data.Bbgbarc_US_Corporate_7_10_years_TR_USD)] #)
        except:
            msg += 'The portfolio setup file is not correctly formatted.  Please review.'
        output.append(('total', sum([i[1] for i in output])))
        output.append(('Predicted_R_Squared', data.Predicted_R_Squared))
        output2 = [(i[0],round(i[1]/100,6)) for i in output]
        outfilenm = name.upper().replace(' ','') 
        outcsv = outfilenm + 'ESG.csv'
        rollforward = outfilenm + '.xlsx'
        shutil.copy(ROLLFORWARD, self.K.outdir + rollforward)
        return msg, outcsv, output2

    def json_details(self,name,fname):
        return {"Description": name,
                "CsvFile": fname,
                "CensusFile": None,
                "Fund":100000000.0,
                "Population": 1000,
                "Runs": 1000,
                "AnnCum": "CUM",
                "Income": .05,
                "Premium": .0055,
                "WMFee": .01,
                "DiscountRate": .03,
                "MortSpreadsheet": "IAM20122581_2582.xlsx",
                "LapseUtilization":"Utilization2020_1_26.xlsx",
                "Stochastic": False,
                "Prudent": True,
                "Debug":False,
                "YearOneIncome": 1.0}

    def chooseoutdirectory(self, evt):
        dlg = wx.DirDialog(self, "Choose a directory for the output files:",
                          style=wx.DD_DEFAULT_STYLE
                           )
        if dlg.ShowModal() == wx.ID_OK:
            self.K.outdir = dlg.GetPath()
            self.makesttext()
        dlg.Destroy()

class MainFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,title="MeritInsurance Monte Carlo Simulator",size=(1000,700))
        p = wx.Panel(self)
        self.Bind(wx.EVT_CLOSE, self.OnClose) # event close
        nb = wx.Notebook(p)
        self.CreateStatusBar()
        paarstab = Paars(nb,self)
        blendtab = Blendo(nb,self)
        easytab = Easy(nb,self)
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
        nb.AddPage(easytab,"QuickRun: one MPI table")
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

