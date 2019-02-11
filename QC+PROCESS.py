
# coding: utf-8

# In[1]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime
from datetime import date
import requests
from bs4 import BeautifulSoup
today=datetime.datetime.today()
today=today.strftime('%m-%d-%y')
import pandas as pd
from pandas import ExcelWriter
import win32com.client as win32


# In[2]:



path =str(r'C:/Users/saransh_arora/AppData/Local/Continuum/anaconda3/chromedriver.exe')
browserpath=str(r'https://sarthak_negi_:Exantas@123@biapps.capitaliq.com/Reports/Pages/Folder.aspx?ItemPath=%2fResearch%2fEntity+Management%2fPricing+and+Index+Data%2fPricing+Related')

browser=webdriver.Chrome(path)
browser.maximize_window()
# browser.add_argument('--headless')
# browser.add_argument('--disable-gpu')                                                        
browser.get(browserpath)


# In[3]:


test = ''
while not test:
    try:
        browser.find_element_by_xpath('//*[@id="ui_a26"]').click()
        test='ok'
    except:
        pass    
browser.find_element_by_id('ctl150_ctl01_ctl05_ctl00').click()
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl00"]/option[6]').click()
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl01"]').click()
browser.get(browserpath)

# Last week Price

test = ''
while not test:
    try:
        browser.find_element_by_xpath('//*[@id="ui_a1"]').click()
        test='ok'
    except:
        pass        
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl00"]/option[6]').click()
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl01"]').click()
browser.get(browserpath)

# Non Public Ticker

test = ''
while not test:
    try:
        browser.find_element_by_xpath('//*[@id="ui_a18"]').click()
        test='ok'
    except:
        pass               
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl00"]/option[6]').click()
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl01"]').click()
browser.get(browserpath)

# Public no FT

test = ''
while not test:
    try:
        browser.find_element_by_xpath('//*[@id="ui_a23"]').click()
        test='ok'
    except:
        pass               
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl00"]/option[6]').click()
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl01"]').click()
browser.get(browserpath)

# Public Acquired OOB

test = ''
while not test:
    try:
        browser.find_element_by_xpath('//*[@id="ui_a22"]').click()
        test='ok'
    except:
        pass        
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl00"]/option[6]').click()
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl01"]').click()
browser.get(browserpath)

# Missing ADR_GDR

test = ''
while not test:
    try:
        browser.find_element_by_xpath('//*[@id="ui_a15"]').click()
        test='ok'
    except:
        pass  
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl00"]/option[6]').click()
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl01"]').click()
browser.get(browserpath)

# NonADRPrimaryInADRList

test = ''
while not test:
    try:
        browser.find_element_by_xpath('//*[@id="ui_a19"]').click()
        test='ok'
    except:
        pass  
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl00"]/option[6]').click()
browser.find_element_by_xpath('//*[@id="ctl150_ctl01_ctl05_ctl01"]').click()
browser.get(browserpath)


# In[4]:


import pandas as pd
from pandas import ExcelWriter

import datetime
from datetime import datetime,timedelta 

weekday = datetime.today().weekday()

if weekday == 0:
    timedelta_val = 3
else:
    timedelta_val = 1

date = datetime.now().strftime("%Y-%m-%d") 
date_minus_one = (datetime.now() - timedelta(timedelta_val)).strftime("%Y-%m-%d") 

print(date)
print(date_minus_one)


# In[5]:


QCpath=r'\\II02FIL001.mhf.mhc\FT\2. Operations\MDCA - Securities Management\MDCA Securities Management Processes\QC Process\To Be Worked'

writer_public_nullticker=QCpath+r'\Public With No Ticker and Exchange\Raw Files\{}.xlsx'.format(date)

writer_lastweekpricing=QCpath+r"\Last Week Pricing wrong type\Raw File\{}.xlsx".format(date)  

writer_nonpublicticker=QCpath+r"\Non Public with ticker or exchange\Raw Files\{}.xlsx".format(date) 

writer_public_noft=QCpath+r"\Public and no FT Trading\Raw Files\{}.xlsx".format(date) 

writer_acquiredoob=QCpath+r"\Acquired OOB companies with Ticker and exchange\{}.xlsx".format(date) 

writer_missingadr=QCpath+r"\Missing ADR & GDR\{}.xlsx".format(date) 

writer_nonadrprimary=QCpath+r"\Non ADR Primary In ADR List\{}.xlsx".format(date) 


# In[6]:


files_processed = ['Files Processed : ' ]
files_not_processed = ['Files not Processed : ' ]


# In[7]:


############################################## PUBLIC WITH NULL TICKER ##################################################################
try:
    df_public_nullticker=pd.read_excel(r'C:\Users\saransh_arora\Downloads\Public With Null Ticker and Exchange.xls')
    df_public_nullticker.to_excel(writer_public_nullticker,header=True,index=False)
    df1_public_nullticker=df_public_nullticker.drop_duplicates(subset='Company ID',keep='first')
    df2_public_nullticker=df1_public_nullticker[(df1_public_nullticker['Company Status']=='Operating')|(df1_public_nullticker['Company Status']=='Operating Subsidiary')|(df1_public_nullticker['Company Status']=='Reorganizing')|(df1_public_nullticker['Company Status']=='Liquidating')]


    public_nullticker_yesterday_path_sameday = QCpath+r'\Public With No Ticker and Exchange\Daily Files\{}.xlsx'.format(date)
    public_nullticker_yesterday_path_yesterday = QCpath+r'\Public With No Ticker and Exchange\Daily Files\{}.xlsx'.format(date_minus_one)

    try:
        df_nullticker_yesterday = pd.read_excel(public_nullticker_yesterday_path_sameday)
    except:
        df_nullticker_yesterday = pd.read_excel(public_nullticker_yesterday_path_yesterday)


    df2_public_nullticker = pd.merge(df2_public_nullticker,df_nullticker_yesterday[['Company ID','Status','Comments']],how='left',on='Company ID')

    df2_public_nullticker['Researcher'] = ''

    df2_public_nullticker['Date'] = ''

    writer_public_nullticker_dailyfile=QCpath+r'\Public With No Ticker and Exchange\Daily Files\{}.xlsx'.format(date)
    df2_public_nullticker.to_excel(writer_public_nullticker_dailyfile,header=True,index=False)
    files_processed.append("Public With Null Ticker,")

except:
    files_not_processed.append("Public With Null Ticker,")
    


# In[8]:


############################################## ACQUIRED OOB WITH TICKER & EXCHANGE ########################################################

try:
    df_acquiredoob=pd.read_excel(r'C:\Users\saransh_arora\Downloads\Public Acquired or OOB with Active Primary.xls')


    yesterday_path_oob_sameday = QCpath+r"\Acquired OOB companies with Ticker and exchange\{}.xlsx".format(date)

    yesterday_path_oob_yesterday = QCpath+r"\Acquired OOB companies with Ticker and exchange\{}.xlsx".format(date_minus_one)

    try:
        df_acquiredoob_yesterday = pd.read_excel(yesterday_path_oob_sameday)
    except:
        df_acquiredoob_yesterday = pd.read_excel(yesterday_path_oob_yesterday)

    df_acquiredoob = pd.merge(df_acquiredoob,df_acquiredoob_yesterday[['Company ID','Status']],how='left',on='Company ID')

    df_acquiredoob['Researcher'] = ''
    df_acquiredoob['Date'] = ''

    df_acquiredoob=df_acquiredoob.drop_duplicates(subset='Company ID',keep='first')

    df_acquiredoob.to_excel(writer_acquiredoob,header=True,index=False)
    
    files_processed.append("Acquired OOB with Ticker & Exchange,")
    
except:
    files_not_processed.append("Acquired OOB with Ticker & Exchange,")
    


# In[9]:


############################################## MISSING ADR & GDR ##########################################################################

try:
    
    df_missing_adr_gdr=pd.read_excel(r'C:\Users\saransh_arora\Downloads\Missing ADR_GDR.xls')

    yesterday_path_adr_gdr_sameday = QCpath+r"\Missing ADR & GDR\{}.xlsx".format(date)

    yesterday_path_adr_gdr_yesterday = QCpath+r"\Missing ADR & GDR\{}.xlsx".format(date_minus_one)

    try:
        df_missing_adr_gdr_yesterday = pd.read_excel(yesterday_path_adr_gdr_sameday)
    except:
        df_missing_adr_gdr_yesterday = pd.read_excel(yesterday_path_adr_gdr_yesterday)

    df_missing_adr_gdr = pd.merge(df_missing_adr_gdr,df_missing_adr_gdr_yesterday[['Company Id','Status']],how='left',on='Company Id')

    df_missing_adr_gdr.to_excel(writer_missingadr,header=True,index=False)
    
    files_processed.append("Missing ADR & GDR,")
    
except:
    
    files_not_processed.append("Missing ADR & GDR,")


# In[10]:


############################################## NON ADR PRIMARY IN ADR LIST ################################################################
try:
    
    df_nonadrprimary=pd.read_excel(r'C:\Users\saransh_arora\Downloads\NonADRPrimaryInADRList.xls')

    yesterday_path_non_adr__in_adrlist_sameday = QCpath+r"\Non ADR Primary In ADR List\{}.xlsx".format(date)
    yesterday_path_non_adr__in_adrlist_yesterday = QCpath+r"\Non ADR Primary In ADR List\{}.xlsx".format(date_minus_one)

    try:
        df_nonadrprimary_yesterday = pd.read_excel(yesterday_path_non_adr__in_adrlist_sameday)
    except:
        df_nonadrprimary_yesterday = pd.read_excel(yesterday_path_non_adr__in_adrlist_yesterday)

    df_nonadrprimary = pd.merge(df_nonadrprimary,df_nonadrprimary_yesterday[['Company ID','Status']],how='left',on='Company ID')

    df_nonadrprimary['Researcher'] = ''

    df_nonadrprimary.to_excel(writer_nonadrprimary,header=True,index=False)
    
    files_processed.append("Non-ADR Primary in ADR List,")

except:
    
    files_not_processed.append("Non-ADR Primary in ADR List,")



# In[11]:


############################################## PUBLIC & NO FT TRADING #####################################################################

try:
    
    df_public_noft=pd.read_excel(r'C:\Users\saransh_arora\Downloads\Public and no FT Trading Items.xls')
    df_public_noft.to_excel(writer_public_noft,header=True,index=False)
    df1_public_noft=df_public_noft[(df_public_noft['Company Status']=='Operating') | (df_public_noft['Company Status']=='Operating Subsidiary')]

    yesterday_path_public_noft_sameday = QCpath+r'\Public and no FT Trading\Daily Files\{}.xlsx'.format(date)
    yesterday_path_public_noft_yesterday = QCpath+r'\Public and no FT Trading\Daily Files\{}.xlsx'.format(date_minus_one)

    try:
        df_public_noft_yesterday = pd.read_excel(yesterday_path_public_noft_sameday)
    except:
        df_public_noft_yesterday = pd.read_excel(yesterday_path_public_noft_yesterday)

    df1_public_noft = pd.merge(df1_public_noft,df_public_noft_yesterday[['Company ID','Status']],how='left',on='Company ID')

    df1_public_noft['Researcher'] = ''
    df1_public_noft['Date'] = ''

    writer_public_noft_dailyfile = QCpath+r'\Public and no FT Trading\Daily Files\{}.xlsx'.format(date)

    df1_public_noft.to_excel(writer_public_noft_dailyfile,header=True,index=False)
    
    files_processed.append("Public & No FT Trading, ")
    
except:
    
    files_not_processed.append("Public & No FT Trading, ")


# In[12]:


############################################## NON PUBLIC WITH TICKER or EXCHANGE #########################################################

try:
    
    df_nonpublic_with_tickerorexchange = pd.read_excel(r'C:\Users\saransh_arora\Downloads\Non Public with ticker or exchange.xls')
    df_nonpublic_with_tickerorexchange.to_excel(writer_nonpublicticker,header=True,index=False)

    yesterday_path_nonpublic_tickerorexchange_sameday = QCpath + r'\Non Public with ticker or exchange\Daily Files\{}.xlsx'.format(date)
    yesterday_path_nonpublic_tickerorexchange_yesterday = QCpath + r'\Non Public with ticker or exchange\Daily Files\{}.xlsx'.format(date_minus_one)

    try:
        df_nonpublic_exchangeorticker_yesterday = pd.read_excel(yesterday_path_nonpublic_tickerorexchange_sameday)
    except:
        df_nonpublic_exchangeorticker_yesterday = pd.read_excel(yesterday_path_nonpublic_tickerorexchange_yesterday)

    try:
        df_nonpublic_with_tickerorexchange = df_nonpublic_with_tickerorexchange[df_nonpublic_with_tickerorexchange['exchange Symbol'].str.contains("FUND|MutualFund|BOIN|OTCUS|UNQ|Xtrakter") == False]
        df_nonpublic_with_tickerorexchange = df_nonpublic_with_tickerorexchange[df_nonpublic_with_tickerorexchange['security Name'].str.contains("%")==False]
    except:
        pass


    df_nonpublic_tickerorexchange_daily = df_nonpublic_with_tickerorexchange

    df_nonpublic_tickerorexchange_daily['Last Sale price'] = ''
    df_nonpublic_tickerorexchange_daily['VOL'] = ''
    df_nonpublic_tickerorexchange_daily['Last Pricing date'] = ''



    df_nonpublic_tickerorexchange_daily =pd.merge(df_nonpublic_tickerorexchange_daily,df_nonpublic_exchangeorticker_yesterday[['Company ID','IQ ','Status','Comments']],how='left',on='Company ID')


    df_nonpublic_tickerorexchange_daily['Researcher'] = ''
    df_nonpublic_tickerorexchange_daily['Date'] = ''
    df_nonpublic_tickerorexchange_daily['TradingItemId'] = ''
    df_nonpublic_tickerorexchange_daily['MarketCap'] = ''
    df_nonpublic_tickerorexchange_daily['>95m?'] = ''


    nonpublic_exchangeorticker_dailyfile_ =  QCpath+r'\Non Public with ticker or exchange\Daily Files\{}.xlsx'.format(date)
    df_nonpublic_tickerorexchange_daily.to_excel(nonpublic_exchangeorticker_dailyfile_,header=True,index=False)
    
    files_processed.append("Non Public with Ticker or Exchange,")
    
except:
    
    files_not_processed.append("Non Public with Ticker or Exchange,")


# In[22]:


########################################## LAST WEEK PRICING WRONG TYPE #################################################################

try:
    
    df_lastweekpricing_raw=pd.read_excel(r'C:\Users\saransh_arora\Downloads\Companies has Last Week Price with Wrong Company Type.xls')
    df_lastweekpricing_raw.to_excel(writer_lastweekpricing,header=True,index=False)

    yesterday_path_lastweekpricing_sameday = QCpath + r'\Last Week Pricing wrong type\Daily File\{}.xlsx'.format(date)
    yesterday_path_lastweekpricing_yesterday = QCpath + r'\Last Week Pricing wrong type\Daily File\{}.xlsx'.format(date_minus_one)

    try:
        df_yesterday_lastweekpricing = pd.read_excel(yesterday_path_lastweekpricing_sameday)
    except:
        df_yesterday_lastweekpricing = pd.read_excel(yesterday_path_lastweekpricing_yesterday)



    try:
        df_lastweekpricing_raw = df_lastweekpricing_raw[df_lastweekpricing_raw['exchange Symbol'].str.contains("FUND|MutualFund|BOIN|OTCUS|UNQ|Xtrakter") == False]
        df_lastweekpricing_raw = df_lastweekpricing_raw[df_lastweekpricing_raw['security Name'].str.contains("%")==False]
    except:
        pass



    df_lastweekpricing = df_lastweekpricing_raw

    df_lastweekpricing = pd.merge(df_lastweekpricing,df_yesterday_lastweekpricing[['company Id','Status']],how='left',on='company Id')


    df_lastweekpricing['Status']=''
    df_lastweekpricing['Researcher']=''
    df_lastweekpricing['LSP'] = ''
    df_lastweekpricing['Security Type'] = ''
    df_lastweekpricing['Volume'] = ''
    df_lastweekpricing['Last price date'] = ''



    writer_lastweekpricing_daily = QCpath + r'\Last Week Pricing wrong type\Daily File\{}.xlsx'.format(date)
    df_lastweekpricing = df_lastweekpricing.drop_duplicates(keep='first')
    df_lastweekpricing.to_excel(writer_lastweekpricing_daily, sheet_name='Sheet1',index=False)
    
    
    files_processed.append("Last Week Pricing Wrong Type")
    
except:
    
    files_not_processed.append("Last Week Pricing Wrong Type")


# In[14]:


import os
try:
    os.remove(r'C:\Users\saransh_arora\Downloads\Public With Null Ticker and Exchange.xls')
    os.remove(r'C:\Users\saransh_arora\Downloads\Public Acquired or OOB with Active Primary.xls')
    os.remove(r'C:\Users\saransh_arora\Downloads\Missing ADR_GDR.xls')
    os.remove(r'C:\Users\saransh_arora\Downloads\NonADRPrimaryInADRList.xls')
    os.remove(r'C:\Users\saransh_arora\Downloads\Public and no FT Trading Items.xls')
    os.remove(r'C:\Users\saransh_arora\Downloads\Non Public with ticker or exchange.xls')
    os.remove(r'C:\Users\saransh_arora\Downloads\Companies has Last Week Price with Wrong Company Type.xls')
except:
    print("Files Don't Exist")
    


# In[15]:


file_location = 'File_location : \\II02FIL001.mhf.mhc\FT\2. Operations\MDCA - Securities Management\MDCA Securities Management Processes\QC Process\To Be Worked'


# In[16]:


location_file = ''.join(file_location)
files_processed_str = ''.join(files_processed)
files_not_processed_str = ''.join(files_not_processed)


# In[17]:


files_not_processed_str


# In[18]:


files_processed_str


# In[19]:


msg_body = location_file + "                " + files_processed_str  + "           " + files_not_processed_str 


# In[20]:


msg_body


# In[21]:


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'saransh_arora@spglobal.com'
mail.Subject = 'QC Processed'
mail.Body = msg_body
mail.send
# mail = outlook.CreateItem(0)
# mail.To = 'aishwarya_bajpayee@spglobal.com'
# mail.Subject = 'QC Processed'
# mail.Body = 'Files Process : ',files_processed ,"     ", 'Files Not Processed : ',files_not_processed,"  ",'File Location : ',file_location
# mail.send

