import traceback
import sys
import eel
import getpass
import sys
import time
import datetime
import platform
import subprocess
import os
  
##get the things to thingy
app_version = '0.1.2p'
web = 'web'
SERVER = 'jogwks0003'


class TodayTime:

   def __init__(self):
     self.is_windows = (sys.platform == 'win32')
     self.wks = platform.uname()[1].upper()
     self.name = self.get_name()
     self.time = self.get_time()
     
     
   def get_name(self):
     print('getting name')
     get_name = getpass.getuser()
     return get_name.replace('.', ' ').title()
     
     
   def get_time(self):
     print('getting login time')
     if self.is_windows:
       from winevt import EventLog
       query = EventLog.Query("System","Event/System[Level=4]",direction='backward')
       for event in query:
           if event.System.Provider['Name'] == 'Microsoft-Windows-Winlogon':
               login_time = event.System.TimeCreated['SystemTime']
               break          
     login_time = login_time.split('.')[0] if self.is_windows else subprocess.check_output("who").split()[3].decode("utf-8") 
     login_time = time.strptime(login_time,'%Y-%m-%dT%H:%M:%S') if self.is_windows else time.strptime(login_time,'%H:%M')
     login_time = datetime.datetime.fromtimestamp(time.mktime(login_time)) 
     if self.is_windows: 
         login_time += datetime.timedelta(hours = 7, minutes = -2) 
     self.is_late = datetime.datetime.now()
     self.is_late = self.is_late.replace(hour=8, minute=0, second=0, microsecond=0)
     self.is_late = 'LATE' if login_time > self.is_late else 'ON TIME'
     login_time_clear = datetime.datetime.strftime(login_time, "%H:%M:%S") if self.is_windows else '{}:{}'.format(login_time.hour, login_time_clear.minute)
     return login_time_clear
     

class ArrivalCahGL(TodayTime):

   def __init__(self):
     print('initializing...')
     super().__init__()
     self.date = datetime.datetime.strftime(datetime.datetime.now(), "%b-%d-%Y")
     self.tglsheet = self.get_date()
     
     if self.is_windows:
       self.retry = 3
       self.filename = 'arrival_' + datetime.datetime.strftime(datetime.datetime.now(), "%b-%d-%Y") + '.csv'
       self.full_filename = '//'+SERVER+'/windowsteam/result/'+self.filename
       os.chdir('c:/users/'+getpass.getuser()+'/')
       eel.init(os.getcwd()+'/'+web)
       text = (self.wks, self.name, self.time)
       self.text_to_write = ','.join(text)
       self.test_csv(self.name)
     print('initializing finished...')
     
   def get_date(self):
     today = datetime.datetime.now()
     hari = str(today.day)
     bulan = str(today.month)
     tahun = str(today.year)
     return '/'.join([bulan, hari, tahun])

     
   def test_csv(self, name):
     try:
       file = open(self.full_filename, 'a')
       file.close()
       file = open(self.full_filename)
       testname = file.read()
       file.close()
       if name in testname: 
           import csv
           csvRows = []
           csvFileObj = open(self.full_filename)
           readerObj = csv.reader(csvFileObj)

           #read then put data to list
           for i in readerObj:
               csvRows.append(i)

           csvFileObj.close()
           total_baris_csv = len(csvRows)

           for n in range(total_baris_csv):
               if csvRows[n][1] == name:
                   login_time_clear = csvRows[n][2]
                   return login_time_clear
       else:          
         self.write_to_file_win(self.text_to_write)
     except Exception as error :
       print(error)
       if self.retry > 0 : self.logging('win_read_error', retry = 1) 
       else: self.logging('error')
       
       
   def logging(self, log_type, retry = 0):
      if log_type == 'win_read_error':
        time.sleep(600)
        self.retry -= retry
        self.test_csv(self.name)
        return
      try:
        self.LOG = open('//'+SERVER+'/windowsteam/result/LOG.log', 'a')  
      except:
        self.LOG = open('LOG.log', 'a')        
      self.logtime = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d %H:%M:%S')
      if log_type == 'error':
        self.LOG.write('{} : {} \n => Failed'.format(self.logtime, self.name))
        self.LOG.write('-'*len(name)+'--\n')
        self.LOG.write(traceback.format_exc())
        self.LOG.write('\n')
        self.LOG.close()        
      elif log_type == 'sent_success':
        self.LOG.write('{} : {} => Success \n '.format(self.logtime, self.name))
      self.LOG.close()
      
   def write_to_file_win(self, text_to_write):
     file  = open(self.full_filename, 'a')
     file.write(text_to_write)
     file.write('\n')
     file.close()
     #return self.time

     
   def write_html(self):
       html = '''
        <html>
        <head>
        
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
        <script data-cfasync="false" src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
        <script type="text/javascript" src="/eel.js"></script>
        <title>Arrival Document</title>
        </head>
        <body>
        <div style='display:none'>
        <form id="foo">
        <p><label for="sid">sid:</label>
        <input id="name" name="name" value="{0}" type="text" ></p>
        <p><label for="name">tgl:</label>
        <input id="apa" name="tgl" value="{1}" type="text" ></p>
        <p><label for="comment">waktu:</label><br>
        <input id="comment" name="waktu" cols="40" value="{2}" type="text" ></input></p>
        <p id="result"></p>
        <input id="sendya" value="Send" type="submit">
        </form>
        </div>
        <p id="info">Hi <b>{0}</b>,<br>I know you arrive <b>{3}</b>. please wait till this window close itself <br>(Retry in: <span id='timer'>00:00</span> minutes)<br><i>If it already > 15 minutes, please close and restart (WIN+R > arr)</i></p>
        

        <script data-cfasync="false" type="text/javascript">
        jQuery( document ).ready(function( $ ) {{
         
         // variable to hold request
         var request;
         // bind to the submit event of our form
         $("#foo").submit(function(event){{
          // abort any pending request
          if (request) {{
           request.abort();
          }}
          // setup some local variables
          var $form = $(this);
          
          // let's select and cache all the fields
          var $inputs = $form.find("input, select, button, textarea");
          // serialize the data in the form
          var serializedData = $form.serialize();
          console.log(serializedData)
          // let's disable the inputs for the duration of the ajax request
          // Note: we disable elements AFTER the form data has been serialized.
          // Disabled form elements will not be serialized.
          $inputs.prop("disabled", true);
          $('#result').text('Sending data...');
          //
          //https://script.google.com/macros/s/AKfycbzV--xTooSkBLufMs4AnrCTdwZxVNtycTE4JNtaCze2UijXAg8/exec
          // fire off the request to /form.php
          request = $.ajax({{
           url: "https://script.google.com/macros/s/AKfycbw9T9iSJ_RM3QrCLhnBrD5qPSHLEEY4OsoLa_Tt_weLGyjkQAsm/exec",
           type: "post",
           data: serializedData
          }});
         
          // callback handler that will be called on success
          request.done(function (response, textStatus, jqXHR){{
           // log a message to the console
           $('#result').html('<a href="https://docs.google.com/spreadsheets/d/1JyHxit11ZMAwnN7eR_6_wDrz2qoNfOi98nfblRQPqgk/edit?ts=585782c7#gid=1680249107" target="_blank">Success - see Google Sheet</a>');
           
           if (response.result == 'error') {{
              console.log("badabumm... retry....!");
              $('#info').text('Error occured on server... Retrying... if it persist till ~30 minute, please restart tool with WIN+R > arr > ENTER');
              //$('#sendya').submit();
            }}
           //$('#hore').innerText = 'Hooray I know you you come come '
           if (response.result == 'success') {{
             console.log("Hooray, it worked!");
             $('#info').text('Your arrival data successfully sent!')
             window.close();
             document.title = "Arrival filled successfully"
             window.location = 'http://www.google.com'
             }}
          }});
         
          // callback handler that will be called on failure
          request.fail(function (jqXHR, textStatus, errorThrown){{
           // log the error to the console
           console.error(
            "The following error occured: "+
            textStatus, errorThrown
           );
           // window.close();
           $('#result').text('Oh nooo... Sending data...');
           $('#sendya').submit();
          }});
         
          // callback handler that will be called regardless
          // if the request failed or succeeded
          request.always(function () {{
           // reenable the inputs
           $inputs.prop("disabled", false);
          }});
         
          // prevent default posting of form
          event.preventDefault();
         }});
        }});
        $(document).ready(function() {{
            $('#sendya').submit();
        }});
        function startTimer(duration, display) {{
            var timer = duration, minutes, seconds;
            setInterval(function () {{
                minutes = parseInt(timer / 60, 10);
                seconds = parseInt(timer % 60, 10);

                minutes = minutes < 10 ? "0" + minutes : minutes;
                seconds = seconds < 10 ? "0" + seconds : seconds;

                display.text(minutes+':'+seconds);

                if (--timer < 0) {{
                    $('#info').text('Error on server.. Retrying...!')
                }}        
                if (timer < -3) {{
                   location.reload();
                }}
            }}, 1000);
        }}

        jQuery(function ($) {{
            var fiveMinutes = 60 * 5,
                display = $('#timer');
            startTimer(fiveMinutes, display);
        }});     
        </script>


        </body>
        </<html>
        '''.format(self.name, self.tglsheet, self.time, self.is_late)
        
       if not os.path.exists(web):
          os.makedirs(web)
       os.chdir(os.getcwd() + r'/' + web)
       today = open(web+'.html', 'w')
       today.write(html)
       today.close()


   def delete_html(self):
     print('deleting any trace...')
     os.remove(web+'.html')
     self.logging('sent_success')

   
   def send_data(self):
      print('i try to sent data of:', self.name, 'arrive at :',self.time)
      self.write_html()
      web_app_options = {
        'mode': "chrome-app", #or "chrome"
        'port': 8081,
        'chromeFlags': ["--window-size=500,112",  "--incognito"]
        }
      print('start eel')
      try:
        eel.start(web+'.html', options=web_app_options)
      except SystemExit:
        self.delete_html()
      
   
def updateclient():
  try:
      arr_file = r'//'+SERVER+r'/windowsteam/script/arr.exe'
      user_arr = 'c:\\users\\'+ getpass.getuser() +'\\arr.exe'  
      latest_ver = open(r'//'+SERVER+r'/windowsteam/script/autoarrivalver.txt')
      latest = latest_ver.read()
      latest_ver.close()
      latest_ver = latest
      
      #if arr_file_size != user_arr_size:
      if app_version != latest_ver:
       user_update = 'c:\\users\\'+ getpass.getuser()+'\\update.exe'
       update_file = r'//'+SERVER+r'/windowsteam/script/update.exe'
       
       update = open(update_file, 'rb')
       copy_update = open(user_update, 'wb')
       copy_update.write(update.read())
       update.close()
       copy_update.close()  
       print('updating autoarrival')
       import win32api
       win32api.MessageBox(0, 'AutoArrival need to update to latest ({}) your app version : {}, press OK to update'.format(latest_ver, app_version), 'AutoArrival Updater')
       os.startfile(user_update)
      else:
       print('autoarrival already uptodate')
  except:
       print('access denied')

if __name__ == '__main__':
   i_am_arrive = ArrivalCahGL()
   i_am_arrive.send_data()
   if i_am_arrive.is_windows:
     print('checkint update')
     updateclient()
     print('done')
