import sys
import os
from optparse import OptionParser
from smb.SMBConnection import SMBConnection
from smb import smb_structs
#from socket import gethostname

def connect(username, password, my_name, server_name, server_ip):
    smb_structs.SUPPORT_SMB2 = True
    conn = SMBConnection(username, password, my_name, server_name, use_ntlm_v2 = True)
    try:
        conn.connect(server_ip, 445) #139=NetBIOS / 445=TCP
    except Exception as e:
        print (e)
    return conn

def upload(username, password, my_name, server_name, server_ip, path, filename, service_name):
  conn = connect(username, password, my_name, server_name, server_ip)
  if conn:
    print ('Upload = ' + path + filename)
    print ('Size = %.1f kB' % (os.path.getsize(filename) / 1024.0))
    print ('start upload')
    with open(filename, 'r') as file_obj:
      filesize = conn.storeFile(service_name, path+filename, file_obj)
    print ('upload finished')
    conn.close()

def getServiceName(username, password, my_name, server_name, server_ip):
    conn = connect(username, password, my_name, server_name, server_ip)
    if conn:
        shares = conn.listShares()
        for s in shares:
            print ("s.type : ",s.type)
            print ("s.name : ",s.name)
            print ("s.isSpecial : ",s.isSpecial)
            if not s.isSpecial and s.name not in ['NETLOGON', 'SYSVOL']:
                #sharedfiles = conn.listPath(s.name, '/home/ls_id/smbdir/wt2_down')
                sharedfiles = conn.listPath('wt2_down', '/file')
                for sharedfile in sharedfiles:
                    print(sharedfile.filename)

            if s.type == 0:  # 0 = DISK_TREE
                print('s.name : ', s.name)
                return s.name
        conn.close()


    else:
        return ''

username = 'ls_id'
password = 'K*vkxldh00'
my_name = 'wipsPC'
server_ip = '172.20.20.109'
server_name = 'wt2_down'
#filename = 'raspi-guide.pdf'
#filename = 'smb-example.py'
#new_filename = 'smb-example2.py'
path = '/file'
filename = 'smb-test.txt'
service_name = getServiceName(username, password, my_name, server_name, server_ip)

#upload(username, password, my_name, server_name, server_ip, path, filename, service_name)
