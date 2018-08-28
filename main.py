import sys
from openpyxl import load_workbook
from openpyxl import Workbook
from netmiko import ConnectHandler
import re
import time
from datetime import datetime

def sw_version(result, type):
    if type == 'cisco_ios':
        ios_sw_re = re.compile(r'Version (\S+),')
        ios_pn_re = re.compile(r'Model\s+number\s+:\s+(\S+)')

        ios_sw = ios_sw_re.search(result)
        ios_pn = ios_pn_re.search(result)

        if ios_sw:
            sw = ios_sw.group(1)
            if ios_pn:
                pn = ios_pn.group(1)
                return sw,pn
    elif type == 'cisco_xe':
        xe_sw_re = re.compile(r'Version (\S+)')
        xe_pn_re = re.compile(r'cisco (WS-\S+)')

        xe_sw = xe_sw_re.search(result)
        xe_pn = xe_pn_re.search(result)

        if xe_sw:
            sw = xe_sw.group(1)
            if xe_pn:
                pn = xe_pn.group(1)
                return sw,pn
    elif type == 'cisco_nx':
        nx_sw_re = re.compile(r'system:\s+version\s+(\S+)')
        nx_pn_re = re.compile(r'cisco (Nexus\w+ \S+)')

        nx_sw = nx_sw_re.search(result)
        nx_pn = nx_pn_re.search(result)

        if nx_sw:
            sw = nx_sw.group(1)
            if nx_pn:
                pn = nx_pn.group(1)
                return sw,pn

def time_log():
    time_log_stamp = datetime.now()
    log_time = time_log_stamp.strftime('%Y-%m-%d %H:%M:%S')
    return log_time

def log(write, filename):
    with open(filename, 'a') as writelog:
        writelog.write(write)

def inventory(filename):
    wb = load_workbook(filename)
    ws = wb.active
    invents = list()
    for i in range(2, ws.max_row+1):
        invent_dic = {
            'hostname': None,
            'address': None,
            'type' : None,
            'user': None,
            'pass': None
        }

        invent_dic['hostname'] = ws.cell(row=i, column=1).value
        invent_dic['address'] = ws.cell(row=i, column=2).value
        invent_dic['type'] = ws.cell(row=i, column=3).value
        invent_dic['user'] = ws.cell(row=i, column=4).value
        invent_dic['pass'] = ws.cell(row=i, column=5).value

        invents.append(invent_dic)
    return invents

def send_command(hostname, type, address, username, password, command):
    connecthandler = {
        'device_type' : type,
        'ip' : address,
        'username' : username,
        'password' : password
    }

    try:
        netconnect = ConnectHandler(**connecthandler)
        result = netconnect.send_command(command)
        return result

    except Exception as msg:
        result = 'sorry connection to %s was failed, %s' % (hostname, msg)
        return result

def main():
    invent_file = 'data/inventory.xlsx'
    data_log = 'syslog/log.txt'
    while True:
        print('\n\n\nthis tools is used to help you to collect data related with network')
        print('this tools is still limited based on Cisco device, Sorry :(')
        print('please make sure you have fill inventory data in "data" directory')
        print('[1] collect software version')
        print('[2] collect address based on arp')
        print('[q] exit\n\n')
        input_select = input('please select function above :')
        input_select = str(input_select)

        if input_select == 'q' or input_select == 'Q':
            sys.exit()
        elif input_select == '1':
            '''create file software inventory'''
            wb = Workbook()
            ws = wb.active
            swinvent_dir = 'data/software_inventory.xlsx'
            ws.cell(row=1, column=1, value='Hostname')
            ws.cell(row=1, column=2, value='Part Number')
            ws.cell(row=1, column=3, value='Software Version')

            wb.save(swinvent_dir)
            inventories = inventory(invent_file)
            command = 'show version'
            row = 2
            for d in inventories:
                output = send_command(d['hostname'], d['type'], d['address'], d['user'], d['pass'], command)
                sw_inventory = sw_version(output, d['type'])
                print('%s : %s --> partnumber : %s, sw_version : %s' % (time_log(), d['hostname'], sw_inventory[1], sw_inventory[0]))
                log('%s : %s --> %s\n' % (time_log(), d['hostname'], sw_inventory), data_log) 
                
                '''load workbook'''
                wb = load_workbook(swinvent_dir)
                ws = wb.active
                ws.cell(row=row, column=1, value=d['hostname'])
                ws.cell(row=row, column=2, value=sw_inventory[1])
                ws.cell(row=row, column=3, value=sw_inventory[0])
                wb.save(swinvent_dir)
                row += 1
        else:
            print('sorry, please choose function above..')
            time.sleep(3)

if __name__ == '__main__':
    main()
