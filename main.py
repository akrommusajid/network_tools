import sys
from openpyxl import load_workbook
from openpyxl import Workbook
from netmiko import ConnectHandler
import re
import time
from datetime import datetime


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

def arp_addr(output):
    arp_re = re.compile(r'\d+\.\d+\.\d+\.\d+')
    arp = arp_re.findall(output)
    return arp

class SendCommand(object):
    def __init__(self, hostname, dev_type, address, username, password):
        self.hostname = hostname
        self.dev_type = dev_type
        self.address = address
        self.username = username
        self.password = password

    def command(self, cmd):
        self.cmd = cmd
        connecthandler = {
            'device_type' : self.dev_type,
            'ip' : self.address,
            'username' : self.username,
            'password' : self.password
        }

        try:
            netconnect = ConnectHandler(**connecthandler)
            result = netconnect.send_command(self.cmd)
            netconnect.disconnect()
            if len(result) == 0:
                return None
            else:
                return result

        except:
            result = 'sorry connection to %s was failed' % (self.hostname)
            return result
        '''
        except (EOFError, SSHException):
            return None
        '''
    def mac_port(self):
        trk_list = self.trunk_port()
        mac_cmd = 'show mac address-table dynamic | exclude %s' % '|'.join(trk_list)
        mac_output = self.command(mac_cmd)
        mac_re = re.compile(r'\w+\.\w+\.\w+')
        int_re = re.compile(r'(Fa[\/\d]+|Gi[\/\d]+|Te[\/\d]+)')
        vlan_re = re.compile(r'(\d+).*')
        mac_line = mac_output.splitlines()
        mac_int_list = list()
        
        for line in mac_line:
            data_mac = {
                    'mac' : None,
                    'port' : None,
                    'vlan' : None
                    }
            mac = mac_re.search(line)
            port = int_re.search(line)
            vlan = vlan_re.search(line)

            if mac and port and vlan:
                data_mac['mac'] = mac.group()
                data_mac['port'] = port.group(1)
                data_mac['vlan'] = vlan.group(1)

                mac_int_list.append(data_mac)

        return mac_int_list

    def trunk_port(self):
        local_cmd = 'show interface trunk | i trunking'
        trunk_cmd = self.command(local_cmd)
        trk_int_re = re.compile(r'(\w+[\/\d]+).*')
        trk_int = trk_int_re.findall(trunk_cmd)

        return trk_int

    def sw_version(self):
        local_cmd = 'show version'
        result = self.command(local_cmd)

        if self.dev_type == 'cisco_ios':
            ios_sw_re = re.compile(r'Version (\S+),')
            ios_pn_re = re.compile(r'Model\s+number\s+:\s+(\S+)')

            ios_sw = ios_sw_re.search(result)
            ios_pn = ios_pn_re.search(result)

            if ios_sw:
                sw = ios_sw.group(1)
                if ios_pn:
                    pn = ios_pn.group(1)
                    return sw,pn
        elif self.dev_type == 'cisco_xe':
            xe_sw_re = re.compile(r'Version (\S+)')
            xe_pn_re = re.compile(r'cisco (WS-\S+)')

            xe_sw = xe_sw_re.search(result)
            xe_pn = xe_pn_re.search(result)

            if xe_sw:
                sw = xe_sw.group(1)
                if xe_pn:
                    pn = xe_pn.group(1)
                    return sw,pn
        elif self.dev_type == 'cisco_nx':
            nx_sw_re = re.compile(r'system:\s+version\s+(\S+)')
            nx_pn_re = re.compile(r'cisco (Nexus\w+ \S+)')

            nx_sw = nx_sw_re.search(result)
            nx_pn = nx_pn_re.search(result)

            if nx_sw:
                sw = nx_sw.group(1)
                if nx_pn:
                    pn = nx_pn.group(1)
                    return sw,pn

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
            cmd = 'show version'
            row = 2
            for d in inventories:
                output = SendCommand(d['hostname'], d['type'], d['address'], d['user'], d['pass'])
                sw_inventory = output.sw_version() 
                print('%s : %s --> partnumber : %s, sw version : %s' % (time_log(), d['hostname'], sw_inventory[1], sw_inventory[0]))
                log('%s : %s --> %s\n' % (time_log(), d['hostname'], sw_inventory), data_log) 
                
                '''load workbook'''
                wb = load_workbook(swinvent_dir)
                ws = wb.active
                ws.cell(row=row, column=1, value=d['hostname'])
                ws.cell(row=row, column=2, value=sw_inventory[1])
                ws.cell(row=row, column=3, value=sw_inventory[0])
                wb.save(swinvent_dir)
                row += 1
        elif input_select == '2':
            inventories = inventory(invent_file)
            for d in inventories:
                output = SendCommand(d['hostname'], d['type'], d['address'], d['user'], d['pass'])
                macs = output.mac_port()
                for mac in macs:
                    mac_add = mac['mac']

                    for e in inventories:
                        arp_cmd = 'show ip arp | include %s' % mac_add
                        print('connect to %s' % e['hostname'])
                        conn = SendCommand(e['hostname'], e['type'], e['address'], e['user'], e['pass'])
                        arp_output = conn.command(arp_cmd)
                        if arp_output == None:
                            continue
                        else:
                            arp = arp_addr(arp_output)
                            mac['addr'] = arp
                            break
                        #print(e['address'],e['user'],e['pass'],e['type'])
                    mac['switch'] = d['hostname']
                    print(mac)
                    log('%s : %s\n' % (time_log(),mac), data_log)


        else:
            print('sorry, please choose function above..')
            time.sleep(3)

if __name__ == '__main__':
    main()
