import sys
from pprint import pprint
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

        netconnect = ConnectHandler(**connecthandler)
        result = netconnect.send_command(self.cmd)
        netconnect.disconnect()
        if len(result) == 0:
            return None
        else:
            return result

    def mac_port(self):
        trk_list = self.trunk_port()
        if self.dev_type == 'cisco_xe':
            list_trk = list()
            for trk in trk_list:
                te_re = re.compile(r'Te[\d\/]+')
                te = te_re.search(trk)
                ge_re = re.compile(r'Gi[\d\/]+')
                ge = ge_re.search(trk)
                fa_re = re.compile(r'Fa[\d\/]+')
                fa = fa_re.search(trk)
                if te:
                    int_te = trk.replace('Te','TenGigabitEthernet')
                    list_trk.append(int_te)
                elif ge:
                    int_te = trk.replace('Gi','GigabitEthernet')
                    list_trk.append(trk)
                elif fa:
                    int_te = trk.replace('Fa','FastEthernet')
                    list_trk.append(trk)
       
            mac_cmd = 'show mac address-table dynamic | exclude %s' % '|'.join(list_trk)
            mac_output = self.command(mac_cmd)
            mac_re = re.compile(r'\w+\.\w+\.\w+')
            int_re = re.compile(r'(FastEthernet[\/\d]+|GigabitEthernet[\/\d]+|TengigabitEthernet[\/\d]+)')
            vlan_re = re.compile(r'(\d+|Te[\d\/]+|Gi[\/\d]+|Fa[\/\d+].*)')
            mac_line = mac_output.splitlines()
            mac_int_list = list()

            for line in mac_line:
                data_mac = {
                        'mac' : None,
                        'port' : None,
                        'vlan' : None,
                        'switch' : None,
                        'addr' : None
                        }

                mac = mac_re.search(line)
                port = int_re.search(line)
                vlan = vlan_re.search(line)

                if mac and port and vlan:
                    data_mac['mac'] = mac.group()
                    data_mac['port'] =port.group(1)
                    data_mac['vlan'] = vlan.group(1)
                    data_mac['switch'] = self.hostname

                    mac_int_list.append(data_mac)
            return mac_int_list

        else:
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
                        'vlan' : None,
                        'switch' : None,
                        'addr': None
                        }
                mac = mac_re.search(line)
                port = int_re.search(line)
                vlan = vlan_re.search(line)

                if mac and port and vlan:
                    data_mac['mac'] = mac.group()
                    data_mac['port'] = port.group(1)
                    data_mac['vlan'] = vlan.group(1)
                    data_mac['switch'] = self.hostname

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

def int_vlan(output):
    intvlan_re = re.compile(r'Vlan(\d+)')
    intvlan = intvlan_re.findall(output)

    return intvlan
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
            dev_inventories = inventory(invent_file)
            address_invent = 'data/address_inventory.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.cell(row=1, column=1, value='Hostname')
            ws.cell(row=1, column=2, value='Port')
            ws.cell(row=1, column=3, value='Mac Address')
            ws.cell(row=1, column=4, value='Vlan')
            ws.cell(row=1, column=5, value='IP Address')
            wb.save(address_invent)

            int_br_cmd = 'show ip interface brief | include .+\..+\..+\..+'
            inventories = list()
            print('identifying interface vlan, please wait')
            for a in dev_inventories:
                try:
                    int_br_conn = SendCommand(
                            hostname=a['hostname'],
                            dev_type=a['type'],
                            username=a['user'],
                            password=a['pass'],
                            address=a['address']
                            )
                    
                    int_br_output = int_br_conn.command(int_br_cmd)
                    int_vlans = int_vlan(int_br_output)
                    if len(int_vlans) != 0:
                        a['int_vlan'] = int_vlans 
                    
                    inventories.append(a)
                except:
                    continue

            #pprint(inventories)

            for d in inventories:
                output = SendCommand(d['hostname'], d['type'], d['address'], d['user'], d['pass'])
                macs = output.mac_port()
                #pprint('%s -> %s '% (d['hostname'], macs))
                row = 2
                for mac in macs:
                    mac_add = mac['mac']
                    
                    for e in inventories:
                        if mac['vlan'] in e['int_vlan']:
                            arp_cmd = 'show ip arp | include %s' % mac_add
                            #print('connect to %s' % e['hostname'])
                            conn = SendCommand(e['hostname'], e['type'], e['address'], e['user'], e['pass'])
                            arp_output = conn.command(arp_cmd)
                            if arp_output == None:
                                continue
                            else:
                                arp = arp_addr(arp_output)
                                mac['addr'] = arp
                                break
                        else:
                            continue
                    print(mac)
                    wb = load_workbook(address_invent)
                    ws = wb.active
                    ws.cell(row=row, column=1, value=mac['switch'])
                    ws.cell(row=row, column=2, value=mac['port'])
                    ws.cell(row=row, column=3, value=mac['mac'])
                    ws.cell(row=row, column=4, value=mac['vlan'])
                    if mac['addr'] != None:
                        ws.cell(row=row, column=5, value=','.join(mac['addr']))

                    wb.save(address_invent)
                    row += 1
                    log('%s : %s\n' % (time_log(),mac), data_log)
        else:
            print('sorry, please choose function above..')
            time.sleep(3)

if __name__ == '__main__':
    main()
