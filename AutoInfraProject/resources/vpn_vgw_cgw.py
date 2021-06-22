from .settings import Common
import openpyxl
import boto3

class VpnVgwCgw(Common):
    def __init__(self, name, workbook, ses, info, log, is_run = False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = info.get('porfile')
            self.region = info.get('region')
            self.client = ses.client(service_name="ec2", region_name=self.region)
            self.run()

    def run(self):
        if len(self.client.describe_vpn_connections()["VpnConnections"]) != 0:
            try:
                print(f"name: {self.name}, profile: {self.profile}, res: {self.client}, reg : {self.region}")
                # Initialize
                self.sheet = self.wb.create_sheet(self.name)
                self.sheet.title = f"{self.name}"
                # Cell width
                cell_widths = [5, 5, 27, 22, 22, 22, 24, 20, 20, 20, 20, 20, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7]
                self.fit_cell_width(cell_widths)
                
                # Header
                self.make_header(self.cell_start, "VPN Connections")
                # Cell header
                cell_headers = ["No.","Name","VPN ID","State","Virtual Private Gateway","Transit Gateway",
                                "Customer Gateway", "Routing", "Type", "Local IPv4 CIDR", "Remote IPv4 CIDR"]
                self.make_cell_header(self.cell_start, cell_headers)
                
                # For loop
                # for idx, tgw in enumerate():

            except Exception as e:
                self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")
        else:
            self.log.write(f"There is no VPN\n")


        if len(self.client.describe_vpn_gateways()["VpnGateways"]) != 0:
            try:
                self.cell_start += 1
                
                # Header
                self.make_header(self.cell_start, "Virtual Private Gateway")
                cell_headers = ["No.", "Name", "ID", "State", "Type", "VPC", "Amazon Side SAN"]
                self.make_cell_header(self.cell_start, cell_headers)
                
                # For Loop
                # for idx, peering in enumerate(self.client.):
                    
            except Exception as e:
                self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")
        else:
            self.log.write(f"There is no VGW\n")
            
        if len(self.client.describe_customer_gateways()["CustomerGateways"]) != 0:
            try:
                self.cell_start += 1
                
                # Header
                self.make_header(self.cell_start, "Customer Gateway")
                cell_headers = ["No.", "Name", "ID", "State", "Type", "IP Address", "BGP ASN", "Device Name"]
                self.make_cell_header(self.cell_start, cell_headers)
                
                # For Loop
                # for idx, peering in enumerate(self.client.):
                    
            except Exception as e:
                self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")
        else:
            self.log.write(f"There is no CGW\n")