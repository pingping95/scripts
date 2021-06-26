from settings import Common

class Nacl(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run = False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="ec2", region_name=self.region)
            self.run()

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = f"{self.name}"
            # Cell width
            cell_widths = [5, 5, 18, 25, 25, 25, 17, 15, 20, 20, 5, 5, 18, 24, 24, 13, 13, 13, 13, 13, 7, 7]
            self.fit_cell_width(cell_widths)
            # 1. NACL Inbound
            # Header
            self.make_header(self.cell_start, self.name + " (Inbound)")
            # Cell header
            cell_headers = ["No.", "Network ACL Name", "Network ACL Id", "VPC ID", "Rule", "Protocol", "Port Range", "Source","Allow / Deny",]
            self.make_cell_header(self.cell_start, cell_headers)
            # For loop
            for idx, acl in enumerate(self.client.describe_network_acls()["NetworkAcls"]):
                acl_start_row = self.cell_start
                
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Name
                acl_name = "-"
                try:
                    for tag in acl.get('tags'):
                        if tag.get('Key') == "Name":
                            acl_name = tag.get('Value')
                except:
                    pass
                self.add_cell(self.cell_start, 3, acl_name)
                self.add_cell(self.cell_start, 4, acl.get('NetworkAclId'))
                self.add_cell(self.cell_start, 5, acl.get('VpcId'))
                in_count = 0
                for entry in acl.get('Entries'):
                    # Ingress일 경우
                    if entry.get('Egress') == False:
                        in_count += 1
                        
                        if "CidrBlock" in entry:
                            source = entry.get('CidrBlock')
                        else:
                            source = "-"
                        protocol = "All"
                        portRange = "All"
                        
                        # # Rule Number
                        if entry["RuleNumber"] >= 32737:
                            ruleNumber = "*"
                        else:
                            ruleNumber = entry["RuleNumber"]
                                
                        # RuleAction
                        ruleAction = entry["RuleAction"]
                        
                        # Protocol, Port Range
                        if entry["Protocol"] == "-1":
                            protocol = "All"
                        elif entry["Protocol"] == "6":
                            protocol = "TCP"
                            if entry["PortRange"]["To"] == entry["PortRange"]["From"]:
                                portRange = entry["PortRange"]["From"]
                            else:
                                portRange = (str(entry["PortRange"]["From"]) + " - " + str(entry["PortRange"]["To"]))
                        elif entry["Protocol"] == "17":
                            protocol = "UDP"
                            if entry["PortRange"]["To"] == entry["PortRange"]["From"]:
                                portRange = entry["PortRange"]["From"]
                            else:
                                portRange = (str(entry["PortRange"]["From"]) + " - " + str(entry["PortRange"]["To"]))
                    
                        self.add_cell(self.cell_start, 6, ruleNumber)
                        self.add_cell(self.cell_start, 7, protocol)
                        self.add_cell(self.cell_start, 8, portRange)
                        self.add_cell(self.cell_start, 9, source)
                        self.add_cell(self.cell_start, 10, ruleAction.capitalize())
                        self.cell_start += 1
                acl_end_row = self.cell_start - 1
                # Merge Cells
                self.sheet.merge_cells(start_row=acl_start_row, end_row=acl_end_row, start_column=2, end_column=2)
                self.sheet.merge_cells(start_row=acl_start_row, end_row=acl_end_row, start_column=3, end_column=3)
                self.sheet.merge_cells(start_row=acl_start_row, end_row=acl_end_row, start_column=4, end_column=4)
                self.sheet.merge_cells(start_row=acl_start_row, end_row=acl_end_row, start_column=5, end_column=5)            
            # 2. NACL Outbound
            self.cell_start += 1
            # Header
            self.make_header(self.cell_start, self.name + " (Outbound)")
            # Cell header
            cell_headers = ["No.","Network ACL Name","Network ACL Id","VPC ID","Rule","Protocol","Port Range","Source","Allow / Deny"]
            self.make_cell_header(self.cell_start, cell_headers)
            # For loop
            for idx, acl in enumerate(self.client.describe_network_acls()["NetworkAcls"]):
                acl_start_row = self.cell_start
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Name
                acl_name = "-"
                try:
                    for tag in acl.get('tags'):
                        if tag.get('Key') == "Name":
                            acl_name = tag.get('Value')
                except:
                    pass
                self.add_cell(self.cell_start, 3, acl_name)
                self.add_cell(self.cell_start, 4, acl.get('NetworkAclId'))
                self.add_cell(self.cell_start, 5, acl.get('VpcId'))
                in_count = 0
                for entry in acl.get('Entries'):
                    # Egress일 경우
                    if entry.get('Egress') == True:
                        in_count += 1
                        
                        if "CidrBlock" in entry:
                            source = entry.get('CidrBlock')
                        else:
                            source = "-"
                        protocol = "All"
                        portRange = "All"
                        # # Rule Number
                        if entry["RuleNumber"] >= 32737:
                            ruleNumber = "*"
                        else:
                            ruleNumber = entry["RuleNumber"]
                        # RuleAction
                        ruleAction = entry["RuleAction"]
                        # Protocol, Port Range
                        if entry["Protocol"] == "-1":
                            protocol = "All"
                        elif entry["Protocol"] == "6":
                            protocol = "TCP"
                            if entry["PortRange"]["To"] == entry["PortRange"]["From"]:
                                portRange = entry["PortRange"]["From"]
                            else:
                                portRange = (str(entry["PortRange"]["From"]) + " - " + str(entry["PortRange"]["To"]))
                        elif entry["Protocol"] == "17":
                            protocol = "UDP"
                            if entry["PortRange"]["To"] == entry["PortRange"]["From"]:
                                portRange = entry["PortRange"]["From"]
                            else:
                                portRange = (str(entry["PortRange"]["From"]) + " - " + str(entry["PortRange"]["To"]))
                    
                        self.add_cell(self.cell_start, 6, ruleNumber)
                        self.add_cell(self.cell_start, 7, protocol)
                        self.add_cell(self.cell_start, 8, portRange)
                        self.add_cell(self.cell_start, 9, source)
                        self.add_cell(self.cell_start, 10, ruleAction.capitalize())
                
                        self.cell_start += 1
                    
                acl_end_row = self.cell_start - 1
                # Merge Cells
                self.sheet.merge_cells(start_row=acl_start_row, end_row=acl_end_row, start_column=2, end_column=2)
                self.sheet.merge_cells(start_row=acl_start_row, end_row=acl_end_row, start_column=3, end_column=3)
                self.sheet.merge_cells(start_row=acl_start_row, end_row=acl_end_row, start_column=4, end_column=4)
                self.sheet.merge_cells(start_row=acl_start_row, end_row=acl_end_row, start_column=5, end_column=5)            
            
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")