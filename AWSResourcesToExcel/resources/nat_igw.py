from settings import Common

class NatIgw(Common):
    def __init__(self, name, workbook, ses, p_name, r_name, log, is_run = False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = p_name
            self.region = r_name
            self.client = ses.client(service_name="ec2", region_name=self.region)
            self.run()
    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = f"{self.name}"
            # Cell width
            cell_widths = [5, 5, 27, 22, 22, 22, 24, 20, 20, 20, 20, 20, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, "Nat Gateway")
            # Cell header
            cell_headers = ["No.", "Name", "ID", "Elastic IP", "VPC", "Subnet", "Status"]
            self.make_cell_header(self.cell_start, cell_headers)
            # For loop
            if len(self.client.describe_nat_gateways()["NatGateways"]) != 0:
                for idx, nat in enumerate(self.client.describe_nat_gateways()["NatGateways"]):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Name
                    natName = "-"
                    try:
                        for i in nat.get("Tags"):
                            if i.get("Key") == "Name":
                                natName = i.get("Value")
                    except Exception:
                        pass
                    id = nat.get("NatGatewayId")
                    eip = nat.get("NatGatewayAddresses")[0].get("PublicIp")
                    vpc = nat.get("VpcId")
                    subnet = nat.get("SubnetId")
                    state = nat.get("State")
                    self.add_cell(self.cell_start, 3, natName)
                    self.add_cell(self.cell_start, 4, id)
                    self.add_cell(self.cell_start, 5, eip)
                    self.add_cell(self.cell_start, 6, vpc)
                    self.add_cell(self.cell_start, 7, subnet)
                    self.add_cell(self.cell_start, 8, state.capitalize())
                    self.cell_start += 1
            else:
                self.log.write(f"There is no Nat Gateway\n")
            self.cell_start += 1
            # Header
            self.make_header(self.cell_start, "IGW")
            # Cell header
            cell_headers = ["No.", "Name", "ID", "VPC", "Status"]
            self.make_cell_header(self.cell_start, cell_headers)
            if len(self.client.describe_internet_gateways()["InternetGateways"]) != 0:
                for idx, igw in enumerate(self.client.describe_internet_gateways()["InternetGateways"]):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Name
                    igwName = "-"
                    try:
                        for i in igw.get("Tags"):
                            if i.get("Key") == "Name":
                                igwName = i.get("Value")
                    except Exception:
                        pass
                    # igw ID, vpc ID, Status
                    igwId = igw.get("InternetGatewayId")
                    if len(igw["Attachments"]) != 0:
                        vpcId = igw["Attachments"][0].get("VpcId")
                        status = igw["Attachments"][0].get("State")
                    else:
                        vpcId = "-"
                        status = "Detached"
                    # Name
                    self.add_cell(self.cell_start, 3, igwName)
                    # IGW ID
                    self.add_cell(self.cell_start, 4, igwId)
                    # VPC ID
                    self.add_cell(self.cell_start, 5, vpcId)
                    # Status
                    self.add_cell(self.cell_start, 6, status.capitalize())
                    self.cell_start += 1
            else:
                self.log.write(f"There is no Internet Gateways\n")
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")