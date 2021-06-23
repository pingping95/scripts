from settings import Common


class TransitPeering(Common):
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
        response = self.client.describe_transit_gateways()["TransitGateways"]
        response2 = self.client.describe_vpc_peering_connections()["VpcPeeringConnections"]
        if len(response) != 0:
            try:
                # Initialize
                self.sheet = self.wb.create_sheet(self.name)
                self.sheet.title = f"{self.name}"
                # Cell width
                cell_widths = [5, 5, 25, 23, 22, 25, 16, 20, 22, 15, 20, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7]
                self.fit_cell_width(cell_widths)
                # Header
                self.make_header(self.cell_start, "Transit")
                # Cell header
                cell_headers = ["No.", "Name", "State", "Transit Gateway ID", "Creation Date", "CIDR Blocks"]
                self.make_cell_header(self.cell_start, cell_headers)
                
                # For loop
                for idx, tgw in enumerate(response):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Name
                    tgw_name = "-"
                    try:
                        for i in tgw.get("Tags"):
                            if i.get("Key") == "Name":
                                tgw_name = i.get("Value")
                    except Exception:
                        pass
                    self.add_cell(self.cell_start, 3, tgw_name)
                    # State
                    self.add_cell(self.cell_start, 4, tgw.get("State"))
                    # Transit Gateway ID
                    self.add_cell(self.cell_start, 5, tgw.get("TransitGatewayId"))
                    # Creation Date
                    tgw_date = str(tgw.get("CreationTime")).split(" ")[0]
                    self.add_cell(self.cell_start, 6, tgw_date)
                    # CIDR Blocks
                    try:
                        tgw_cidr = ""
                        for i, cidr in enumerate(tgw["Options"]["TransitGatewayCidrBlocks"]):
                            if i != 0:
                                tgw_cidr += ", \n"
                            tgw_cidr += cidr
                        self.add_cell(self.cell_start, 7, tgw_cidr)
                    except Exception as e:
                        self.add_cell(self.cell_start, 7, "-")
                    self.cell_start += 1
            except Exception as e:
                self.log.write(f"Error 발생, 리소스: TGW, 내용: {e}\n")
        else:
            self.log.write(f"There is no Transit Gateway\n")
        
        if len(response) == 0 and len(response2) != 0:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = f"{self.name}"
        if len(response2) != 0:
            try:
                self.cell_start += 1
                # Cell width
                cell_widths = [5, 5, 25, 23, 22, 25, 16, 20, 22, 15, 20, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7]
                self.fit_cell_width(cell_widths)
                # Header
                self.make_header(self.cell_start, "Peering Connection")
                cell_headers = ["No.","Name","Status","Peering Connection ID","Requester VPC ID","Requester Region",
                "Requester CIDR Block","Accepter VPC ID","Accepter Region","Accepter CIDR Block"]
                self.make_cell_header(self.cell_start, cell_headers)
                for idx, peering in enumerate(self.client.describe_vpc_peering_connections()["VpcPeeringConnections"]):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Name
                    peer_name = "-"
                    try:
                        for i in peering.get("Tags"):
                            if i.get("Key") == "Name":
                                peer_name = i.get("Value")
                    except Exception:
                        pass
                    self.add_cell(self.cell_start, 3, peer_name)
                    # Status, Accepter CIDR Block
                    # Statue가 Active가 아닐 경우 Accepter의 CIDR Block을 알아낼 수 없음
                    peer_status = peering["Status"].get("Code")
                    if peer_status == "active":
                        self.add_cell(self.cell_start, 4, peer_status)
                        # Accepter CIDR Block
                        accepter_cidr = peering["AccepterVpcInfo"].get("CidrBlock")
                        self.add_cell(self.cell_start, 11, accepter_cidr)
                    else:
                        self.add_cell(self.cell_start, 4, peer_status)
                        self.add_cell(self.cell_start, 11, "-")
                    # Peering Connection ID
                    self.add_cell(self.cell_start, 5, peering.get("VpcPeeringConnectionId"))
                    # Requester
                    requester = peering.get("RequesterVpcInfo")
                    # Requester VPC ID
                    req_id = requester.get("VpcId")
                    self.add_cell(self.cell_start, 6, req_id)
                    # Requester Region
                    self.add_cell(self.cell_start, 7, requester.get("Region"))
                    # Requester CIDR Block
                    self.add_cell(self.cell_start, 8, requester.get("CidrBlock"))
                    # Accepter
                    accepter = peering.get("AccepterVpcInfo")
                    # Accepter VPC ID
                    self.add_cell(self.cell_start, 9, accepter.get("VpcId"))
                    # Accepter Region
                    self.add_cell(self.cell_start, 10, accepter.get("Region"))

                    self.cell_start += 1
            except Exception as e:
                self.log.write(f"Error 발생, 리소스: Peering Connection, 내용: {e}\n")
        else:
            self.log.write(f"There is no Peering Connection\n")