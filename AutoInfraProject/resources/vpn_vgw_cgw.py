from settings import Common


class VpnVgwCgw(Common):
    def __init__(self, name, workbook, ses, p_name, r_name, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = p_name
            self.region = r_name
            self.client = ses.client(service_name="ec2", region_name=self.region)
            self.run()

    def run(self):
        response = self.client.describe_vpn_connections()["VpnConnections"]
        response2 = self.client.describe_vpn_gateways()["VpnGateways"]
        response3 = self.client.describe_customer_gateways()["CustomerGateways"]
        if len(response) != 0 or len(response2) != 0 or len(response3) != 0:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = f"{self.name}"
            # Cell width
            cell_widths = [5, 5, 27, 22, 22, 22, 24, 22, 20, 20, 20, 20, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
        if len(response) != 0:
            try:
                # Header
                self.make_header(self.cell_start, "VPN Connections")
                # Cell header
                cell_headers = ["No.", "Name", "VPN ID", "State", "Virtual Private Gateway", "Transit Gateway",
                                "Customer Gateway", "Routing", "Type", "Local IPv4 CIDR", "Remote IPv4 CIDR"]
                self.make_cell_header(self.cell_start, cell_headers)
                # For loop
                for idx, vpn in enumerate(response):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Name
                    vpn_name = "-"
                    try:
                        for i in vpn.get("Tags"):
                            if i.get("Key") == "Name":
                                vpn_name = i.get("Value")
                    except Exception:
                        pass
                    self.add_cell(self.cell_start, 3, vpn_name)
                    # VPN ID
                    self.add_cell(self.cell_start, 4, vpn.get("VpnConnectionId"))
                    # State
                    vpn_state = vpn.get("State")
                    self.add_cell(self.cell_start, 5, vpn_state.capitalize())
                    # Virtual Private Gateway ID
                    try:
                        self.add_cell(self.cell_start, 6, vpn.get("VpnGatewayId"))
                    except Exception:
                        self.add_cell(self.cell_start, 6, "-")
                    # Transit Gateway ID
                    try:
                        self.add_cell(self.cell_start, 7, vpn.get("TransitGatewayId"))
                    except Exception:
                        self.add_cell(self.cell_start, 7, "-")
                    # Customer Gateway ID
                    try:
                        self.add_cell(self.cell_start, 8, vpn.get("CustomerGatewayId"))
                    except Exception:
                        self.add_cell(self.cell_start, 8, "-")
                    # Routing
                    routing = vpn["Options"]["StaticRoutesOnly"]
                    if routing == False:
                        self.add_cell(self.cell_start, 9, "Dynamic")
                    else:
                        self.add_cell(self.cell_start, 9, "Static")
                    # Type
                    self.add_cell(self.cell_start, 10, vpn.get("Type"))
                    # Local IPv4 CIDR
                    self.add_cell(self.cell_start, 11, vpn["Options"]["LocalIpv4NetworkCidr"])
                    # Remote IPV4 CIDR
                    self.add_cell(self.cell_start, 12, vpn["Options"]["RemoteIpv4NetworkCidr"])
                    self.cell_start += 1

            except Exception as e:
                self.log.write(f"Error 발생, 리소스: VPN Connections, 내용: {e}\n")
        else:
            self.log.write(f"There is no VPN\n")
        if len(response2) != 0:
            try:
                self.cell_start += 1

                # Header
                self.make_header(self.cell_start, "Virtual Private Gateway")
                cell_headers = ["No.", "Name", "ID", "State", "Type", "VPC", "Amazon Side SAN"]
                self.make_cell_header(self.cell_start, cell_headers)
                # For Loop
                for idx, vgw in enumerate(response2):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Name
                    vgw_name = "-"
                    try:
                        for i in vgw.get("Tags"):
                            if i.get("Key") == "Name":
                                vgw_name = i.get("Value")
                        self.add_cell(self.cell_start, 3, vgw_name)
                    except Exception:
                        self.add_cell(self.cell_start, 3, vgw_name)
                    # ID
                    self.add_cell(self.cell_start, 4, vgw.get("VpnGatewayId"))
                    # State
                    self.add_cell(self.cell_start, 5, str(vgw.get("State")).capitalize())
                    # Type
                    self.add_cell(self.cell_start, 6, vgw.get("Type"))
                    # VPC
                    vpc_id = ""
                    try:
                        for i in vgw.get("VpcAttachments"):
                            if i.get("State") == "attached":
                                vpc_id += i.get("VpcId")
                        self.add_cell(self.cell_start, 7, vpc_id)
                    except Exception:
                        self.add_cell(self.cell_start, 7, vpc_id)
                    # Amazon Side SAN
                    try:
                        vgw_san = vgw.get("AmazonSideAsn")
                        self.add_cell(self.cell_start, 8, vgw_san)
                    except Exception:
                        self.add_cell(self.cell_start, 8, "-")
                    self.cell_start += 1

            except Exception as e:
                self.log.write(f"Error 발생, 리소스: VGW, 내용: {e}\n")
        else:
            self.log.write(f"There is no VGW\n")

        if len(response3) != 0:
            try:
                self.cell_start += 1
                # Header
                self.make_header(self.cell_start, "Customer Gateway")
                cell_headers = ["No.", "Name", "ID", "State", "Type", "IP Address", "BGP ASN", "Device Name"]
                self.make_cell_header(self.cell_start, cell_headers)
                # For Loop
                for idx, cgw in enumerate(response3):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Name
                    cgw_name = "-"
                    try:
                        for i in cgw.get("Tags"):
                            if i.get("Key") == "Name":
                                cgw_name = i.get("Value")
                        self.add_cell(self.cell_start, 3, cgw_name)
                    except Exception:
                        self.add_cell(self.cell_start, 3, cgw_name)
                    # CGW ID
                    self.add_cell(self.cell_start, 4, cgw.get("CustomerGatewayId"))
                    # State
                    self.add_cell(self.cell_start, 5, str(cgw.get("State")).capitalize())
                    # Type
                    self.add_cell(self.cell_start, 6, cgw.get("Type"))
                    # IP Address
                    self.add_cell(self.cell_start, 7, cgw.get("IpAddress"))
                    # BGP SAN
                    self.add_cell(self.cell_start, 8, cgw.get("BgpAsn"))

                    # Device Name
                    try:
                        self.add_cell(self.cell_start, 9, cgw["DeviceName"])
                    except Exception:
                        self.add_cell(self.cell_start, 9, "-")
                    self.cell_start += 1
            except Exception as e:
                self.log.write(f"Error 발생, 리소스: CGW, 내용: {e}\n")
        else:
            self.log.write(f"There is no CGW\n")
