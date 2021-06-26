from settings import Common


class Elb(Common):
    def __init__(self, name, workbook, ses, p_name, r_name, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = p_name
            self.region = r_name
            self.client_elb = ses.client(service_name="elb", region_name=self.region)
            self.client_elb2 = ses.client(service_name="elbv2", region_name=self.region)
            response = self.client_elb.describe_load_balancers().get("LoadBalancerDescriptions")
            response2 = self.client_elb2.describe_load_balancers().get('LoadBalancers')
            if len(response) != 0 or len(response2) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [5, 5, 13, 28, 65, 10, 25, 32, 18, 20.5, 12, 14, 11, 7, 7, 7, 7, 7, 7, 7,7,7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Scheme", "ELB Name", "DNS Name", "Type", "Port Configuration", "Instance IDs or Target Groups",
                            "Availability Zones", "ELB Security Group", "Cross-Zone","Idle Timeout (s)", "Access Logs"]
            self.make_cell_header(self.cell_start, cell_headers)
            # 1. CLB
            for idx, response in enumerate(self.client_elb.describe_load_balancers().get("LoadBalancerDescriptions")):
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Scheme
                self.add_cell(self.cell_start, 3, str(response["Scheme"]).capitalize())
                # ELB Name
                self.add_cell(self.cell_start, 4, response["LoadBalancerName"])
                # DNS Name
                self.add_cell(self.cell_start, 5, response["DNSName"])
                # Type
                self.add_cell(self.cell_start, 6, "Classic")
                # Port Configuration
                ListenerDesc = response["ListenerDescriptions"]
                Listenerstr = ""
                for idx, Listener in enumerate(ListenerDesc):
                    if idx != 0:
                        Listenerstr += ", \n"
                    Listenerstr += (
                            str(Listener["Listener"].get("LoadBalancerPort"))
                            + " forwarding to "
                            + str(Listener["Listener"].get("InstancePort"))
                    )
                self.add_cell(self.cell_start, 7, Listenerstr)
                # Instance IDs
                idstr = ""
                if len(response["Instances"]) != 0:
                    for idx, id in enumerate(response["Instances"]):
                        if idx != 0:
                            idstr += ", \n"
                        idstr += id.get("InstanceId")
                    self.add_cell(self.cell_start, 8, idstr)
                else:
                    self.add_cell(self.cell_start, 8, "-")
                # Availability Zones
                zonestr = ""
                for idx, zone in enumerate(response["AvailabilityZones"]):
                    if idx != 0:
                        zonestr += ", \n"
                    zonestr += zone
                self.add_cell(self.cell_start, 9, zonestr)
                # Security Groups
                secstr = ""
                for idx, sec in enumerate(response["SecurityGroups"]):
                    if idx != 0:
                        secstr += ", \n"
                    secstr += sec
                self.add_cell(self.cell_start, 10, secstr)
                response_attr = self.client_elb.describe_load_balancer_attributes(
                    LoadBalancerName=response["LoadBalancerName"])["LoadBalancerAttributes"]
                # Cross-Zone Load Balancing
                crossZone = str(response_attr.get('CrossZoneLoadBalancing').get('Enabled')).lower().capitalize()
                self.add_cell(self.cell_start, 11, crossZone)
                # Idle Timeout
                self.add_cell(self.cell_start, 12, int(response_attr["ConnectionSettings"]["IdleTimeout"]))
                # Access Logs
                access_log = str(response_attr.get('AccessLog').get('Enabled')).lower().capitalize()
                self.add_cell(self.cell_start, 13, access_log)
                self.cell_start += 1
            
            # 2. ALB, NLB, GWLB
            for idx, response in enumerate(self.client_elb2.describe_load_balancers().get('LoadBalancers')):
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Scheme
                schemeType = ""
                if "Scheme" in response:
                    schemeType = str(response["Scheme"]).capitalize()
                if len(schemeType) != 0:
                    self.add_cell(self.cell_start, 3, schemeType)
                else:
                    self.add_cell(self.cell_start, 3, "-")
                # ELB Name
                ElbName = response["LoadBalancerName"]
                self.add_cell(self.cell_start, 4, ElbName)
                # DNS Name
                DnsName = ""
                if "DNSName" in response:
                    DnsName = response["DNSName"]
                if len(DnsName) != 0:
                    self.add_cell(self.cell_start, 5, DnsName)
                else:
                    self.add_cell(self.cell_start, 5, "-")
                # Type
                self.add_cell(self.cell_start, 6, str(response["Type"]).capitalize())
                # Port Configuration
                # Instance IDs
                target_group = self.client_elb2.describe_target_groups(LoadBalancerArn=response["LoadBalancerArn"])
                portstr = ""
                targetstr = ""
                try:
                    for idx, port in enumerate(target_group["TargetGroups"]):
                        if idx != 0:
                            portstr += ", \n"
                            targetstr += ", \n"
                        portstr += str(port.get("Port")) + " " + port.get("Protocol")
                        targetstr += port.get("TargetGroupName")
                    self.add_cell(self.cell_start, 7, portstr)
                    self.add_cell(self.cell_start, 8, targetstr)
                except:
                    self.add_cell(self.cell_start, 7, "-")
                    self.add_cell(self.cell_start, 8, "-")
                # Availability Zones
                zonestr = ""
                for idx, zone in enumerate(response["AvailabilityZones"]):
                    if idx != 0:
                        zonestr += ", \n"
                    zonestr += zone.get("ZoneName")
                self.add_cell(self.cell_start, 9, zonestr)
                # Security Groups
                secstr = ""
                try:
                    for idx, sec in enumerate(response["SecurityGroups"]):
                        if idx != 0:
                            secstr += ", \n"
                        secstr += sec
                        self.add_cell(self.cell_start, 10, secstr)
                except:
                    self.add_cell(self.cell_start, 10, "-")
                # Attributes
                Attributes = self.client_elb2.describe_load_balancer_attributes(LoadBalancerArn=response["LoadBalancerArn"])["Attributes"]
                # GWLB : CorssZone
                if response["Type"] == "gateway":
                    accessLogsEnabled = "-"
                    crossZoneLBEnabled = Attributes[1]["Value"]
                    idleTimeout = "-"
                # NLB : AccessLogs, CrossZone
                elif response["Type"] == "network":
                    accessLogsEnabled = Attributes[0]["Value"]
                    crossZoneLBEnabled = Attributes[4]["Value"]
                    idleTimeout = "-"
                # ALB : AccessLogs, CrossZone, Idle Timeout
                else:
                    accessLogsEnabled = Attributes[0]["Value"]
                    crossZoneLBEnabled = "true"  # 기본으로 활성화되어 있음
                    idleTimeout = Attributes[3]["Value"]
                # Cross Zone
                self.add_cell(self.cell_start, 11, str(crossZoneLBEnabled).capitalize())
                # IdleTimeout
                self.add_cell(self.cell_start, 12, idleTimeout)
                # Access Logs
                try:
                    self.add_cell(self.cell_start, 13, str(accessLogsEnabled).lower().capitalize())
                except:
                    self.add_cell(self.cell_start, 13, "-")
                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")