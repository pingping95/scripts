from settings import Common


class SecurityGroup(Common):
    def __init__(self, name, workbook, ses, p_name, r_name, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = p_name
            self.region = r_name
            self.resource = ses.resource(service_name="ec2", region_name=self.region)
            self.run()

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = f"{self.name}"
            # Cell width
            cell_widths = [5, 5, 55, 20, 22, 13, 24, 45, 7, 7.8, 55, 23, 22, 13, 20, 30, 10, 11, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name + " (Inbound)")
            # Cell header
            cell_headers = ["No.", "Security Groups Name", "Group ID", "Type", "Port Range", "source",
                            "비고(Description)"]
            self.make_cell_header(self.cell_start, cell_headers)
            # For loop
            for idx, security_group in enumerate(self.resource.security_groups.all()):
                short_start = self.cell_start
                long_start1 = self.cell_start
                short_start1 = self.cell_start
                long_start2 = self.cell_start
                sec_group = self.resource.SecurityGroup(security_group.id)
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Security Group Name
                self.add_cell(self.cell_start, 3, sec_group.group_name)
                # Security Group Id
                self.add_cell(self.cell_start, 4, sec_group.group_id)
                # Inbound
                # Type, Port Range, Source, Description
                for inbound in sec_group.ip_permissions:
                    ip_range = []
                    desc = []
                    # Type
                    if inbound.get("IpProtocol") == "-1":
                        ipType = "ALL Traffic"
                    else:
                        ipType = inbound.get("IpProtocol")
                    # portRange
                    if inbound.get("FromPort") == inbound.get("ToPort"):
                        portrange = inbound.get("FromPort")
                    else:
                        portrange = str(inbound.get("FromPort")) + " - " + str(inbound.get("ToPort"))
                    if portrange == "0--1" or portrange == -1:
                        portrange = "N/A"
                    if ipType == "ALL Traffic":
                        portrange = "All"
                    ipType = self.check_type(ipType, portrange)
                    # source
                    for ips in inbound.get("IpRanges"):
                        ip_range.append(ips.get("CidrIp"))
                        description = ips.get("Description")
                        if description is None:
                            desc.append("-")
                        else:
                            desc.append(description)
                    # 비고
                    for ips in inbound.get("Ipv6Ranges"):
                        ip_range.append(ips.get("CidrIpv6"))
                        description = ips.get("Description")
                        if description is None:
                            desc.append("-")
                        else:
                            desc.append(description)
                    for group in inbound.get("UserIdGroupPairs"):
                        ip_range.append(group.get("GroupId"))
                        description = group.get("Description")
                        if description is None:
                            desc.append("-")
                        else:
                            desc.append(description)
                    self.add_cell(short_start1, 5, ipType)
                    self.add_cell(short_start1, 6, portrange)
                    tmp1 = long_start1
                    for ip, desc in zip(ip_range, desc):
                        self.add_cell(long_start1, 7, ip)
                        self.add_cell(long_start1, 8, desc)
                        long_start1 += 1
                    tmp1_1 = long_start1
                    self.sheet.merge_cells(start_row=tmp1, end_row=tmp1_1 - 1, start_column=5, end_column=5)
                    self.sheet.merge_cells(start_row=tmp1, end_row=tmp1_1 - 1, start_column=6, end_column=6)
                    short_start1 = long_start1
                try:
                    if long_start1 >= long_start2:
                        self.cell_start = long_start1
                        # cell merge
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start1 - 1, start_column=2, end_column=2)
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start1 - 1, start_column=3, end_column=3)
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start1 - 1, start_column=4, end_column=4)
                    else:
                        self.cell_start = long_start2
                        # cell merge
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start2 - 1, start_column=2, end_column=2)
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start2 - 1, start_column=3, end_column=3)
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start2 - 1, start_column=4, end_column=4)
                except:
                    pass

            # 2. Outbound
            self.cell_start += 1
            # Header
            self.make_header(self.cell_start, self.name + " (Outbound)")
            # Cell header
            cell_headers = ["No.", "Security Groups Name", "Group ID", "Type", "Port Range", "source",
                            "비고(Description)"]
            self.make_cell_header(self.cell_start, cell_headers)
            # For loop
            for idx, security_group in enumerate(self.resource.security_groups.all()):
                short_start = self.cell_start
                long_start1 = self.cell_start
                short_start1 = self.cell_start
                long_start2 = self.cell_start
                sec_group = self.resource.SecurityGroup(security_group.id)
                # outbound
                for outbound in sec_group.ip_permissions_egress:
                    if outbound:
                        ip_range = []
                        desc = []
                        # Type
                        if outbound.get("IpProtocol") == "-1":
                            ipType = "ALL Traffic"
                        else:
                            ipType = outbound.get("IpProtocol")
                        # portRange
                        if outbound.get("FromPort") == outbound.get("ToPort"):
                            portrange = outbound.get("FromPort")
                        else:
                            portrange = str(outbound.get("FromPort")) + " - " + str(outbound.get("ToPort"))
                        if portrange == "0--1" or portrange == -1:
                            portrange = "N/A"
                        if ipType == "ALL Traffic":
                            portrange = "All"
                        ipType = self.check_type(ipType, portrange)
                        # source
                        for ips in outbound.get("IpRanges"):
                            ip_range.append(ips.get("CidrIp"))
                            description = ips.get("Description")
                            if description is None:
                                desc.append("-")
                            else:
                                desc.append(description)
                        # 비고
                        for ips in outbound.get("Ipv6Ranges"):
                            ip_range.append(ips.get("CidrIpv6"))
                            description = ips.get("Description")
                            if description is None:
                                desc.append("-")
                            else:
                                desc.append(description)
                        for group in outbound.get("UserIdGroupPairs"):
                            ip_range.append(group.get("GroupId"))
                            description = group.get("Description")
                            if description is None:
                                desc.append("-")
                            else:
                                desc.append(description)
                        # No.
                        self.add_cell(self.cell_start, 2, idx + 1)
                        # Security Group Name
                        self.add_cell(self.cell_start, 3, sec_group.group_name)
                        # Security Group Id
                        self.add_cell(self.cell_start, 4, sec_group.group_id)

                        self.add_cell(short_start1, 5, ipType)
                        self.add_cell(short_start1, 6, portrange)
                        tmp1 = long_start1
                        for ip, desc in zip(ip_range, desc):
                            self.add_cell(long_start1, 7, ip)
                            self.add_cell(long_start1, 8, desc)
                            # Security Group Name
                            self.add_cell(self.cell_start, 3, sec_group.group_name)
                            # Security Group Id
                            self.add_cell(self.cell_start, 4, sec_group.group_id)
                            long_start1 += 1

                        tmp1_1 = long_start1
                        try:
                            self.sheet.merge_cells(start_row=tmp1, end_row=tmp1_1 - 1, start_column=5, end_column=5)
                            self.sheet.merge_cells(start_row=tmp1, end_row=tmp1_1 - 1, start_column=6, end_column=6)
                        except:
                            pass
                        short_start1 = long_start1

                    self.cell_start = long_start1
                try:
                    if long_start1 >= long_start2:
                        self.cell_start = long_start1
                        # cell merge
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start1 - 1, start_column=2, end_column=2)
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start1 - 1, start_column=3, end_column=3)
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start1 - 1, start_column=4, end_column=4)
                    else:
                        self.cell_start = long_start2
                        # cell merge
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start2 - 1, start_column=2, end_column=2)
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start2 - 1, start_column=3, end_column=3)
                        self.sheet.merge_cells(start_row=short_start, end_row=long_start2 - 1, start_column=4, end_column=4)
                except:
                    pass
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")