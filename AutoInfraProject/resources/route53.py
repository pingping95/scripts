from settings import Common


class Route53(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="route53", region_name=self.region)
            if self.client.list_hosted_zones_by_name()["HostedZones"]:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 33, 9, 65, 11, 75, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Domain Name", "Type", "Record Name", "Record Type", "Record Value"]
            self.make_cell_header(self.cell_start, cell_headers)
            # CT
            for idx, zone in enumerate(self.client.list_hosted_zones_by_name()["HostedZones"]):
                zoneId = zone["Id"]
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Domain Name
                zoneName = zone["Name"]
                real_zoneName = zoneName[:-1]
                self.add_cell(self.cell_start, 3, real_zoneName)
                # Type
                if zone["Config"]["PrivateZone"] == False:
                    self.add_cell(self.cell_start, 4, "Public")
                else:
                    self.add_cell(self.cell_start, 4, "Private")
                # Record Sets
                init_row_cnt = self.cell_start
                row_cnt = 0
                for record_set in self.client.list_resource_record_sets(HostedZoneId=zoneId)["ResourceRecordSets"]:
                    row_cnt += 1
                    # Record Name
                    record_name = record_set.get("Name")
                    # Record Type
                    record_type = record_set.get("Type")
                    # Record Values
                    # Alias : AliasTarget
                    # Value : ResourceRecords
                    record_value = ""
                    if "ResourceRecords" in record_set:
                        for idx, value in enumerate(record_set.get("ResourceRecords")):
                            if idx != 0:
                                record_value = record_value + ", \n"
                            record_value += value.get("Value")
                    elif "AliasTarget" in record_set:
                        record_value = record_set["AliasTarget"]["DNSName"]
                    else:
                        record_value = "-"
                    self.add_cell(self.cell_start, 5, record_name)
                    self.add_cell(self.cell_start, 6, record_type)
                    self.add_cell(self.cell_start, 7, record_value)
                    self.cell_start += 1

                self.sheet.merge_cells(start_row=init_row_cnt, end_row=init_row_cnt + row_cnt - 1, start_column=2, end_column=2)
                self.sheet.merge_cells(start_row=init_row_cnt, end_row=init_row_cnt + row_cnt - 1, start_column=3, end_column=3)
                self.sheet.merge_cells(start_row=init_row_cnt, end_row=init_row_cnt + row_cnt - 1, start_column=4, end_column=4)
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")