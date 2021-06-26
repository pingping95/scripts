from settings import Common


class RDS(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="rds", region_name=self.region)

            if len(self.client.describe_db_instances().get('DBInstances')) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [5, 5, 14, 35, 11, 13, 22, 15, 15, 21, 30, 30, 23, 23, 12, 27,22, 20, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Availability Zone", "RDS Name", "RDS Engine", "Engine Version", "DB Instance Class", "Storage Type",
                            "Master Username", "Master Password", "VPC ID", "Subnet Group", "Parameter Group", "Option Group",
                            "Database Port","Preferred Maintenance Window", "Preferred Backup Window", "Backup Retention Time"]
            self.make_cell_header(self.cell_start, cell_headers)
            # RDS
            for idx, rdsdata in enumerate(self.client.describe_db_instances().get('DBInstances')):
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Availability Zone
                self.add_cell(self.cell_start, 3, rdsdata["AvailabilityZone"])
                # RDS Name
                self.add_cell(self.cell_start, 4, rdsdata["DBInstanceIdentifier"])
                # DB Engine
                self.add_cell(self.cell_start, 5, rdsdata["Engine"])
                # Engine Version
                self.add_cell(self.cell_start, 6, rdsdata["EngineVersion"])
                # DB Instance Class
                self.add_cell(self.cell_start, 7, rdsdata["DBInstanceClass"])
                # Storage Type
                self.add_cell(self.cell_start, 8, rdsdata["StorageType"])
                # Master Username
                self.add_cell(self.cell_start, 9, rdsdata["MasterUsername"])
                # Master Password
                self.add_cell(self.cell_start, 10, "-")
                # VPC ID
                self.add_cell(self.cell_start, 11, rdsdata["DBSubnetGroup"].get("VpcId"))
                # Subnet Group
                subnetstr = rdsdata["DBSubnetGroup"].get("DBSubnetGroupName")
                subnetstr += "\n( "
                for idx, subnet in enumerate(rdsdata["DBSubnetGroup"].get("Subnets")):
                    if idx != 0:
                        subnetstr += ", \n"
                    subnetstr += subnet.get("SubnetIdentifier")
                subnetstr += " )"
                self.add_cell(self.cell_start, 12, subnetstr)
                # Parameter Group
                for dbparam in rdsdata["DBParameterGroups"]:
                    self.add_cell(self.cell_start, 13, dbparam.get("DBParameterGroupName"))
                # Option Group
                for dboption in rdsdata["OptionGroupMemberships"]:
                    self.add_cell(self.cell_start, 14, dboption.get("OptionGroupName"))
                # Database Port
                self.add_cell(self.cell_start, 15, rdsdata["Endpoint"].get("Port"))
                # Maintenance Time
                self.add_cell(self.cell_start, 16, rdsdata["PreferredMaintenanceWindow"])
                # PreferredMaintenanceWindowPreferredBackupWindow Backup Retention Time
                self.add_cell(self.cell_start, 17, rdsdata["PreferredBackupWindow"])
                self.add_cell(self.cell_start, 18, rdsdata["BackupRetentionPeriod"])
                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")