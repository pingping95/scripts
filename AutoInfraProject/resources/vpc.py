from .settings import Common
import openpyxl
import boto3

class Vpc(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run = False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.resource = ses.resource(service_name="ec2")
            self.run()

    def run(self):
        # try:
        print(f"name: {self.name}, profile: {self.profile}, res: {self.resource}, reg : {self.region}")
        # Initialize
        self.sheet = self.wb.active
        self.sheet.title = self.name
        # Cell width
        cell_widths = [5, 5, 18, 25, 25, 25, 17, 15, 20, 20, 5, 5, 18, 24, 24, 13, 13, 13, 13, 13, 7, 7]
        self.fit_cell_width(cell_widths)
        # Header
        self.make_header(self.cell_start, self.sheet.title)
        # Cell header
        cell_headers = ["No.", "VPC Name", "VPC ID", "VPC CIDR Block"]
        self.make_cell_header(self.cell_start, cell_headers)
        # For loop
        for idx, vpc in enumerate(self.resource.vpcs.all()):
            # No.
            self.add_cell(self.cell_start, 2, idx + 1)
            # Name
            try:
                for tag in vpc.tags:
                    if tag["Key"] == "Name":
                        self.add_cell(self.cell_start, 3, tag['Value'])
            except Exception as e:
                self.add_cell(self.cell_start, 3, "-")
            # Id
            self.add_cell(self.cell_start, 4, vpc.id)
            # CIDR
            self.add_cell(self.cell_start, 5, vpc.cidr_block)
            
            self.cell_start += 1
        # except Exception as e:
        #     self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}")