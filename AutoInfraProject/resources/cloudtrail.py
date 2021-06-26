from settings import Common


class CT(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="cloudtrail", region_name=self.region)
            if len(self.client.describe_trails()["trailList"]) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 25, 14, 16, 13, 15, 45, 8,8, 8, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Trail name", "Home Region", "Multi-region Trail", "Trail Insights",
                            "Organization Trail", "S3 bucket"]
            self.make_cell_header(self.cell_start, cell_headers)
            # CT
            for idx, trail in enumerate(self.client.describe_trails()["trailList"]):
                self.add_cell(self.cell_start, 2, idx + 1)
                self.add_cell(self.cell_start, 3, trail.get("Name"))
                self.add_cell(self.cell_start, 4, trail.get("HomeRegion"))
                self.add_cell(self.cell_start, 5, str(trail.get("IsMultiRegionTrail")).lower().capitalize())
                self.add_cell(self.cell_start, 6, str(trail.get("HasInsightSelectors")).lower().capitalize())
                self.add_cell(self.cell_start, 7, str(trail.get("IsOrganizationTrail")).lower().capitalize())
                self.add_cell(self.cell_start, 8, trail.get("S3BucketName"))
                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")