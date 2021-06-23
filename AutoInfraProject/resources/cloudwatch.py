from settings import Common


class CW(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="cloudwatch", region_name=self.region)
            self.run()

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 56, 56, 15, 23, 28, 10.75, 19, 8, 8, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, "CloudWatch Dashboards")
            # Cell header
            cell_headers = ["No.", "Dashboard Name", "Last Updated(UTC)", "Size"]
            self.make_cell_header(self.cell_start, cell_headers)
            # 1. CW Dashboards
            for idx, dashboard in enumerate(self.client.list_dashboards().get('DashboardEntries')):
                self.add_cell(self.cell_start, 2, idx + 1)
                self.add_cell(self.cell_start, 3, dashboard.get("DashboardName"))
                date = str(dashboard.get("LastModified"))
                self.add_cell(self.cell_start, 4, date.split("+")[0])
                self.add_cell(self.cell_start, 5, dashboard.get("Size"))
                self.cell_start += 1

            # 2. CW Metric Alarms
            # Header
            self.cell_start += 1
            self.make_header(self.cell_start, "CloudWatch Metric Alarms")
            # Cell header
            cell_headers = ["No.", "Alarm Name", "AlarmDescription", "Namespace", "MetricName", "ComparisonOperator", "Threshold", "StateValue"]
            self.make_cell_header(self.cell_start, cell_headers)

            for idx, alarm in enumerate(self.client.describe_alarms()["MetricAlarms"]):
                self.add_cell(self.cell_start, 2, idx + 1)
                self.add_cell(self.cell_start, 3, alarm.get("AlarmName"))
                self.add_cell(self.cell_start, 4, alarm.get("AlarmDescription"))
                self.add_cell(self.cell_start, 5, alarm.get("Namespace"))
                self.add_cell(self.cell_start, 6, alarm.get("MetricName"))
                self.add_cell(self.cell_start, 7, alarm.get("ComparisonOperator"))
                self.add_cell(self.cell_start, 8, alarm.get("Threshold"))
                self.add_cell(self.cell_start, 9, alarm.get("StateValue"))
                self.cell_start += 1

        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}")