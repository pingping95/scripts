from settings import Common


class Lambda(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="lambda", region_name=self.region)
            if self.client.list_functions().get("Functions"):
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 30, 12, 55, 11, 12, 9, 12, 8, 8, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Function Name", "Runtime", "Role", "Code Size", "Memory Size", "Time Out", "Package Type"]
            self.make_cell_header(self.cell_start, cell_headers)
            # CT
            for idx, each_lambda in enumerate(self.client.list_functions().get("Functions")):
                self.add_cell(self.cell_start, 2, idx + 1)
                # Function Name
                self.add_cell(self.cell_start, 3, each_lambda.get("FunctionName"))
                # Runtime
                self.add_cell(self.cell_start, 4, each_lambda.get("Runtime"))
                # Role Name
                role_name = each_lambda.get("Role").split(":")[-1]
                self.add_cell(self.cell_start, 5, role_name)
                # Code Size
                readable_code_size = self.return_humanbytes(each_lambda.get("CodeSize"))
                self.add_cell(self.cell_start, 6, readable_code_size)
                # Memory Size
                self.add_cell(self.cell_start, 7, each_lambda.get("MemorySize"))
                # Timeout
                self.add_cell(self.cell_start, 8, each_lambda.get("Timeout"))
                # Package Type
                self.add_cell(self.cell_start, 9, each_lambda.get("PackageType"))
                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")