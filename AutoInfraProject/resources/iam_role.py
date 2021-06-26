from settings import Common


class IamRole(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.resource = ses.resource(service_name="iam", region_name=self.region)
            self.client = ses.client(service_name="iam", region_name=self.region)

            iam_roles = self.client.list_roles()["Roles"]
            iam_roles_cnt = 0
            for i in iam_roles:
                if i["Path"] == "/":
                    iam_roles_cnt += 1

            if iam_roles_cnt != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 60, 20, 70, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Role Name", "Creation Date", "Description"]
            self.make_cell_header(self.cell_start, cell_headers)

            # IAM Role
            count = 1
            for iam_role in self.client.list_roles()["Roles"]:
                # 직접 생성한 IAM Roles만 엑셀에 추가함
                if iam_role["Path"] == "/":
                    # No.
                    self.add_cell(self.cell_start, 2, count)

                    # Role Name
                    roleName = iam_role["RoleName"]
                    self.add_cell(self.cell_start, 3, roleName)

                    # Creation Date
                    role_creation_date = str(iam_role.get('CreateDate')).split(" ")[0]
                    self.add_cell(self.cell_start, 4, role_creation_date)

                    # Description
                    try:
                        iamRoleDesc = iam_role["Description"]
                        self.add_cell(self.cell_start, 5, iamRoleDesc)
                    except Exception:
                        self.add_cell(self.cell_start, 5, "-")
                    self.cell_start += 1
                    count += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")