from settings import Common


class IamUser(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.resource = ses.resource(service_name="iam", region_name=self.region)
            self.client = ses.client(service_name="iam", region_name=self.region)
            self.run()

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 30, 30, 38, 33, 13, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "User Name", "Group Names", "Policies (Attached to Groups)", "Policies (Attached to User)", "Creation Date"]
            self.make_cell_header(self.cell_start, cell_headers)

            # IAM User
            for idx, user_detail in enumerate(self.client.get_account_authorization_details(Filter=["User"])["UserDetailList"]):
                self.add_cell(self.cell_start, 2, idx + 1)
                self.add_cell(self.cell_start, 3, str(user_detail.get("UserName")))

                if len(user_detail.get("GroupList")) != 0:
                    # Group 리스트 생성
                    group_list = []
                    for idx, group in enumerate(user_detail.get("GroupList")):
                        group_list.append(group)
                    str_group_list = ", \n".join(group_list)
                    self.add_cell(self.cell_start, 4, str_group_list)
                    # Group Policies
                    group_policies = []
                    for each_group in group_list:
                        iam_res_group = self.resource.Group(each_group)
                        policy_generator = iam_res_group.attached_policies.all()
                        for policy in policy_generator:
                            group_policies.append(policy.policy_name)
                    if len(group_policies) != 0:
                        str_group_policies = ", \n".join(group_policies)
                        self.add_cell(self.cell_start, 5, str_group_policies)
                    else:
                        self.add_cell(self.cell_start, 5, "-")
                    # User Policies
                    user_policies = []
                    for policy in user_detail.get("AttachedManagedPolicies"):
                        user_policies.append(str(policy["PolicyName"]))
                    only_user_policies = [x for x in user_policies if x not in group_policies]
                    if len(only_user_policies) != 0:
                        str_only_user_policies = ", \n".join(only_user_policies)
                        self.add_cell(self.cell_start, 6, str_only_user_policies)
                    else:
                        self.add_cell(self.cell_start, 6, "-")
                else:
                    self.add_cell(self.cell_start, 4, "-")
                    self.add_cell(self.cell_start, 5, "-")
                    gpolicy = ""
                    for idx, attachpolicy in enumerate(user_detail.get("AttachedManagedPolicies")):
                        if idx != 0:
                            gpolicy += ", \n"
                        gpolicy += attachpolicy["PolicyName"]
                    self.add_cell(self.cell_start, 6, gpolicy)
                # Creation Date
                created_date = str(user_detail.get("CreateDate"))
                iamUser_date = created_date.split(" ")[0]
                self.add_cell(self.cell_start, 7, iamUser_date)
                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")