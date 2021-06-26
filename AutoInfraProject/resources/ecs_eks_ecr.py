from settings import Common


class EcsEks(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="ecs", region_name=self.region)
            self.client2 = ses.client(service_name="eks", region_name=self.region)
            self.client3 = ses.client(service_name="ecr", region_name=self.region)
            if len(self.client.list_clusters().get('clusterArns')) != 0 or len(self.client2.list_clusters()["clusters"]) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        # Initialize
        self.sheet = self.wb.create_sheet(self.name)
        self.sheet.title = self.name
        # Cell width
        cell_widths = [6, 6, 30, 14, 80, 15, 47, 23, 19, 25, 20, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
        self.fit_cell_width(cell_widths)

        # ECS
        if len(self.client.list_clusters().get('clusterArns')) != 0:
            try:
                # Header
                self.make_header(self.cell_start, "ECS")
                # Cell header
                cell_headers = ["No.", "Cluster Name", "Status", "Registered Container Instances", "Running Tasks",
                                "Active Services", "Container Insight"]
                self.make_cell_header(self.cell_start, cell_headers)
                cluster_list = []
                for clus in self.client.list_clusters().get('clusterArns'):
                    name = str(clus).split('/')[-1]
                    cluster_list.append(name)

                for idx, cluster in enumerate(self.client.describe_clusters(clusters=cluster_list).get('clusters')):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Cluster Name
                    self.add_cell(self.cell_start, 3, cluster.get('clusterName'))
                    # Status
                    self.add_cell(self.cell_start, 4, str(cluster.get('status')).lower().capitalize())
                    # Registered Container Instances
                    self.add_cell(self.cell_start, 5, cluster.get('registeredContainerInstancesCount'))
                    # Running Tasks
                    self.add_cell(self.cell_start, 6, cluster.get('runningTasksCount'))
                    # Active Services
                    self.add_cell(self.cell_start, 7, cluster.get('activeServicesCount'))
                    # Container Insight
                    for setting in cluster.get('settings'):
                        if setting.get('name') == 'containerInsights':
                            cont_insight = str(setting.get('value')).capitalize()
                            break
                    self.add_cell(self.cell_start, 8, cont_insight)
                    self.cell_start += 1
            except Exception as e:
                self.log.write(f"Error 발생, 리소스: ECS, 내용: {e}\n")
        else:
            self.log.write(f"There is no ECS\n")

        # EKS
        if len(self.client2.list_clusters()["clusters"]) != 0:
            try:
                self.cell_start += 1
                # Header
                self.make_header(self.cell_start, "EKS")
                # Cell header
                cell_headers = ["No.", "Cluster Name", "Cluster Version", "Platform Version", "Status", "Subnet IDs",
                        "Cluster Security Group IDs", "Security Group IDs", "Node Group Names", "Fargate Profile Names"]
                self.make_cell_header(self.cell_start, cell_headers)

                for idx, cluster_name in enumerate(self.client2.list_clusters()["clusters"]):
                    cluster = self.client2.describe_cluster(name=cluster_name)["cluster"]
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Cluster name
                    name = cluster.get("name")
                    self.add_cell(self.cell_start, 3, name)
                    # Version
                    self.add_cell(self.cell_start, 4, cluster.get("version"))
                    # Platform Version
                    self.add_cell(self.cell_start, 5, cluster.get("platformVersion"))
                    # Status
                    self.add_cell(self.cell_start, 6, str(cluster.get("status")).lower().capitalize())
                    # Subnet IDs
                    subnetIds = ""
                    for i, subnetId in enumerate(cluster["resourcesVpcConfig"]["subnetIds"]):
                        if i != 0:
                            subnetIds += ", "
                        elif i != 0 and i % 2 == 0:
                            subnetIds += "\n"
                        subnetIds += subnetId
                    self.add_cell(self.cell_start, 7, subnetIds)
                    # Cluster Security Group IDs
                    cluster_sg_ids = cluster["resourcesVpcConfig"].get("clusterSecurityGroupId")
                    if cluster_sg_ids:
                        self.add_cell(self.cell_start, 8, cluster_sg_ids)
                    else:
                        self.add_cell(self.cell_start, 8, "-")
                    # Security Group IDs
                    sg_ids_list = cluster["resourcesVpcConfig"].get("securityGroupIds")
                    sg_ids = ""
                    for i, sg_id in enumerate(sg_ids_list):
                        if i != 0:
                            sg_ids += ", \n"
                        sg_ids += sg_id
                    self.add_cell(self.cell_start, 9, sg_ids)
                    # Node Group Names
                    node_groups = self.client2.list_nodegroups(clusterName=name)["nodegroups"]
                    if len(node_groups) != 0:
                        node_group_names = ""
                        for i, node_group in enumerate(node_groups):
                            if i != 0:
                                node_group_names += ", \n"
                            node_group_names += node_group
                        self.add_cell(self.cell_start, 10, node_group_names)
                    else:
                        self.add_cell(self.cell_start, 10, "-")
                    # Fargate Profile Names
                    try:
                        fargates = self.client2.list_fargate_profiles(clusterName=name)["fargateProfileNames"]
                    except:
                        fargates = []
                    if len(fargates) != 0:
                        fg_names = ""
                        for i, fargate in enumerate(fargates):
                            if i != 0:
                                fg_names += ", \n"
                            fg_names += fargate
                        self.add_cell(self.cell_start, 11, fg_names)
                    else:
                        self.add_cell(self.cell_start, 11, "-")
                    self.cell_start += 1
            except Exception as e:
                self.log.write(f"Error 발생, 리소스: ECS, 내용: {e}\n")
        else:
            self.log.write(f"There is no EKS\n")


        # ECR
        if len(self.client3.describe_repositories().get('repositories')) != 0:
            try:
                self.cell_start += 1
                # Header
                self.make_header(self.cell_start, "ECR")
                # Cell header
                cell_headers = ["No.", "Repo Name", "Creation Date", "Repo URI", "Scan on Push", "Encryption config"]
                self.make_cell_header(self.cell_start, cell_headers)

                for idx, ecr in enumerate(self.client3.describe_repositories().get('repositories')):
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Repository Name
                    self.add_cell(self.cell_start, 3, ecr.get('repositoryName'))
                    # Creation Date
                    self.add_cell(self.cell_start, 4, str(ecr.get('createdAt')).split(' ')[0])
                    # Repository URI
                    self.add_cell(self.cell_start, 5, ecr.get('repositoryUri'))
                    # Scan on push
                    self.add_cell(self.cell_start, 6, str(ecr.get('imageScanningConfiguration').get('scanOnPush')).lower().capitalize())
                    # Encryption Configuration
                    self.add_cell(self.cell_start, 7, ecr.get('encryptionConfiguration').get('encryptionType'))
                    self.cell_start += 1
            except Exception as e:
                self.log.write(f"Error 발생, 리소스: ECR, 내용: {e}\n")
        else:
            self.log.write(f"There is no ECR\n")

