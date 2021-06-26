from settings import Common


class ElastiCache(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="elasticache", region_name=self.region)
            if len(self.client.describe_cache_clusters()["CacheClusters"]) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 25, 15, 12, 13, 12, 12, 21, 28, 13, 8, 8.10, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Cluster Id", "Node Type", "Engine", "Engine Version", "Status", "Nodes (EA)", "Security Groups", "Subnet Group", "Creation Date"]
            self.make_cell_header(self.cell_start, cell_headers)

            # ElastiCache
            for idx, node in enumerate(self.client.describe_cache_clusters()["CacheClusters"]):
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Cluster Id
                self.add_cell(self.cell_start, 3, node.get("CacheClusterId"))
                # Node Type
                self.add_cell(self.cell_start, 4, node.get("CacheNodeType"))
                # Engine
                self.add_cell(self.cell_start, 5, str(node.get("Engine")).capitalize())
                # Engine Version
                self.add_cell(self.cell_start, 6, node.get("EngineVersion"))
                # Status
                self.add_cell(self.cell_start, 7, node.get("CacheClusterStatus"))
                # Number of Nodes
                self.add_cell(self.cell_start, 8, node.get("NumCacheNodes"))
                # Security Groups
                node_sgs = ""
                for idx, sg in enumerate(node.get("SecurityGroups")):
                    if idx != 0:
                        node_sgs += ", \n"
                    node_sgs += sg.get("SecurityGroupId")
                self.add_cell(self.cell_start, 9, node_sgs)
                # Subnet Group
                self.add_cell(self.cell_start, 10, node.get('CacheSubnetGroupName'))
                # Creation Date
                self.add_cell(self.cell_start, 11, str(node.get('CacheClusterCreateTime')).split(" ")[0])

                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")