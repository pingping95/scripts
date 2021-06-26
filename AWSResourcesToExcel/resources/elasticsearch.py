from settings import Common


class ElasticSearch(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="es", region_name=self.region)
            if len(self.client.list_domain_names().get('DomainNames')) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 5, 25, 10, 13, 16, 20, 24, 20, 12, 20, 12, 10, 12, 16, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Domain Name", "ES Version", "Creation Date", "Availability Zones", "VPC ID",
                            "Subnet IDs", "Security Group IDs", "Cluster State", "Instance Type", "Instances (EA)",
                            "EBS Type", "EBS Size (GB)", "Auto Tune Option"]
            self.make_cell_header(self.cell_start, cell_headers)

            # Elastic Search
            for idx, domain_name in enumerate(self.client.list_domain_names().get('DomainNames')):
                es_domain_name = domain_name.get('DomainName')
                es_config = self.client.describe_elasticsearch_domain_config(
                    DomainName=es_domain_name
                ).get('DomainConfig')
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Domain Name
                self.add_cell(self.cell_start, 3, es_domain_name)
                # ES Version
                self.add_cell(self.cell_start, 4, es_config.get('ElasticsearchVersion').get('Options'))
                # Creation Date
                self.add_cell(self.cell_start, 5, str(es_config.get('ElasticsearchClusterConfig').
                                                      get('Status').get('CreationDate')).split(" ")[0])
                # Availability Zones
                es_azs = ""
                azs = es_config.get('VPCOptions').get('Options').get('AvailabilityZones')
                for i, az in enumerate(azs):
                    if i != 0:
                        es_azs += ", \n"
                    es_azs += az
                self.add_cell(self.cell_start, 6, es_azs)
                # VPC
                self.add_cell(self.cell_start, 7, es_config.get('VPCOptions').get('Options').get('VPCId'))
                # Subnet IDs
                es_subnet_ids = ""
                subnets = es_config.get('VPCOptions').get('Options').get('SubnetIds')
                for i, subnet in enumerate(subnets):
                    if i != 0:
                        es_subnet_ids += ", \n"
                    es_subnet_ids += subnet
                self.add_cell(self.cell_start, 8, es_subnet_ids)
                # Security Group IDs
                es_sg_ids = ""
                sgs = es_config.get('VPCOptions').get('Options').get('SecurityGroupIds')
                for i, sg in enumerate(sgs):
                    if i != 0:
                        es_sg_ids += ", \n"
                    es_sg_ids += sg
                self.add_cell(self.cell_start, 9, es_sg_ids)
                # ES Cluster State
                self.add_cell(self.cell_start, 10, es_config.get('ElasticsearchClusterConfig').get('Status').get('State'))

                # Instance Type
                self.add_cell(self.cell_start, 11, es_config.get('ElasticsearchClusterConfig').
                              get('Options').get('InstanceType'))
                # Instance Count
                self.add_cell(self.cell_start, 12, es_config.get('ElasticsearchClusterConfig').
                              get('Options').get('InstanceCount'))
                # EBS (Volume Type, Size)
                is_es_ebs_enabled = es_config.get('EBSOptions').get('Options').get('EBSEnabled')
                if is_es_ebs_enabled:
                    # EBS Volume Type
                    es_ebs_type = es_config.get('EBSOptions').get('Options').get('VolumeType')
                    # EBS Volume Size
                    es_ebs_size = es_config.get('EBSOptions').get('Options').get('VolumeSize')
                else:
                    es_ebs_type = "-"
                    es_ebs_size = "-"
                self.add_cell(self.cell_start, 13, es_ebs_type)
                self.add_cell(self.cell_start, 14, es_ebs_size)
                # Auto Tune Options
                es_auto_tune = str(
                    es_config.get("AutoTuneOptions").get("Options").get("DesiredState")).lower().capitalize()
                self.add_cell(self.cell_start, 15, es_auto_tune)
                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")