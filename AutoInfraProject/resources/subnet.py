from .settings import Common
import openpyxl
import boto3

class Subnet(Common):
    def __init__(self, name, workbook, ses, info, log, is_run = False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = info.get('porfile')
            self.region = info.get('region')
            self.resource = ses.resource(service_name="ec2", region_name=self.region)
            self.client = ses.client(service_name="ec2", region_name=self.region)
            self.run()

    def run(self):
        try:
            print(f"name: {self.name}, profile: {self.profile}, res: {self.client}, reg : {self.region}")
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            
            # Cell width
            cell_widths = [5, 5, 18, 25, 25, 25, 17, 15, 20, 20, 5, 5, 18, 24, 24, 13, 13, 13, 13, 13, 7, 7]
            self.fit_cell_width(cell_widths)
            
            # Header
            self.make_header(self.cell_start, self.name)
            
            # Cell header
            cell_headers = ["No.", "Subnet Type", "VPC ID", "Subnet Name", "Subnet ID", "Subnet CIDR Block", "Availability Zone",
                            "Network ACLs", "Route Tables"]
            self.make_cell_header(self.cell_start, cell_headers)
            
            # Public Subnet을 pubSubnetDict에 넣음
            pubSubnetDict = {}
            
            # For loop
            for rt in self.client.describe_route_tables().get('RouteTables'):
                associations = rt["Associations"]
                routes = rt["Routes"]
                isPublic = False
                
                for route in routes:
                    gid = route.get("GatewayId", "")
                    if gid.startswith("igw-"):
                        isPublic = True

                if not isPublic:
                    continue
                
                for assoc in associations:
                    subnetId = assoc.get("SubnetId", None)  # This checks for explicit associations, only
                    if subnetId:
                        pubSubnetDict[subnetId] = isPublic
            
            # For Loop 2

            for idx, subnet in enumerate(self.resource.subnets.all()):
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Subnet Type (Check if this subnet is Public or Private)
                if subnet.id in pubSubnetDict:
                    self.add_cell(self.cell_start, 3, "Public Network")
                else:
                    self.add_cell(self.cell_start, 3, "Private Network")
                # VPC Id
                self.add_cell(self.cell_start, 4, subnet.vpc_id)

                # Subnet Name
                try:
                    for tags in subnet.tags:
                        if tags["Key"] == "Name":
                            add_cell(sheet1, sheet1_cell_start, 5, tags["Value"])
                except:
                    self.add_cell(self.cell_start, 5, "-")
                # Subnet Id
                self.add_cell(self.cell_start, 6, subnet.subnet_id)
                # Subnet CIDR Block
                self.add_cell(self.cell_start, 7, subnet.cidr_block)
                # Subnet Availability Zone
                self.add_cell(self.cell_start, 8, subnet.availability_zone)
                # NetworkAcls
                for acls in self.client.describe_network_acls()["NetworkAcls"]:
                    for i in acls["Associations"]:
                        if i.get("SubnetId") == subnet.subnet_id:
                            self.add_cell(self.cell_start, 9, i.get("NetworkAclId"))

                # Route Tables
                try:
                    route_table = self.client.describe_route_tables(
                        Filters=[
                            {
                                "Name": "association.subnet-id",
                                "Values": [
                                    subnet.subnet_id,
                                ],
                            },
                        ],
                    )["RouteTables"]
                    self.add_cell(self.cell_start, 10, route_table[0].get("Associations")[0].get("RouteTableId"))
                except:
                    self.add_cell(self.cell_start, 10, "-")
                
                self.cell_start += 1

                
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}")