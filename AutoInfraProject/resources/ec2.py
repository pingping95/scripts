from settings import Common


class Ec2(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="ec2", region_name=self.region)

            if len(self.client.describe_instances().get("Reservations")) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Elastic IP 리스트 뽑아냄
            eip_list = []
            try:
                for eip in self.client.describe_addresses().get("Addresses"):
                    eip_list.append(eip.get("PublicIp"))
            except Exception:
                pass
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [5, 5, 14, 35, 8, 8, 12, 11, 20, 23, 19, 13, 13, 13, 21, 15, 20, 38, 15, 21, 15, 10]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Availability Zone", "Instance Name", "Status", "AMI (OS)", "Instance Type", "Subnet Type",
                            "VPC ID", "Subnet ID", "Instance ID", "Private IP", "Public IP", "Elstic IP", "Root Volume ID",
                            "Root Volume (GB)", "Key Pair", "Security Groups", "IAM role", "Data Volume ID", "Data Volume (GB)"]
            self.make_cell_header(self.cell_start, cell_headers)
            # EC2
            indx = 0
            for reservation in self.client.describe_instances().get("Reservations"):
                for Instances in reservation["Instances"]:
                    indx += 1
                    # No.
                    self.add_cell(self.cell_start, 2, indx)
                    # Availability Zone
                    Zone = Instances["Placement"].get("AvailabilityZone")
                    self.add_cell(self.cell_start, 3, Zone)
                    # Instance Name
                    InstanceName = "-"
                    try:
                        for i in Instances["Tags"]:
                            if i.get("Key") == "Name":
                                InstanceName = i.get("Value")
                                break
                    except:
                        pass
                    self.add_cell(self.cell_start, 4, InstanceName)
                    # Status
                    # 'pending'|'running'|'shutting-down'|'terminated'|'stopping'|'stopped'
                    Status = Instances["State"].get("Name")
                    self.add_cell(self.cell_start, 5, Instances["State"].get("Name"))
                    if Status == "running" or Status == "stopped":
                        # AMI (OS)
                        try:
                            OS = Instances["Platform"]
                            self.add_cell(self.cell_start, 6, OS)
                        except:
                            self.add_cell(self.cell_start, 6, "Linux")
                        # Instance Type
                        self.add_cell(self.cell_start, 7, Instances["InstanceType"])
                        # Subnet Type
                        if Instances.get("PublicIpAddress"):
                            self.add_cell(self.cell_start, 8, "Public")
                        else:
                            self.add_cell(self.cell_start, 8, "Private")
                        # VPC ID
                        self.add_cell(self.cell_start, 9, Instances["VpcId"])
                        # Subnet ID
                        self.add_cell(self.cell_start, 10, Instances["SubnetId"])
                        # Instance ID
                        self.add_cell(self.cell_start, 11, Instances["InstanceId"])

                        # Private IP
                        Private_IP = Instances["PrivateIpAddress"]
                        self.add_cell(self.cell_start, 12, Private_IP)

                        # Public IP
                        if Instances.get("PublicIpAddress"):
                            Public_IP = Instances["PublicIpAddress"]
                            self.add_cell(self.cell_start, 13, Public_IP)
                        else:
                            Public_IP = ""
                            self.add_cell(self.cell_start, 13, Public_IP)

                        # EIP
                        if Public_IP in eip_list:
                            self.add_cell(self.cell_start, 14, Public_IP)
                        else:
                            self.add_cell(self.cell_start, 14, "-")
                        # data volume 리스트
                        data_volume_list = []
                        # Root Volume ID
                        root_volume_id = ""
                        for ebs in Instances["BlockDeviceMappings"]:
                            data_volume_list.append(ebs["Ebs"].get("VolumeId"))  # EC2 인스턴스에 있는 모든 volume을 추가
                            if ebs.get("DeviceName") == Instances["RootDeviceName"]:
                                data_volume_list.remove(ebs["Ebs"].get("VolumeId"))  # Root Volume은 리스트에서 제외
                                root_volume_id = str(ebs["Ebs"].get("VolumeId"))
                                self.add_cell(self.cell_start, 15, root_volume_id)

                        # Root Volume (GB)
                        try:
                            response = self.client.describe_volumes(VolumeIds=[root_volume_id])
                            for volume in response["Volumes"]:
                                Size = volume.get("Size")
                                self.add_cell(self.cell_start, 16, Size)
                        # Root Volume이 없을 경우 오류 발생 가능성 있음
                        except Exception:
                            self.add_cell(self.cell_start, 16, "-")
                        # Key Pair
                        try:
                            KeyPair = Instances["KeyName"]
                            self.add_cell(self.cell_start, 17, KeyPair)
                        except:
                            pass
                        # Security Group
                        sec_groups = ""
                        for idx, sec in enumerate(Instances["SecurityGroups"]):
                            if idx != 0:
                                sec_groups += ", \n"
                            sec_groups += sec.get("GroupName")
                            # SG = sec.get('GroupName')
                        self.add_cell(self.cell_start, 18, sec_groups)
                        # IAM role
                        try:
                            IAM = Instances["IamInstanceProfile"].get("Arn").split("/")[-1]
                            self.add_cell(self.cell_start, 19, IAM)
                        except:
                            self.add_cell(self.cell_start, 19, "-")
                        # Data Volume이 존재할 경우
                        if len(data_volume_list) != 0:
                            # Data Volume ID
                            data_volume_id = ""
                            data_volume_size = ""
                            for idx, data_volume in enumerate(data_volume_list):
                                if idx != 0:
                                    data_volume_id += ", \n"
                                    data_volume_size += ", \n"
                                data_volume_id += data_volume
                                response2 = self.client.describe_volumes(VolumeIds=[data_volume])
                                for volume in response2["Volumes"]:
                                    data_volume_size += str(volume.get("Size"))
                            self.add_cell(self.cell_start, 20, data_volume_id)
                            self.add_cell(self.cell_start, 21, data_volume_size)
                        else:
                            self.add_cell(self.cell_start, 20, "-")
                            self.add_cell(self.cell_start, 21, "-")
                    else:
                        pass
                    self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")