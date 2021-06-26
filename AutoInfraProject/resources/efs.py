from settings import Common


class Efs(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="efs", region_name=self.region)
            if len(self.client.describe_file_systems().get("FileSystems")) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 13, 25, 13, 18, 14, 12, 21 , 17, 9, 17, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers =["No.", "File System Id", "Name", "Creation Date", "Mount Targets (EA)", "Size (Standard)",
                           "Size (IA)", "LifeCycle Configuration", "Performance Mode", "Encrypted", "Throughput Mode"]
            self.make_cell_header(self.cell_start, cell_headers)

            # EFS
            for idx, efs in enumerate(self.client.describe_file_systems().get("FileSystems")):
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # File System Id
                file_sys_id = efs.get("FileSystemId")
                self.add_cell(self.cell_start, 3, efs.get("FileSystemId"))
                # Name
                self.add_cell(self.cell_start, 4, efs.get("Name"))
                # Creation Date
                self.add_cell(self.cell_start, 5, str(efs.get('CreationTime')).split(" ")[0])
                # Number of Mount Targets
                self.add_cell(self.cell_start, 6, efs.get("NumberOfMountTargets"))
                # Size in Standard
                efs_standard_size = efs["SizeInBytes"].get("ValueInStandard")
                readable_standard_size = self.return_humanbytes(efs_standard_size)
                self.add_cell(self.cell_start, 7, readable_standard_size)
                # Size in IA
                efs_ia_size = efs["SizeInBytes"].get("ValueInIA")
                readable_ia_size = self.return_humanbytes(efs_ia_size)
                self.add_cell(self.cell_start, 8, readable_ia_size)
                # LifeCycle Configuration
                efs_lifecycle = self.client.describe_lifecycle_configuration(
                    FileSystemId=file_sys_id
                ).get('LifecyclePolicies')
                if len(efs_lifecycle) != 0:
                    efs_transition_to_IA = efs_lifecycle[0].get('TransitionToIA')
                else:
                    efs_transition_to_IA = "None"
                self.add_cell(self.cell_start, 9, efs_transition_to_IA)
                # Performance Mode
                self.add_cell(self.cell_start, 10, str(efs.get("PerformanceMode")).capitalize())
                # Enctypted
                self.add_cell(self.cell_start, 11, str(efs.get("Encrypted")).lower().capitalize())
                # Throughput Mode
                self.add_cell(self.cell_start, 12, str(efs.get("ThroughputMode")).capitalize())

                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")