from settings import Common


class S3(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.session = ses
            self.client = ses.client(service_name="s3", region_name=self.region)
            self.resource = ses.resource(service_name="s3", region_name=self.region)

            if len(self.client.list_buckets().get('Buckets')) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")
    def run(self):
        import datetime
        from dateutil.relativedelta import relativedelta
        right_now = datetime.datetime.now()
        now = right_now.date()
        target_date = str(now + relativedelta(days=-2))

        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [5, 5, 55, 13, 14, 22, 11, 11, 11, 8, 8, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Bucket Name", "Creation Date", "Region", "Size\n (" + str(target_date) + " 측정값)",
                            "Logging", "Versioning",]
            self.make_cell_header(self.cell_start, cell_headers)
            # S3
            for idx, bucket in enumerate(self.resource.buckets.all()):
                try:
                    # No.
                    self.add_cell(self.cell_start, 2, idx + 1)
                    # Name
                    bucket_name = bucket.name
                    self.add_cell(self.cell_start, 3, bucket_name)
                    # Creation Date
                    bucket_creation_date = str(bucket.creation_date).split(" ")[0]
                    self.add_cell(self.cell_start, 4, bucket_creation_date)
                    # Bucket Region
                    bucket_region = self.client.get_bucket_location(
                        Bucket=bucket_name
                    ).get('LocationConstraint')
                    if bucket_region == None:
                        bucket_region = "us-east-1"
                    self.add_cell(self.cell_start, 5, bucket_region)
                    # Bucket Size
                    cloudwatch_temp_cli = self.session.client(service_name="cloudwatch", region_name=bucket_region)
                    readable_bucket_size = 0
                    response = cloudwatch_temp_cli.get_metric_statistics(
                        Namespace="AWS/S3",
                        MetricName="BucketSizeBytes",
                        Dimensions=[
                            {"Name": "BucketName", "Value": bucket_name},
                            {"Name": "StorageType", "Value": "StandardStorage"},
                        ],
                        Statistics=["Average"],
                        Period=3600,
                        StartTime=(now - datetime.timedelta(days=2)).isoformat(),
                        EndTime=now.isoformat(),
                    )
                    if len(response["Datapoints"]) != 0:
                        bucket_size = int(response["Datapoints"][0].get("Average"))
                    else:
                        bucket_size = "-"
                    if type(bucket_size) == int:
                        readable_bucket_size = self.return_humanbytes(bucket_size)
                    else:
                        readable_bucket_size = "-"
                    self.add_cell(self.cell_start, 6, readable_bucket_size)
                    try:
                        # Logging
                        bucket_logging = self.client.get_bucket_logging(Bucket=bucket_name)
                        if "LoggingEnabled" in bucket_logging:
                            self.add_cell(self.cell_start, 7, "Enabled")
                        else:
                            self.add_cell(self.cell_start, 7, "Disabled")
                        # Versioning
                        bucket_versioning = self.client.get_bucket_versioning(Bucket=bucket_name)
                        if "Status" in bucket_versioning:
                            self.add_cell(self.cell_start, 8, "Enabled")
                        else:
                            self.add_cell(self.cell_start, 8, "Disabled")
                    except Exception as e:
                        self.add_cell(self.cell_start, 7, "-")
                        self.add_cell(self.cell_start, 8, "-")
                        self.log.write(f"{bucket_name} : 비 정상적인 Bucket으로 추정", e)
                except Exception as e:
                    self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")

                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")