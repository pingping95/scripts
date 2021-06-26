from settings import Common


class DynamoDb(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="dynamodb", region_name=self.region)
            if len(self.client.list_tables().get('TableNames')) != 0:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [5, 5, 35, 7, 13, 20, 20, 18, 18, 10, 10, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "Table Name", "Status", "Creation Date", "Partition Key", "Sort Key",
                            "Total Read Capacity", "Total Write Capacity", "Table Size", "Item Count"]
            self.make_cell_header(self.cell_start, cell_headers)

            # EFS
            for idx, table_name in enumerate(self.client.list_tables().get('TableNames')):
                table = self.client.describe_table(TableName=table_name).get('Table')
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Table name
                self.add_cell(self.cell_start, 3, table_name)
                # Status
                dynamo_status = str(table.get('TableStatus')).lower().capitalize()
                self.add_cell(self.cell_start, 4, dynamo_status)
                # Creation Date
                dynamo_created_date = str(table.get('CreationDateTime')).split(' ')[0]
                self.add_cell(self.cell_start, 5, dynamo_created_date)
                # Get Partition Key, Sort Key from KeySchema
                dynamo_part_key = "-"
                dynamo_sort_key = "-"
                for key in table.get('KeySchema'):
                    # HASH : Partition Key
                    if key.get('KeyType') == "HASH":
                        dynamo_part_key = key.get('AttributeName')
                    # RANGE : Sort Key
                    if key.get('KeyType') == 'RANGE':
                        dynamo_sort_key = key.get('AttributeName')

                # Get Key`s Attribute Type (Number, String, Binary)
                for attribute in table.get('AttributeDefinitions'):
                    if attribute.get('AttributeName') == dynamo_part_key:
                        dynamo_part_key_type = str(attribute.get('AttributeType'))
                        if dynamo_part_key_type == "S":
                            dynamo_part_key_type = "String"
                        elif dynamo_part_key_type == "N":
                            dynamo_part_key_type = "Number"
                        else:
                            dynamo_part_key_type = "Binary"
                    try:
                        if attribute.get('AttributeName') == dynamo_sort_key:
                            dynamo_sort_key_type = attribute.get('AttributeType')
                            if dynamo_sort_key_type == "S":
                                dynamo_sort_key_type = "String"
                            elif dynamo_sort_key_type == "N":
                                dynamo_sort_key_type = "Number"
                            else:
                                dynamo_sort_key_type = "Binary"
                    except:
                        pass
                # Primary Key
                dynamo_partition_key = "-"
                if dynamo_part_key != "-":
                    dynamo_partition_key = str(dynamo_part_key) + " (" + dynamo_part_key_type + ")"
                self.add_cell(self.cell_start, 6, dynamo_partition_key)
                # Sort Key
                dynamo_sorted_key = "-"
                if dynamo_sort_key != "-":
                    dynamo_sorted_key = str(dynamo_sort_key) + " (" + dynamo_sort_key_type + ")"
                self.add_cell(self.cell_start, 7, dynamo_sorted_key)
                # Total Read Capacity
                self.add_cell(self.cell_start, 8, table.get('ProvisionedThroughput').get('ReadCapacityUnits'))
                # Total Write Capacity
                self.add_cell(self.cell_start, 9, table.get('ProvisionedThroughput').get('WriteCapacityUnits'))
                # Table Size
                readable_dynamo_table_size = self.return_humanbytes(table.get('TableSizeBytes'))
                self.add_cell(self.cell_start, 10, readable_dynamo_table_size)
                # Item Count
                self.add_cell(self.cell_start, 11, table.get('ItemCount'))
                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")