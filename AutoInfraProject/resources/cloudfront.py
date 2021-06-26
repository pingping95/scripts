from settings import Common


class CF(Common):
    def __init__(self, name, workbook, ses, profile, region, log, is_run=False):
        Common.__init__(self, name)
        if is_run:
            self.log = log
            self.wb = workbook
            self.profile = profile
            self.region = region
            self.client = ses.client(service_name="cloudfront", region_name=self.region)

            self.cloudfront = self.client.list_distributions().get('DistributionList')
            if "Items" in self.cloudfront:
                self.run()
            else:
                self.log.write(f"There is no {self.name}\n")

    def run(self):
        try:
            # Initialize
            self.sheet = self.wb.create_sheet(self.name)
            self.sheet.title = self.name
            # Cell width
            cell_widths = [6, 6, 16, 28, 47, 16, 30, 10, 10, 25, 8, 8, 8, 8, 8, 8, 8, 8, 7, 7, 7, 7]
            self.fit_cell_width(cell_widths)
            # Header
            self.make_header(self.cell_start, self.name)
            # Cell header
            cell_headers = ["No.", "ID", "Domain Name", "Origin Domain Names", "Origin Groups (EA)", "CNAMEs", "Status", "State", "Price Class"]
            self.make_cell_header(self.cell_start, cell_headers)
            # CF
            for idx, items in enumerate(self.cloudfront.get('Items')):
                # No.
                self.add_cell(self.cell_start, 2, idx + 1)
                # Id
                item_id = items.get("Id")
                self.add_cell(self.cell_start, 3, item_id)
                # Domain Name
                distribution_domain_name = items.get('DomainName')
                self.add_cell(self.cell_start, 4, distribution_domain_name)
                # Origin Domain Names
                origin_domain = ""
                origin = items.get('Origins')
                if origin.get('Quantity') != 0:
                    for i, item in enumerate(origin["Items"]):
                        if i != 0:
                            origin_domain += ", \n"
                        origin_domain += item.get('DomainName')
                else:
                    origin_domain = "-"
                self.add_cell(self.cell_start, 5, origin_domain)
                # Origin Groups (EA)
                origin_group_cnt = items.get('OriginGroups').get('Quantity')
                self.add_cell(self.cell_start, 6, origin_group_cnt)
                # cf_cnames
                cf_cnames = ""
                aliases = items.get('Aliases')
                if aliases.get('Quantity') != 0:
                    for i, cname in enumerate(aliases.get('Items')):
                        if i != 0:
                            cf_cnames += ", \n"
                        cf_cnames += cname
                else:
                    cf_cnames = "-"
                self.add_cell(self.cell_start, 7, cf_cnames)
                # cf_status (Deployed, ..)
                try:
                    cf_status = items.get('Status')
                except Exception:
                    cf_status = "-"
                self.add_cell(self.cell_start, 8, cf_status)
                # cf_state (Enabled, Disabled)
                cf_state = items.get('Enabled')
                if cf_state == True:
                    cf_state = "Enabled"
                elif cf_state == False:
                    cf_state = "Disabled"
                self.add_cell(self.cell_start, 9, cf_state)
                # PriceClass
                price_class = items.get('PriceClass')
                if "All" in price_class:
                    price_class = "All Edge Locations \n(Best Performance)"
                elif "100" in price_class:
                    price_class = "Only U.S, Canada and Europe"
                elif "200" in price_class:
                    price_class = "Use U.S, Canada, Europe, \nAsia, Middle East and Afreeca"
                self.add_cell(self.cell_start, 10, price_class)
                self.cell_start += 1
        except Exception as e:
            self.log.write(f"Error 발생, 리소스: {self.name}, 내용: {e}\n")