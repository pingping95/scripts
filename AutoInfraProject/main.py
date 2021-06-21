import boto3
import openpyxl
from resources import vpc, nacl, subnet, routetable, transit_peering, nat_igw, vpn_vgw_cgw

if __name__ == "__main__":
    try:
        info_dict = {"porfile": "default", 
                    "region": "ap-northeast-2"}
        session = boto3.session.Session(profile_name=info_dict.get('profile'))
        log_txt = open("log.txt", "w")
        wb = openpyxl.Workbook()
        log_txt.write(f"Log\n\n")
        log_txt.write(f"##################\nProfile: {info_dict.get('porfile')}\n##################\n")
        
        # For
        vpc.Vpc("VPC", wb, session, info_dict, log_txt, True)
        nacl.Nacl("NACL", wb, session, info_dict, log_txt, True)
        subnet.Subnet("Subnet Group", wb, session, info_dict, log_txt, True)
        routetable.RouteTable("Route Table", wb, session, info_dict, log_txt, True)
        transit_peering.TransitPeering("Transit & Peering", wb, session, info_dict, log_txt, True)
        nat_igw.NatIgw("NAT & IGW", wb, session, info_dict, log_txt, True)
        vpn_vgw_cgw.VpnVgwCgw("VPN & VGW & CGW", wb, session, info_dict, log_txt, True)
    
    finally:
        log_txt.close()
        wb.save("hello.xlsx")
        wb.close()