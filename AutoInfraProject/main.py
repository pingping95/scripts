import boto3
import openpyxl
from resources import vpc, nacl, subnet, routetable, transit_peering, nat_igw, vpn_vgw_cgw, security_group
from resources import elb

p_name = "zet"
r_name = "ap-northeast-2"

session = boto3.session.Session(profile_name=p_name, region_name=r_name)
log_txt = open("log.txt", "w")
wb = openpyxl.Workbook()
log_txt.write(f"Log\n\n")
log_txt.write(f"##################\nProfile: {p_name}\n##################\n")

if __name__ == "__main__":
        
    # For
    vpc.Vpc("VPC", wb, session, p_name, r_name, log_txt, True)
    # nacl.Nacl("NACL", wb, session, info_dict, log_txt, True)
    # subnet.Subnet("Subnet Group", wb, session, info_dict, log_txt, True)
    # routetable.RouteTable("Route Table", wb, session, info_dict, log_txt, True)
    # transit_peering.TransitPeering("Transit & Peering", wb, session, info_dict, log_txt, True)
    # nat_igw.NatIgw("NAT & IGW", wb, session, info_dict, log_txt, True)
    # vpn_vgw_cgw.VpnVgwCgw("VPN & VGW & CGW", wb, session, info_dict, log_txt, True)
    # security_group.SecurityGroup("Security Group", wb, session, info_dict, log_txt, True)
    # elb.Elb("ELB", wb, session, info_dict, log_txt, True)

    log_txt.close()
    wb.save("hello.xlsx")
    wb.close()