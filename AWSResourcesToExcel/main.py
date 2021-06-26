import boto3
import openpyxl
from resources import vpc, nacl, subnet, routetable, transit_peering, nat_igw, vpn_vgw_cgw, security_group, elb, ec2, \
    rds, s3, cloudfront, cloudtrail, cloudwatch, lmda, route53, iam_user, iam_role, ecs_eks_ecr, elasticache, efs, \
    elasticsearch
import os
import datetime

storing_path = f'G:\\공유 드라이브\\CSU 공유 드라이브\\팀별 공유 드라이브\\DevOps 상품 본부\\SRE2 센터\\SRE 5팀\\SMB3\\[기타] 내부공유자료\\고객사 자산 내역\\'

today = datetime.date.today()

# dd/mm/YY
today_date = today.strftime("%Y%m%d")

f = open(f'C:\\Users\\{os.getlogin()}\\.aws\\config', "r")
lines = f.readlines()

p_name_list = []
r_name_list = []

for line in lines:
    if "profile" in line:
        p_name = line.split(" ")[-1].split("]")[0]
        p_name_list.append(p_name)
    if "region" in line:
        r_name = line.split(" = ")[-1]
        if "\n" in r_name:
            r_name = r_name.split("\n")[0]
        r_name_list.append(r_name)

f.close()

cust_name_list = dict(zip(p_name_list, r_name_list))

log_txt = open(f"{storing_path}log.txt", "w")
log_txt.write(f"Log\n")


if __name__ == "__main__":

    for p_name, r_name in cust_name_list.items():
        try:
            wb = openpyxl.Workbook()
            session = boto3.session.Session(profile_name=p_name, region_name=r_name)
            log_txt.write(f"\n##################\nProfile: {p_name}\n##################\n")
            # For
            vpc.Vpc("VPC", wb, session, p_name, r_name, log_txt, True)
            nacl.Nacl("NACL", wb, session, p_name, r_name, log_txt, True)
            subnet.Subnet("Subnet Group", wb, session, p_name, r_name, log_txt, True)
            routetable.RouteTable("Route Table", wb, session, p_name, r_name, log_txt, True)
            security_group.SecurityGroup("Security Group", wb, session, p_name, r_name, log_txt, True)
            transit_peering.TransitPeering("Transit & Peering", wb, session, p_name, r_name, log_txt, True)
            nat_igw.NatIgw("NAT & IGW", wb, session, p_name, r_name, log_txt, True)
            vpn_vgw_cgw.VpnVgwCgw("VPN & VGW & CGW", wb, session, p_name, r_name, log_txt, True)
            elb.Elb("ELB", wb, session, p_name, r_name, log_txt, True)
            ec2.Ec2("EC2", wb, session, p_name, r_name, log_txt, True)
            rds.RDS("RDS", wb, session, p_name, r_name, log_txt, True)
            s3.S3("s3", wb, session, p_name, r_name, log_txt, True)
            cloudfront.CF("CloudFront", wb, session, p_name, r_name, log_txt, True)
            cloudtrail.CT("CloudTrail", wb, session, p_name, r_name, log_txt, True)
            cloudwatch.CW("CloudWatch", wb, session, p_name, r_name, log_txt, True)
            lmda.Lambda("Lambda", wb, session, p_name, r_name, log_txt, True)
            route53.Route53("Route53", wb, session, p_name, r_name, log_txt, True)
            iam_user.IamUser("Iam User", wb, session, p_name, r_name, log_txt, True)
            iam_role.IamRole("Iam Role", wb, session, p_name, r_name, log_txt, True)
            ecs_eks_ecr.EcsEks("ECS, EKS", wb, session, p_name, r_name, log_txt, True)
            elasticache.ElastiCache("ElastiCache", wb, session, p_name, r_name, log_txt, True)
            efs.Efs("EFS", wb, session, p_name, r_name, log_txt, True)
            elasticsearch.ElasticSearch("ElasticSearch", wb, session, p_name, r_name, log_txt, True)

            wb.save(storing_path + p_name + "_자산 내역_" + today_date + ".xlsx")

            print(f"{p_name} : 자산 내역 완료")

        except Exception as e:
            print(f"{p_name} : 자산 내역 실패, 내용: {e}")
            log_txt.write(f"main 문에서 Error 발생, 고객사 이름: {p_name}, 내용: {e}\n")

        finally:
            wb.close()

    log_txt.close()
