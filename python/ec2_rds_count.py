# Ctrl+F -> User Name 변경 필수

import boto3
def count_ec2(r_n):
    ec2_cli = session.client('ec2', region_name=r_n)
    response = ec2_cli.describe_instances()

    EC2_REGION_COUNT = 0
    if response['Reservations']:
        for idx, reservation in enumerate(response['Reservations']):
            for Instances in reservation['Instances']:
                # No.
                EC2_REGION_COUNT = EC2_REGION_COUNT+1
                # print(EC2_REGION_COUNT)
                # Availability Zone
                Zone = Instances['Placement'].get('AvailabilityZone')
                # print(Zone)
                # Instance Name
                InstanceName = '-'
                try:
                    for i in Instances['Tags']:
                        if i.get('Key') == 'Name':
                            InstanceName = i.get('Value')
                            break
                    # print(InstanceName)
                except:
                    pass
                    # print('-')
                # Status
                Status = Instances['State'].get('Name')
                # print(Status)
                # AMI (OS)
                try:
                    OS = Instances['Platform']
                    # print(OS)
                except:
                    pass
                    # print('Linux')
                # Instance Type
                Type = Instances['InstanceType']
                # print(Type)
                # Subnet Type
                if Instances.get('PublicIpAddress'):
                    pass
                    # print('public')
                else:
                    pass
                    # print('private')
                # VPC ID
                VPC_ID = Instances['VpcId']
                # print(VPC_ID)
                # Subnet ID
                Subnet_ID = Instances['SubnetId']
                # print(Subnet_ID)
                # Instance ID
                Instance_ID = Instances['InstanceId']
                # print(Instance_ID)
                # Private IP
                Private_IP = Instances['PrivateIpAddress']
                # print(Private_IP)
                # Public IP
                if Instances.get('PublicIpAddress'):
                    Public_IP = Instances['PublicIpAddress']
                    # print( Public_IP)
                else:
                    pass
                    # print('-')
                # EIP
                # Key Pair
                try:
                    KeyPair = Instances['KeyName']
                    # print(KeyPair)
                except:
                    pass
                    # print('-')
                # Security Group
                try:
                    for sec in Instances['SecurityGroups']:
                        SG = sec.get('GroupName')
                        # print(SG)
                except:
                    pass
                    # print('-')
                # IAM role
                try:
                    IAM = Instances['IamInstanceProfile'].get('Arn').split('/')[-1]
                    pass
                    # print(IAM)
                except:
                    pass
                    # print('-')
        print("EC2 resources are "+str(EC2_REGION_COUNT)+"EA in "+ r_n)
        return EC2_REGION_COUNT
    else:
        print("EC2 resources are not in " +r_n + " region")

def count_rds(r_n):
    RDS_REGION_COUNT = 0
    rds_cli = session.client('rds', region_name=r_n)

    rdsres = rds_cli.describe_db_instances()

    if rdsres['DBInstances']:

        for idx, rdsdata in enumerate(rdsres['DBInstances']):
            # No.
            RDS_REGION_COUNT = RDS_REGION_COUNT + 1
            # print(RDS_REGION_COUNT)
            # Availability Zone
            # print(rdsdata['AvailabilityZone'])
            # RDS Name
            # print(rdsdata['DBInstanceIdentifier'])
            # DB Engine
            # print(rdsdata['Engine'])
            # Engine Version
            # print(rdsdata['EngineVersion'])
            # DB Instance Class
            # print(rdsdata['DBInstanceClass'])
            # Storage Type
            # print(rdsdata['StorageType'])
            # Master Username
            # print(rdsdata['MasterUsername'])
            # VPC ID
            # print(rdsdata['DBSubnetGroup'].get('VpcId'))
        print("RDS resources are " + str(RDS_REGION_COUNT) + "EA in " + r_n)
        return RDS_REGION_COUNT
    else:
        print("RDS resources are not in " + r_n + " region")


if __name__ == "__main__":

    # 수정
    # User Name 변경 필수
    f = open('C:\/Users\/taehun.kim\/.aws/credentials', 'r')

    regions = ['us-east-1', 'us-east-2', 'us-west-1', 'us-west-2', 'ap-east-1', 'ap-south-1', 'ap-northeast-2',
               'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1', 'ca-central-1', 'eu-central-1', 'eu-west-1',
               'eu-west-2', 'eu-west-3', 'eu-north-1', 'me-south-1', 'sa-east-1']


    lines = f.readlines()
    name_list=[]
    access_list=[]
    secret_list=[]

    for line in lines:
        if '[' in line:
            name_list.append(line.split('[')[-1].split(']')[0])
        if 'aws_access_key_id =' in line:
            access_list.append(line.split('=')[-1].strip())
        if 'aws_secret_access_key =' in line:
            secret_list.append(line.split('=')[-1].strip())


    for name,access,secret in zip(name_list,access_list,secret_list):
        EC2_TOTAL_COUNT = 0
        RDS_TOTAL_COUNT = 0

        session = boto3.session.Session(
            aws_access_key_id= access,
            aws_secret_access_key=secret
        )
        try:
            print("++++++++++++" + name + "++++++++++++++++")
            for region in regions:
                try:
                    EC2_TOTAL_COUNT += count_ec2(region)
                except:
                    pass
                try:
                    RDS_TOTAL_COUNT += count_rds(region)
                except:
                    pass
        except Exception:
            print(name+"은 접근 권한이 없습니다.")
        print("++++++++++++" + "Total" + "++++++++++++++++")
        print(name +" EC2 resources are "+ str(EC2_TOTAL_COUNT)+ " EA")
        print(name + " RDS resources are " + str(RDS_TOTAL_COUNT) + " EA")