import boto3
import pprint


# 0. 변수 설정
p_name = ''
r_name = ''

print("=============================================================================================")
print(f"|고객사: {p_name}, 리전: {r_name} 의 EC2 정보 리스트입니다.")
print("==============================================================================================")

# 1. boto3 session을 이용하여 EC2 Res, Cli 객체 생성
session = boto3.session.Session(profile_name=p_name)
ec2_res = session.resource(service_name="ec2", region_name=r_name)
ec2_cli = session.client(service_name="ec2", region_name=r_name)


# 2. EC2 info
print('%5s' % "No.".ljust(5), '%21s' % "Region".ljust(21),'%31s' % "Instance Name".ljust(31), 
      '%16s' % "Private IP".ljust(16), '%16s' % "Public IP".ljust(16), '%25s' % "Key Pair".ljust(25))

response = ec2_cli.describe_instances()
if response['Reservations']:
    cnt = 0
    for idx, reservation in enumerate(response['Reservations']):
        try:
            for Instances in reservation['Instances']:
                # No.
                print('%4s' % str(idx + 1).ljust(4), end=': ')
                
                # Availability Zone
                Zone = Instances['Placement'].get('AvailabilityZone')
                print('%20s' % Zone.ljust(20), end = ', ')
                
                # Instance Name
                InstanceName = '-'
                try:
                    for i in Instances['Tags']:
                        if i.get('Key') == 'Name':
                            InstanceName = i.get('Value')
                            break
                except:
                    pass
                print('%30s' % InstanceName.ljust(30), end = ', ')
                
                # Private IP
                Private_IP = Instances['PrivateIpAddress']
                print('%15s' % str(Private_IP).ljust(15), end = ', ')

                # Public IP
                if Instances.get('PublicIpAddress'):
                    Public_IP = Instances['PublicIpAddress']
                    print('%15s' % str(Public_IP).ljust(15), end = ', ')

                else:
                    Public_IP = ''
                    print('%15s' % 'No Public IP'.ljust(15), end = ', ')
                    
                # Key Pair
                try:
                    KeyPair = Instances['KeyName']
                    print('%25s' % str(KeyPair).ljust(25))
                except:
                    print('%25s' % 'No Key Pair'.ljust(15))
                
        
        except Exception as e:
            cnt += 1
            print(e)

else:
    print("EC2가 없습니다.")