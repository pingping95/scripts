AWS Resource를 Excel 파일로 자동 저장시켜주는 Script입니다.

1. ~/.aws/config와 ~/.aws/credentials에 있는 Profile Name을 기준으로 Access Key와 Secret key를 통해 SDK 인증을 실시합니다.

2. 모든 profile name을 dict in list 형식으로 만들어 for문을 돌립니다.
    => 모든 Profile에 대하여 Excel 파일을 생성합니다.

3. main.py 내에 어떤 디렉터리에 저장할 것인지 지정할 수 있습니다.

4. log_txt 파일을 통해 어느 리소스가 없었으며, 어느 부분에서 Error가 발생하였는지 파악할 수 있습니다.