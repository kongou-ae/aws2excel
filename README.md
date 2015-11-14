# aws2excel

## 説明

AWSの構成情報をExcelに書き出すLambdaファンクションを作りたい。
現時点では、以下サービスの情報をローカルのxlsxファイルに書き出すだけのスクリプトです。

## 対応しているサービス

- EC2
- EBS
- Security Group
- IAM(User)

## EC2

![ec2_summary](https://raw.githubusercontent.com/kongou-ae/aws2excel/master/sample/ex2_summary.png)

![ec2_detail](https://raw.githubusercontent.com/kongou-ae/aws2excel/master/sample/ec2_detail.png)

## SecurityGroup

![securitygroup](https://raw.githubusercontent.com/kongou-ae/aws2excel/master/sample/securitygroup.png)

## EBS

![EBS](https://raw.githubusercontent.com/kongou-ae/aws2excel/master/sample/ebs.png)

## IAM(Users)

![IAM(Users)](https://raw.githubusercontent.com/kongou-ae/aws2excel/master/sample/iam_users.png)

## 使い方

```
git clone https://github.com/kongou-ae/aws2excel.git
cd aws2excel
npm install
node main.js
```