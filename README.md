# aws2excel

## 説明

AWSの構成情報をExcelに書き出すLambdaファンクションを作りたい。
現時点では、EC2の情報をローカルのxlsxファイルに書き出すだけのスクリプトです。

## サンプル

[aws_configuration.xlsx](https://github.com/kongou-ae/aws2excel/raw/master/sample/aws_configuration.xlsx)

## EC2

![ec2_summary](https://raw.githubusercontent.com/kongou-ae/aws2excel/develop/sample/ex2_summary.png)

![ec2_detail](https://raw.githubusercontent.com/kongou-ae/aws2excel/develop/sample/ec2_detail.png)

## SecurityGroup

![securitygrou@](https://raw.githubusercontent.com/kongou-ae/aws2excel/develop/sample/securitygroup.png)


## 使い方

```
git clone https://github.com/kongou-ae/aws2excel.git
cd aws2excel
npm install
node main.js
```