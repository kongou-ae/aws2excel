# aws2excel

## 説明

AWSの構成情報をExcelに書き出すLambdaファンクションを作りたい。
現時点では、EC2の情報をローカルのxlsxファイルに書き出すだけのスクリプトです。

## サンプル

[aws_configuration.xlsx](https://github.com/kongou-ae/aws2excel/blob/master/sample/aws_configuration.xlsx)

## 使い方

```
git clone https://github.com/kongou-ae/aws2excel.git
cd aws2excel
npm install
node main.js
```