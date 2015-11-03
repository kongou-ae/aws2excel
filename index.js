exports.handler = function(event, context) {

    var async = require('async'); 
    var AWS = require('aws-sdk');
    var Excel = require("exceljs");
    
    // この書き方で.aws/credentialsの情報は取れるみたい。
    var ec2 = new AWS.EC2({
        region: 'ap-northeast-1'
    });
 
    var ec2Obj = {};
    var ec2Array = []; 
    
    async.series([
        // sdkを使い、インスタンスの情報を取得する
        // ToDo：エラーハンドリングを書く
        function getEc2Info(next){
            ec2.describeInstances(function(err, result) {
                for (var k = 0; k < Object.keys(result.Reservations).length; k++) {
                    ec2Obj = result.Reservations[k].Instances[0]
                    ec2Array.push(ec2Obj);
                }
                next()
            })
        },
        // 動作確認用
        /*
        function printInfo (next){
            var util = require('util'); // 

            console.log(util.inspect(ec2Array,false,null))
            next()
        },
        */
        function buildExcel (next) {
            // -----------------------------------------------------------------------------------------
            // EC2(summary)シートを作る
            // -----------------------------------------------------------------------------------------
            var workbook = new Excel.Workbook();
            var worksheet_ec2sum = workbook.addWorksheet("EC2(summary)");

            // 1行目を追加
            worksheet_ec2sum.columns = [
                { header: "InstanceId", key: "InstanceId", width: 15 },
                { header: "ImageId", key: "ImageId", width: 15 },
                { header: "state", key: "state", width: 15 },
                { header: "InstanceType", key:"InstanceType", width: 15 },
                { header: "PrivateIpAddress", key:"PrivateIpAddress", width:15},
                { header: "PublicIpAddress", key:"PublicIpAddress", width:15},
                { header: "AvailabilityZone", key:"AvailabilityZone",width:15}
            ];
            
            // 1行目を装飾する
            worksheet_ec2sum.getRow(1).eachCell({ includeEmpty: true }, function(cell, colNumber) {
                cell.border = {
                    top: {style:"thin"},
                    left: {style:"thin"},
                    bottom: {style:"thin"},
                    right: {style:"thin"}
                };
                cell.fill = {
                    type: "pattern",
                    pattern:"solid",
                    fgColor:{argb:"00191970"}
                };
                cell.font = {
                    color: { argb: "00FFFFFF" },
                    bold: true
                };
            });
                

            for (var i = 0; i < ec2Array.length; i++) {
                // 2行目以降にインスタンスの情報を記載
                worksheet_ec2sum.addRow(
                    {   InstanceId: ec2Array[i].InstanceId,
                        ImageId: ec2Array[i].ImageId,
                        state: ec2Array[i].State.Name,
                        InstanceType: ec2Array[i].InstanceType,
                        PrivateIpAddress: ec2Array[i].PrivateIpAddress,
                        PublicIpAddress: ec2Array[i].PublicIpAddress,
                        AvailabilityZone: ec2Array[i].Placement.AvailabilityZone
                    });
                // 2行目以降に罫線を描画
                // ToDo：要素が空だと線が引かれない
                worksheet_ec2sum.getRow(2+i).eachCell({ includeEmpty: true }, function(cell, colNumber) {
                    cell.border = {
                        top: {style:"thin"},
                        left: {style:"thin"},
                        bottom: {style:"thin"},
                        right: {style:"thin"}
                    };
                });                    

            }
            // -----------------------------------------------------------------------------------------
            // EC2(detail)シートを作る
            // -----------------------------------------------------------------------------------------

            // ---------------------------------------------------------------------
            // 1列目を作成する
            // ---------------------------------------------------------------------

            // RootDeviceNameの最大枠を確認。

            var maxValueOfBlockDeviceMappings = 1
            for (var i = 0; i < ec2Array.length; i++) {
                if (Object.keys(ec2Array[i].BlockDeviceMappings).length > maxValueOfBlockDeviceMappings) {
                    maxValueOfBlockDeviceMappings = Object.keys(ec2Array[i].BlockDeviceMappings).length
                }
            }

            // SecurityGroupsの最大枠を確認。
            var maxValueOfSecurityGroups = 1
            for (var i = 0; i < ec2Array.length; i++) {
                if (Object.keys(ec2Array[i].SecurityGroups).length > maxValueOfSecurityGroups) {
                    maxValueOfSecurityGroups = Object.keys(ec2Array[i].SecurityGroups).length
                }
            }
            
            // Tagsの最大枠を確認。
            var maxValueOfTags = 1
            for (var i = 0; i < ec2Array.length; i++) {
                if (Object.keys(ec2Array[i].Tags).length > maxValueOfTags) {
                    maxValueOfTags = Object.keys(ec2Array[i].Tags).length
                }                
            }
            // NetworkInterfacesの最大枠を確認。
            var maxValueOfNetworkInterfaces = 1
            for (var i = 0; i < ec2Array.length; i++) {
                if (Object.keys(ec2Array[i].NetworkInterfaces).length > maxValueOfNetworkInterfaces) {
                    maxValueOfNetworkInterfaces = Object.keys(ec2Array[i].NetworkInterfaces).length
                }                
            }


            // 1列目のitemを定義
            var worksheet_ec2detail = workbook.addWorksheet("EC2(detail)");
            var choiceAry = [
                    'InstanceId',           // 0
                    'ImageId',              // 1
                    'State',                // 2
                    'PrivateDnsName',       // 3
                    'PublicDnsName',        // 4
                    'KeyName',              // 5
                    'InstanceType',         // 6
                    'AvailabilityZone',     // 7
                    'Tenancy',              // 8
                    'SubnetId',             // 9
                    'VpcId',                // 10
                    'PrivateIpAddress',     // 11
                    'Architecture',         // 12
                    'RootDeviceType',       // 13
                    'RootDeviceName',       // 14
                    'BlockDeviceMappings-1',// 15
                    'VirtualizationType',   // 16
                    'Tags-1',               // 17
                    'SecurityGroups-1',     // 18
                    'SourceDestCheck',      // 19
                    'Hypervisor',           // 20
                    'NetworkInterfaces-1',  // 21
                    'EbsOptimized'          // 22
                ]
            
            // 全てのインスタンスの中の最大値分、行を増やす
            // カッコ悪いので、もっとカッコいい処理を考える
            var ebsCount = 0
            for (var i = 0; i < maxValueOfBlockDeviceMappings -1; i++) { 
                ebsCount = 2 + i
                choiceAry.splice(16 + i,0,'BlockDeviceMappings-'+ ebsCount)                
            }
            
            // この時点でchoiceAryの総数が変更になっているので、位置をチェック
            var SecurityGroupsPostion = choiceAry.indexOf('SecurityGroups-1')

            // 全てのインスタンスの中の最大値分、行を増やす
            // カッコ悪いので、もっとカッコいい処理を考える
            var SecurityGroupsCount = 0
            for (var i = 0; i < maxValueOfSecurityGroups -1; i++) { 
                SecurityGroupsCount = 2 + i
                choiceAry.splice(SecurityGroupsPostion + 1 + i ,0,'SecurityGroups-'+ SecurityGroupsCount)                
            }

            // この時点でchoiceAryの総数が変更になっているので、位置をチェック
            var TagsPostion = choiceAry.indexOf('Tags-1')

            // 全てのインスタンスの中の最大値分、行を増やす
            // カッコ悪いので、もっとカッコいい処理を考える
            var TagsCount = 0
            for (var i = 0; i < maxValueOfTags -1; i++) { 
                TagsCount = 2 + i
                choiceAry.splice(TagsPostion + 1 + i ,0,'Tags-'+ TagsCount)                
            }
            
            // この時点でchoiceAryの総数が変更になっているので、位置をチェック
            var NetworkInterfacesPostion = choiceAry.indexOf('NetworkInterfaces-1')

            // 全てのインスタンスの中の最大値分、行を増やす
            // カッコ悪いので、もっとカッコいい処理を考える
            var NetworkInterfacesCount = 0
            for (var i = 0; i < maxValueOfNetworkInterfaces -1; i++) { 
                NetworkInterfacesCount = 2 + i
                choiceAry.splice(NetworkInterfacesPostion + 1 + i ,0,'NetworkInterfaces-'+ NetworkInterfacesCount)                
            }

            //console.log(choiceAry) // 動作確認用

            // 1列目を作成           
            var itemAry = []
            var itemColumn = worksheet_ec2detail.getColumn(1);
            itemColumn.width = '25'
            itemColumn.header = choiceAry
            itemColumn.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                cell.border = {
                    top: {style:"thin"},
                    left: {style:"thin"},
                    bottom: {style:"thin"},
                    right: {style:"thin"}
                };
                cell.fill = {
                    type: "pattern",
                    pattern:"solid",
                    fgColor:{argb:"00191970"}
                };
                cell.font = {
                    color: { argb: "00FFFFFF" },
                    bold: true
                };
            });   


            // ---------------------------------------------------------------------
            //2列目以降を作成 
            // ---------------------------------------------------------------------
            var valueAry = []
            var valueColumn;
            
            for (var i = 0; i < ec2Array.length; i++) {
                valueColumn = worksheet_ec2detail.getColumn(i+2);
                valueColumn.width = '46'

                // 2列目以降に入力する内容を作成
                valueAry = [
                    ec2Array[i].InstanceId,                     
                    ec2Array[i].ImageId,                        
                    ec2Array[i].State.Name,                     
                    ec2Array[i].PrivateDnsName,                 
                    ec2Array[i].PublicDnsName,                  
                    ec2Array[i].KeyName,                        
                    ec2Array[i].InstanceType,
                    ec2Array[i].Placement.AvailabilityZone,
                    ec2Array[i].Placement.Tenancy,
                    ec2Array[i].SubnetId,
                    ec2Array[i].VpcId,
                    ec2Array[i].PrivateIpAddress,
                    ec2Array[i].Architecture,
                    ec2Array[i].RootDeviceType,
                    ec2Array[i].RootDeviceName,
                    //'BlockDeviceMappings',
                    ec2Array[i].VirtualizationType,
                    'Tags',//Tags // SecurityGroupsのvalueを入れ替える位置を決定するための仮value
                    //SecurityGroups ec2Array
                    ec2Array[i].SourceDestCheck.toString(),
                    ec2Array[i].Hypervisor,
                    //NetworkInterfaces
                    ec2Array[i].EbsOptimized.toString()
                ]
                //console.log(valueAry) // 動作確認用
                
                // BlockDeviceMappingsを実際の値に入れ替え。不足分は-で埋める
                // dev/xvda is vol-6907d396 的なvalueにする
                for (var j = 0; j < maxValueOfBlockDeviceMappings; j++) {
                    if ( j < ec2Array[i].BlockDeviceMappings.length){
                        valueAry.splice(valueAry.indexOf(ec2Array[i].RootDeviceName) + 1 + j , 0 , ec2Array[i].BlockDeviceMappings[j].DeviceName + ' is ' + ec2Array[i].BlockDeviceMappings[j].Ebs['VolumeId'])
                    } else {
                        valueAry.splice(valueAry.indexOf(ec2Array[i].RootDeviceName) + 1 + j , 0 , '-') 
                    }
                }
                
                // SecurityGroupsを実際の値に入れ替え。不足分は-で埋める
                // GroupId(GroupName) 的なvalueにする
                for (var j = 0; j < maxValueOfSecurityGroups; j++) {
                    if ( j < ec2Array[i].SecurityGroups.length){
                        valueAry.splice(valueAry.indexOf('Tags') + 1 + j , 0 , ec2Array[i].SecurityGroups[j].GroupId + '(' + ec2Array[i].SecurityGroups[j].GroupName + ')')
                    } else {
                        valueAry.splice(valueAry.indexOf('Tags') + 1 + j , 0 , '-') 
                    }
                }
                // 後段の挿入処理のため、仮のTagsが残っていると、全体として1行多くなってしまう。
                // これを回避するため、SecurityGroups用に作っておいた仮のTagsを削除する
                valueAry.splice(valueAry.indexOf(ec2Array[i].VirtualizationType) + 1,1)

                // Tagsを実際の値に入れ替え。不足分は-で埋める
                // Key is Value的なvalueにする
                for (var j = 0; j < maxValueOfTags; j++) {
                    if ( j < ec2Array[i].Tags.length){
                        valueAry.splice(valueAry.indexOf(ec2Array[i].VirtualizationType) + 1 + j , 0 , ec2Array[i].Tags[j].Key + ' is ' + ec2Array[i].Tags[j].Value)
                    } else {
                        valueAry.splice(valueAry.indexOf(ec2Array[i].VirtualizationType) + 1 + j , 0 , '-') 
                    }
                }
                
                // NetworkInterfacesを実際の値に入れ替え。不足分は-で埋める
                // とりあえずENIのIDだけ
                for (var j = 0; j < maxValueOfNetworkInterfaces; j++) {
                    if ( j < ec2Array[i].NetworkInterfaces.length){
                        valueAry.splice(valueAry.indexOf(ec2Array[i].Hypervisor) + 1 + j , 0 , ec2Array[i].NetworkInterfaces[j].NetworkInterfaceId)
                    } else {
                        valueAry.splice(valueAry.indexOf(ec2Array[i].Hypervisor) + 1 + j , 0 , '-') 
                    }
                }                

                valueColumn.header = valueAry
                valueColumn.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                    cell.border = {
                        top: {style:"thin"},
                        left: {style:"thin"},
                        bottom: {style:"thin"},
                        right: {style:"thin"}
                    };
                });                    
            }

            // --------------------------------------------------------------
            // ファイルに書き出し
            // --------------------------------------------------------------
            workbook.xlsx.writeFile('aws_configuration.xlsx')
                .then(function() {
                    console.log('complete')
                    next()
                });
        }
    ])
}