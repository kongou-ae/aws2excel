exports.handler = function(event, context) {

    var async = require('async'); 
    var AWS = require('aws-sdk');
    var Excel = require("exceljs");
    var ebs = require('./ebs.js')
    var iamUsers = require('./iam_users.js')
    var workbook = new Excel.Workbook();
    var ec2Array = [];
    var sgArray = []
    var iamUserArray = []

    // この書き方で.aws/credentialsの情報は取れるみたい。
    var ec2 = new AWS.EC2({
        region: 'ap-northeast-1'
    });

    var iam = new AWS.IAM({
        region: 'ap-northeast-1'
    });
    
    async.series([
        // sdkを使い、インスタンスの情報を取得する
        // ToDo：エラーハンドリングを書く
        function getEc2Info(next){

            var ec2Obj = {};
            ec2.describeInstances(function(err, result) {
                for (var k = 0; k < Object.keys(result.Reservations).length; k++) {
                    ec2Obj = result.Reservations[k].Instances[0]
                    ec2Array.push(ec2Obj);
                }
                next()
            })
        },
        function getSecurityGroupInfo(next){

            var sgObj = {};
            ec2.describeSecurityGroups(function(err, result) {
                for (var k = 0; k < result.SecurityGroups.length; k++) {
                    sgObj = result.SecurityGroups[k]
                    sgArray.push(sgObj)
                }
                next()
            })
        },
        // 動作確認用
        /*
        function printInfo (next){
            var util = require('util'); // 
            console.log(util.inspect(sgArray,false,null))
            console.log(util.inspect(ec2Array,false,null))
            next()
        },
        */
        function buildEc2 (next) {
            // -----------------------------------------------------------------------------------------
            // EC2(summary)シートを作る
            // -----------------------------------------------------------------------------------------

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
                cell.alignment = {
                    horizontal: "center" 
                    
                };
            });
                

            for (var i = 0; i < ec2Array.length; i++) {
                // 2行目以降にインスタンスの情報を記載
                
                if (ec2Array[i].PublicIpAddress == null){
                    var PublicIpAddress = '-'
                } else {
                    var PublicIpAddress = ec2Array[i].PublicIpAddress
                }
                
                worksheet_ec2sum.addRow(
                    {   InstanceId: ec2Array[i].InstanceId,
                        ImageId: ec2Array[i].ImageId,
                        state: ec2Array[i].State.Name,
                        InstanceType: ec2Array[i].InstanceType,
                        PrivateIpAddress: ec2Array[i].PrivateIpAddress,
                        PublicIpAddress: PublicIpAddress,
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
            next()
        },
        function buildSg (next) {
            
            // PortRangeを元に、APIの戻り値に対応するマネジメントコンソールの見た目を生成する
            function checkPortRange (from,to) {
                
                if (from == -1) {
                    return 'N/A'
                }
                
                if (from == null){
                    return 'All'
                }
                
                if (from == to){
                    return from
                } else {
                    return from + ' - ' + to   
                }
            }
            
            // IpProtocolを元に、APIの戻り値に対応するマネジメントコンソールの見た目を生成する
            function checkIpProtocol (IpProtocol){
                // sg-a9822eccで要動作確認。何かが間違っている
                if (IpProtocol == -1){
                    return 'All'
                } else {
                    return IpProtocol
                }
                
            }
            
            // IpProtocolとPortRangeから、Typeフィールドのマネジメントコンソールの見た目を生成する
            function determineType (IpProtocol,PortRange) {
                
                var sample = IpProtocol+':'+PortRange
                
                var TypeObj = {
                    'All:All':'All traffic',
                    'icmp:N/A':'All ICMP',
                    'tcp:All':'All TCP',
                    'udp:All':'All UDP',
                    'tcp:22':'SSH',
                    'tcp:23':'telnet',
                    'tcp:25':'SMTP',
                    'tcp:42':'nameserver',
                    'udp:53':'DNS(UDP)',
                    'tcp:53':'DNS(TCP)',
                    'tcp:80':'HTTP',
                    'tcp:110':'POP3',
                    'tcp:143':'IMAP',
                    'tcp:389':'LDAP',
                    'tcp:443':'HTTPS',
                    'tcp:465':'SMTPS',
                    'tcp:993':'IMAPS',
                    'tcp:9955':'POP3S',
                    'tcp:1433':'MS SQL',
                    'tcp:3306':'MySQL/Aurora',
                    'tcp:3389':'RDP',
                }
                
                if (sample in TypeObj){
                    return TypeObj[sample]
                }
                
                if (/tcp:/.test(sample)) {
                    return 'Custom TCP Rule'
                }
                    
                if (/udp:/.test(sample)){
                    return 'Custom UDP Rule'
                }

                if (/icmp:/.test(sample)){
                    return 'Custom ICMP Rule'
                }
                
                return sample
                
            }
            
            // -----------------------------------------------------------------------------------------
            // SecurityGroupシートを作る
            // -----------------------------------------------------------------------------------------

            // 枠だけ作る
            var worksheet_sg = workbook.addWorksheet("SecurityGroup");
            var rowPos = 0
            var row = 0
            var sgSource = ''
            var sgDestination = ''
             
            //var row = worksheet_sg.getRow(3);
            //row = worksheet_sg.lastRow;

            // A-Eのwidthを15に変更
            for (var j = 1; j <= 5 ; j++) {
                worksheet_sg.getColumn(j).width = '17'
            }
            
            var borderStyle = {
                top: {style:"thin"},
                left: {style:"thin"},
                bottom: {style:"thin"},
                right: {style:"thin"}                   
            }
            
            var fillStyle = {
                type: "pattern",
                pattern:"solid",
                fgColor:{argb:"00191970"}
            };
            var fontStyle = {
                color: { argb: "00FFFFFF" },
                bold: true
            };

            // 各セキュリティグループごとに処理を実施
            for (var j = 0; j < sgArray.length; j++) {

                // まずはルール以外の箇所を描画
                row = worksheet_sg.getRow(1 + rowPos);
                row.values = ['GroupId',sgArray[j].GroupId]
                worksheet_sg.mergeCells("B" +(1 + rowPos) + ":E" + (1 + rowPos));
                row.eachCell(function(cell,colNumber){
                    cell.border = borderStyle
                    // 先頭列の装飾を実施
                    if (colNumber == 1){
                        cell.fill = fillStyle
                        cell.font = fontStyle
                    } 
                })

                row = worksheet_sg.getRow(2 + rowPos);
                row.values = ['GroupName',sgArray[j].GroupName]
                worksheet_sg.mergeCells("B" +(2 + rowPos) + ":E" + (2 + rowPos));
                row.eachCell(function(cell,colNumber){
                    cell.border = borderStyle                        
                    // 先頭列の装飾を実施
                    if (colNumber == 1){
                        cell.fill = fillStyle
                        cell.font = fontStyle
                    } 
                })

                row = worksheet_sg.getRow(3 + rowPos);
                row.values = ['Description',sgArray[j].Description]
                worksheet_sg.mergeCells("B" +(3 + rowPos) + ":E" + (3 + rowPos));
                row.eachCell(function(cell,colNumber){
                    cell.border = borderStyle
                    // 先頭列の装飾を実施
                    if (colNumber == 1){
                        cell.fill = fillStyle
                        cell.font = fontStyle
                    } 

                })

                row = worksheet_sg.getRow(4 + rowPos);
                if (sgArray[j].VpcId == null ){
                    row.values = ['VpcId','-']   
                } else {
                    row.values = ['VpcId',sgArray[j].VpcId] 
                }
                worksheet_sg.mergeCells("B" +(4 + rowPos) + ":E" + (4 + rowPos));
                row.eachCell(function(cell,colNumber){
                    cell.border = borderStyle
                    // 先頭列の装飾を実施
                    if (colNumber == 1){
                        cell.fill = fillStyle
                        cell.font = fontStyle
                    } 
                    
                })

                row = worksheet_sg.getRow(5 + rowPos);
                row.values = ['Rule','Type','Protocol','Port Range','Src/Dst']
                row.eachCell(function(cell,colNumber){
                    // ここの行だけ、全てのセルを装飾
                    cell.border = borderStyle                        
                    cell.fill = fillStyle
                    cell.font = fontStyle

                    if (colNumber != 1){
                        cell.alignment = { horizontal:'center' }
                    }                    
                    
                })

                // Ingressのルールを描画
                for (var k = 0; k < sgArray[j].IpPermissions.length; k++) {
                    row = worksheet_sg.getRow(worksheet_sg.lastRow.number + 1)
                    
                    // 送信元がIPアドレスかSGかのチェック
                    if ( sgArray[j].IpPermissions[k].IpRanges[0] == null ) {
                        sgSource = sgArray[j].IpPermissions[k].UserIdGroupPairs[0].GroupId
                    } else {
                        sgSource = sgArray[j].IpPermissions[k].IpRanges[0].CidrIp
                    }
                    
                    // FromとToからsgPortRangeを生成
                    var sgPortRange = checkPortRange(sgArray[j].IpPermissions[k].FromPort,sgArray[j].IpPermissions[k].ToPort)

                    // IpProtocolからsgIpProtocolを生成
                    var sgIpProtocol = checkIpProtocol(sgArray[j].IpPermissions[k].IpProtocol)

                    //  sgIpProtocolとsgPortRangeからsgTypeを生成
                    var sgType = determineType(sgIpProtocol,sgPortRange)

                    row.values = [
                        'Ingress', // Rule
                        sgType, // Type
                        sgIpProtocol, // Protocol
                        sgPortRange,
                        sgSource, // Source 
                        ]
                        
                    row.eachCell(function(cell,colNumber){
                        cell.border = borderStyle                        
                        // 先頭列の装飾を実施
                        if (colNumber == 1){
                            cell.fill = fillStyle
                            cell.font = fontStyle
                        } else {
                            // それ以外はセンタリング
                            cell.alignment = { horizontal:'center' }
                        }

                    })

                }
                
                // Egressのルールを描画
                for (var k = 0; k < sgArray[j].IpPermissionsEgress.length; k++) {
                    row = worksheet_sg.getRow(worksheet_sg.lastRow.number + 1)
                    
                    // 宛先IPアドレスかSGかのチェック
                    if ( sgArray[j].IpPermissionsEgress[k].IpRanges[0] == null ) {
                        sgDestination = sgArray[j].IpPermissionsEgress[k].UserIdGroupPairs[0].GroupId
                    } else {
                        sgDestination = sgArray[j].IpPermissionsEgress[k].IpRanges[0].CidrIp
                    }

                    // FromとToからsgPortRangeを生成
                    var sgPortRange = checkPortRange(sgArray[j].IpPermissionsEgress[k].FromPort,sgArray[j].IpPermissionsEgress[k].ToPort)

                    // IpProtocolからsgIpProtocolを生成
                    var sgIpProtocol = checkIpProtocol(sgArray[j].IpPermissionsEgress[k].IpProtocol)

                    //  sgIpProtocolとsgPortRangeからsgTypeを生成
                    var sgType = determineType(sgIpProtocol,sgPortRange)

                    row.values = [
                        'Egress', // Rule
                        sgType, // Type
                        sgIpProtocol , // Protocol
                        sgPortRange, // Port Range
                        sgDestination, // Source 
                        ]

                    row.eachCell(function(cell,colNumber){
                        cell.border = borderStyle                        
                        // 先頭列の装飾を実施
                        if (colNumber == 1){
                            cell.fill = fillStyle
                            cell.font = fontStyle
                        } else {
                            // それ以外はセンタリング
                            cell.alignment = { horizontal:'center' }
                        }

                    })
                }
                // 次のセキュリティグループに行くタイミングで、1行開けるために＋１
                rowPos = worksheet_sg.lastRow.number + 1
                
            } 
            next()
        },
        function buildEbs(next){
            console.log('EBS is started')
            ec2.describeVolumes(function(err, result){
                console.log(ebs(result.Volumes,workbook))
                next()
            })
        },
        // ToDo書き直せるか考える。とりあえず動くのをリリース
        function buildIam(next){
            console.log('IAM(Users) is started')
            async.series([
                function listUsers(next){
                    // UserNameとCreateDateを生成
                    iam.listUsers(function(err, listUsers){
                        for (var k = 0; k < listUsers['Users'].length; k++){
                            var iamUserObj = {}
                            iamUserObj.UserName = listUsers['Users'][k].UserName
                            iamUserObj.CreateDate = listUsers['Users'][k].CreateDate
                            iamUserObj.LoginProfile = ''
                            iamUserArray.push(iamUserObj)
                        }
                        next()
                    })
                },
                function getLoginProfile(next){
                    k = 0
                    async.forEachSeries(iamUserArray,function(data,finish){
                        iam.getLoginProfile( {UserName:data.UserName} ,function(err,getLoginProfile){
                            var data = iamUserArray[k]
                            if (getLoginProfile == null){
                                data.LoginProfile = 'false'
                            } else {
                                data.LoginProfile  = 'true'
                            }
                            iamUserArray.splice(k,1,data)
                            k = k + 1
                            finish()
                        })
                    },function(){
                        next()
                    })
                },
                function ListAccessKeys(next){
                    k = 0
                    async.forEachSeries(iamUserArray,function(data,finish){
                        iam.listAccessKeys( {UserName:data.UserName} ,function(err,ListAccessKeys){
                            var data = iamUserArray[k]
                            switch (ListAccessKeys.AccessKeyMetadata.length){
                                case 0:
                                    data.ListAccessKeys1 = '-'
                                    data.ListAccessKeys2 = '-'
                                    break;
                                case 1:
                                    data.ListAccessKeys1 = ListAccessKeys.AccessKeyMetadata[0].AccessKeyId + '(' + ListAccessKeys.AccessKeyMetadata[0].Status + ')'
                                    data.ListAccessKeys2 = '-'
                                    break;
                                case 2:
                                    data.ListAccessKeys1 = ListAccessKeys.AccessKeyMetadata[0].AccessKeyId + '(' + ListAccessKeys.AccessKeyMetadata[0].Status + ')'
                                    data.ListAccessKeys2 = ListAccessKeys.AccessKeyMetadata[1].AccessKeyId + '(' + ListAccessKeys.AccessKeyMetadata[1].Status + ')'
                                    break;
                            }
                            iamUserArray.splice(k,1,data)
                            k = k + 1
                            finish()
                        })
                    },function(){
                        next()
                    })
                },
                function listGroupsForUser(next){
                    k = 0
                    async.forEachSeries(iamUserArray,function(data,finish){                    
                        iam.listGroupsForUser( {UserName:data.UserName} ,function(err,listGroupsForUser){
                            var data = iamUserArray[k]
                            if ( listGroupsForUser == null)
                                data.listGroupsForUser == '-'
                            else {
                                var tmp = ''
                                for(var l = 0; l < listGroupsForUser.Groups.length; l++){
                                    tmp = tmp + listGroupsForUser.Groups[l].GroupName + '\r\n'
                                }
                                data.listGroupsForUser = tmp
                            }
                            iamUserArray.splice(k,1,data)
                            k = k + 1
                            finish()
                        })
                    },function(){
                        next()
                    })
                },
                function listAttachedUserPolicies(next){
                    k = 0
                    async.forEachSeries(iamUserArray,function(data,finish){                    
                        iam.listAttachedUserPolicies( {UserName:data.UserName} ,function(err,listAttachedUserPolicies){
                            var data = iamUserArray[k]
                            if ( listAttachedUserPolicies == null){
                                data.AttachedPolicies == '-'
                            } else {
                                var tmp = ''
                                for(var l = 0; l < listAttachedUserPolicies.AttachedPolicies.length; l++){
                                    tmp = tmp + listAttachedUserPolicies.AttachedPolicies[l].PolicyName + '\r\n'
                                }
                                data.AttachedPolicies = tmp
                            }
                            iamUserArray.splice(k,1,data)
                            k = k + 1
                            finish()
                        })
                    },function(){
                        next()
                    })
                },
                function buildIamUsers(next){
                    console.log(iamUsers(iamUserArray,workbook))
                    next()
                }
            ],function(){
                next()                
            })

        },
        function outputToExcel(next){
            console.log('building Excel....')
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