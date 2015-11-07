module.exports = function(Volumes,workbook){
        var worksheet_ebs = workbook.addWorksheet("EBS");
        var ebsObj = {}
        var ebsAry = []

        for (var i = 0; i < Object.keys(Volumes).length; i++) {
            ebsObj = Volumes[i]
            ebsAry.push(ebsObj);
        }
        
        // 1行目を追加
        worksheet_ebs.columns = [
            { header: 'VolumeId', key: 'VolumeId', width: 15 },
            { header: 'Size', key: 'Size', width: 15 },
            { header: 'VolumeType', key: 'VolumeType', width: 15 },
            { header: 'Iops', key: 'Iops', width: 15 },
            { header: 'SnapshotId', key: 'SnapshotId', width: 15 },
            { header: 'CreateTime', key: 'CreateTime', width: 15 },
            { header: 'AvailabilityZone', key: 'AvailabilityZone', width: 15 },
            { header: 'State:', key: 'State', width: 15 },
            { header: 'AttachmentInformation',  key: 'AttachmentInformation', width: 24 },
            { header: 'DeleteOnTermination', key: 'DeleteOnTermination', width : 15},
            { header: 'Tags', key: 'Tags', width: 15 },
            { header: 'Encrypted', key: 'Encrypted', width: 15 },
            { header: 'KmsKeyId', key: 'KmsKeyId', width: 81 },
        ]
        
        // 1行目を装飾する
        worksheet_ebs.getRow(1).eachCell({ includeEmpty: true }, function(cell, colNumber) {
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

        for (var i = 0; i < ebsAry.length; i++) {
            // 2行目以降にEBSの情報を記載

            // attacheされていないとDeleteOnTerminationがエラーになるので、事前に-を挿入
            if (ebsAry[i].Attachments[0] == null){
                var DeleteOnTermination = '-'
                var AttachmentInformation = '-'
            } else {
                var DeleteOnTermination = ebsAry[i].Attachments[0].DeleteOnTermination.toString()
                var AttachmentInformation = ebsAry[i].Attachments[0].InstanceId + ':' + ebsAry[i].Attachments[0].Device
            }

            // 暗号化されていないとebsAry[i].KmsKeyIdが存在せずエラーになるので、-を挿入
            if (ebsAry[i].KmsKeyId == null){
                var KmsKeyId = '-' 
            } else {
                var KmsKeyId = ebsAry[i].KmsKeyId
            }

            if (ebsAry[i].SnapshotId == ''){
                var SnapshotId = '-' 
            } else {
                var SnapshotId = ebsAry[i].SnapshotId
            }
            
            worksheet_ebs.addRow(
                {   VolumeId: ebsAry[i].VolumeId,
                    Size: ebsAry[i].Size + 'GB',
                    VolumeType: ebsAry[i].VolumeType,
                    Iops: ebsAry[i].Iops,
                    SnapshotId: SnapshotId,
                    CreateTime: ebsAry[i].CreateTime,
                    AvailabilityZone: ebsAry[i].AvailabilityZone,
                    State: ebsAry[i].State,
                    AttachmentInformation: AttachmentInformation,
                    DeleteOnTermination: DeleteOnTermination,
                    Tags: 'unsupported!!',//ebsAry[i].Tags,
                    Encrypted: ebsAry[i].Encrypted.toString(),
                    KmsKeyId: KmsKeyId
                });
            // 2行目以降に罫線を描画
            worksheet_ebs.getRow(2+i).eachCell({ includeEmpty: true }, function(cell, colNumber) {
                cell.border = {
                    top: {style:"thin"},
                    left: {style:"thin"},
                    bottom: {style:"thin"},
                    right: {style:"thin"}
                };
            });                    
        }

        return 'EBS is finished'
}


