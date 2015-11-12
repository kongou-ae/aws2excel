module.exports = function(iamUserArray,workbook){
    var worksheet_iamUsers = workbook.addWorksheet("IAM(Users)");

    // 1行目を追加
    worksheet_iamUsers.columns = [
        { header: 'UserName', key: 'UserName', width: 15 },
        { header: 'CreateDate', key: 'CreateDate', width: 15 },
        { header: 'Login', key: 'Login', width: 15 },
        { header: 'ListAccessKeys1', key: 'ListAccessKeys1', width: 32 },
        { header: 'ListAccessKeys2', key: 'ListAccessKeys2', width: 32 },
        { header: 'listGroupsForUser', key: 'listGroupsForUser', width: 20 },
        { header: 'AttachedPolicies', key: 'AttachedPolicies', width: 32 },
    ]

    // 1行目を装飾する
    worksheet_iamUsers.getRow(1).eachCell({ includeEmpty: true }, function(cell, colNumber) {
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
            vertical: "middle"
        };
    });

    for (var i = 0; i < iamUserArray.length; i++) {
        worksheet_iamUsers.addRow(
            {   UserName: iamUserArray[i].UserName,
                CreateDate: iamUserArray[i].CreateDate,
                Login: iamUserArray[i].LoginProfile,
                ListAccessKeys1: iamUserArray[i].ListAccessKeys1,
                ListAccessKeys2: iamUserArray[i].ListAccessKeys2,
                listGroupsForUser: iamUserArray[i].listGroupsForUser,
                AttachedPolicies: iamUserArray[i].AttachedPolicies,
            });        
            
        worksheet_iamUsers.getRow(2+i).eachCell({ includeEmpty: true }, function(cell, colNumber) {
            cell.border = {
                top: {style:"thin"},
                left: {style:"thin"},
                bottom: {style:"thin"},
                right: {style:"thin"}
            };
            cell.alignment = {
                vertical: "middle"
            };
        });  
    }
    return 'IAM(USers) is finished'

}