var http = require("http"),
    mongo = require("mongodb"),
    url = require("url");
    //querystring = require("querystring");
var xlsx = require("node-xlsx");

function saveDateToMongo(fileName){
    console.log("kfjak");
    var db = new mongo.Db("test", new mongo.Server('127.0.0.1', 27017, {auto_reconnect:true}), {safe: true});
    db.open(function () {
        db.collection("user", function (err, collection) {

            //console.log("fileName:"+fileName);
            //var list = xlsx.parse("C:\\Users\\Administrator\\Desktop\\测序单\\副本RNA+20150430+CG-1Ad WGS出库信息.xlsx");
            var list = xlsx.parse(fileName);

            for(i=2;i<list[0]["data"].length;i++) {

                /*
                 '建库原始板号',          originalPlateNo
                 '建库原始孔位',          originalHoleNo
                 '任务单名称',           taskName
                 'CLS',                 CLS
                 'Well',                 Well
                 'Sample ID',            SampleID
                 'DID',                 DID
                 'SID',                 SID
                 'Pooling基数',           PoolingBase
                 'barcode',             barcode
                 '样品名称*',           sampleName
                 '样品编号*',           sampleNo
                 '建库起始时间',          libraryBeginDate
                 '建库结束时间',          libraryEndDate
                 '出库浓度（ng/ul）',     libraryConcentration
                 '芯片名称',             chipName
                 '备注'                  remark
                 */
                var sample;
                var oneLine = list[0]["data"][i];
                if (typeof(oneLine[0]) == "string")
                    var originalPlateNo = oneLine[0];
                var originalHoleNo = oneLine[1];
                if (typeof(oneLine[2]) == "string")
                    var taskName = oneLine[2];
                if (typeof(oneLine[3]) == "string")
                    var CLS = oneLine[3];
                if (typeof(oneLine[4]) == "string")
                    var Well = oneLine[4];
                if (typeof(oneLine[5]) == "string")
                    var SampleID = oneLine[5];
                if (typeof(oneLine[6]) == "string")
                    var DID = oneLine[6];
                if (typeof(oneLine[7]) == "string")
                    var SID = oneLine[7];
                if (typeof(oneLine[8]) == "number")
                    var PoolingBase = oneLine[8];
                var barcode = oneLine[9];
                var sampleName = oneLine[10];
                var sampleNo = oneLine[11];
                var libraryBeginDate = oneLine[12];
                var libraryEndDate = oneLine[13];
                var libraryConcentration = oneLine[14];
                var chipName = oneLine[15];
                var remark = oneLine[16];
                sample = {
                    "originalPlateNo": originalPlateNo,
                    "originalHoleNo": originalHoleNo,
                    "taskName": taskName,
                    "CLS": CLS,
                    "Well": Well,
                    "SampleID": SampleID,
                    "DID": DID,
                    "SID": SID,
                    "PoolingBase": PoolingBase,
                    "barcode": barcode,
                    "sampleName": sampleName,
                    "sampleNo": sampleNo,
                    "libraryBeginDate": libraryBeginDate,
                    "libraryEndDate": libraryEndDate,
                    "libraryConcentration": libraryConcentration,
                    "chipName": chipName,
                    "remark": remark
                }
                collection.insert(sample, { safe: true }, function (err,result) {});
            }

        });
    });
}



var server = http.createServer();

var querystring = require('querystring');

var firstPage = function(res){

    res.writeHead(200, {'Content-Type': 'text/html'});

    var html = '<html>'+
        '<head>'+
        '<meta http-equiv="Content-Type" '+
        'content="text/html; charset=UTF-8" />'+
        '</script>'+
        '</head>'+
        '<body>'+
        '<form action="/save" method="post">'+
        'name:<input type="file" name="fileName"> </br>'+
        '<input type="submit" value="save">'+
        '</form>'+
        '</body></html>';

    res.end(html);

}

var save = function(req, res) {

    var info ='';

    req.addListener('data', function(chunk){

        info += chunk;

    })

        .addListener('end', function(){

            info = querystring.parse(info);
            console.log("info: "+ info);
            console.log("type: "+typeof(info.fileName));
            var fileNameStr=new String(info.fileName);
            var postfix=fileNameStr.slice(-4);
            console.log(postfix);

            if(postfix == "xlsx"){

                //
                var db = new mongo.Db("test", new mongo.Server('127.0.0.1', 27017, {auto_reconnect:true}), {safe: true});
                db.open(function () {
                    db.collection("user", function (err, collection) {

                        var filename=info.fileName.replace(/\\/g,"\\");
                        //var list = xlsx.parse(info.fileName);
                        var list = xlsx.parse(filename);

                        for(i=2;i<list[0]["data"].length;i++) {

                            /*
                             '建库原始板号',          originalPlateNo
                             '建库原始孔位',          originalHoleNo
                             '任务单名称',           taskName
                             'CLS',                 CLS
                             'Well',                 Well
                             'Sample ID',            SampleID
                             'DID',                 DID
                             'SID',                 SID
                             'Pooling基数',           PoolingBase
                             'barcode',             barcode
                             '样品名称*',           sampleName
                             '样品编号*',           sampleNo
                             '建库起始时间',          libraryBeginDate
                             '建库结束时间',          libraryEndDate
                             '出库浓度（ng/ul）',     libraryConcentration
                             '芯片名称',             chipName
                             '备注'                  remark
                             */
                            var sample;
                            var oneLine = list[0]["data"][i];
                            if (typeof(oneLine[0]) == "string")
                                var originalPlateNo = oneLine[0];
                            var originalHoleNo = oneLine[1];
                            if (typeof(oneLine[2]) == "string")
                                var taskName = oneLine[2];
                            if (typeof(oneLine[3]) == "string")
                                var CLS = oneLine[3];
                            if (typeof(oneLine[4]) == "string")
                                var Well = oneLine[4];
                            if (typeof(oneLine[5]) == "string")
                                var SampleID = oneLine[5];
                            if (typeof(oneLine[6]) == "string")
                                var DID = oneLine[6];
                            if (typeof(oneLine[7]) == "string")
                                var SID = oneLine[7];
                            if (typeof(oneLine[8]) == "number")
                                var PoolingBase = oneLine[8];
                            var barcode = oneLine[9];
                            var sampleName = oneLine[10];
                            var sampleNo = oneLine[11];
                            var libraryBeginDate = oneLine[12];
                            var libraryEndDate = oneLine[13];
                            var libraryConcentration = oneLine[14];
                            var chipName = oneLine[15];
                            var remark = oneLine[16];
                            sample = {
                                "originalPlateNo": originalPlateNo,
                                "originalHoleNo": originalHoleNo,
                                "taskName": taskName,
                                "CLS": CLS,
                                "Well": Well,
                                "SampleID": SampleID,
                                "DID": DID,
                                "SID": SID,
                                "PoolingBase": PoolingBase,
                                "barcode": barcode,
                                "sampleName": sampleName,
                                "sampleNo": sampleNo,
                                "libraryBeginDate": libraryBeginDate,
                                "libraryEndDate": libraryEndDate,
                                "libraryConcentration": libraryConcentration,
                                "chipName": chipName,
                                "remark": remark
                            }
                            collection.insert(sample, { safe: true }, function (err,result) {});
                        }

                    });
                });
                //
                res.writeHead(200, {"Content-Type": "text/html;charset:utf-8"});
                res.end('save success ' + info.fileName);

            }else{

                res.writeHead(200, {"Content-Type": "text/html;charset:utf-8"});
                res.end('save failed ');

            }

        })

}

var requestFunction = function (req, res){

    if(req.url == '/'){

        return firstPage(res);

    }

    if(req.url == '/save'){

        if (req.method != 'POST'){

            return;

        }

        return save(req, res);

        console.log(req,res);

    }

}

server.on('request',requestFunction);

server.listen(8089);

console.log('Server is running');