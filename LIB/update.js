var https = require('https');
var fs = require('fs');
var packageConfig = require('./package.json');

function sendToWin(type, msg) {
  console.log(msg)
}

function checkUpdate(cb) {
  getHttpsData('https://raw.githubusercontent.com/guanyuxin/fyvalid/master/package.json', function (res) {
    var data = JSON.parse(res);
    
    packageConfig.build = packageConfig.build || -1;
    if (packageConfig.build < data.build) {
      sendToWin('updateInfo', '检测到更新');
      cb && cb(data);
    } else {
      sendToWin('updateInfo', '最新版本');
    }
  })
}

checkUpdate(function (data) {
  fs.mkdir('./tmp', function () {
    var updateing = data.files.map(function(file, i) {
      return new Promise(function (resolve, reject) {
        getHttpsData('https://raw.githubusercontent.com/guanyuxin/fyvalid/master/' + file, function (res) {
          console.log('downloaded' + file);
          fs.writeFile('./tmp/' + file, res, function () {
            sendToWin('updateInfo', '下载' + file);
            resolve(file);
          });
        }, () => {
          reject();
        });
      });
    });
    Promise.all(updateing).then((res) => {
      sendToWin('updateInfo', '下载完成');
      var moveing = res.map((file, i) => {
        return new Promise((resolve, reject) =>{
          fs.rename('./tmp/' + file, './' + file, () => {
            resolve(file);
          })
        })
      })
      return Promise.all(moveing)
    }, () => {
      sendToWin('updateInfo', '更新失败');
    }).then(()=>{
      sendToWin('updateInfo', '更新完毕，重启生效');
    })
  })
});


function getHttpsData(filepath, success, error) {
  // 回调缺省时候的处理
  success = success || function () {};
  error = error || function () {};

  var url = filepath + '?r=' + Math.random();

  https.get(url, function (res) {
    var statusCode = res.statusCode;

    if (statusCode !== 200) {
      // 出错回调
      console.log(statusCode);
      error();
      // 消耗响应数据以释放内存
      res.resume();
      return;
    }

    res.setEncoding('utf8');
    var rawData = '';
    res.on('data', function (chunk) {
      rawData += chunk;
    });

    // 请求结束
    res.on('end', function () {
      // 成功回调
      success(rawData);
    }).on('error', function (e) {
      // 出错回调
      error();
    });
  });
};