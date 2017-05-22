const http = require("http");
const cheerio = require('cheerio');
// const nodeExcel = require('excel-export');
const xlsx = require('node-xlsx').default;
const express = require('express');
const opn = require('opn');
const fs = require('fs');
let app = express();
let $;
let page = 1;
const url = `http://you.ctrip.com/searchsite/Travels?query=${encodeURI('德国乡村')}&isAnswered=&isRecommended=&publishDate=&PageNo=`
const context = 'http://you.ctrip.com';
let totalPage = 1;
let excelList = [];
// 获取http数据
function loadPage(url) {
    console.log('url:', url);
    var http = require('http');
    var pm = new Promise(function (resolve, reject) {
        http.get(url, function (res) {
            var body = [];
            res.on('data', function (chunck) {
                backData = body.push(chunck);
            });
            res.on('end', function () {
                body = Buffer.concat(body).toString();
                if (res.headers['content-type'].indexOf('text/html') !== -1) {
                    resolve(cheerio.load(body));
                }
                else if (res.headers['content-type'].indexOf('application/json') !== -1) {
                    resolve(JSON.parse(body));
                }
                else {
                    resolve(body);
                }
            });
        }).on('error', function (e) {
            reject(e)
        });
    });
    return pm;
}
loadPage(url + page).then(function ($) {
    setToatalPage($);
}).then(function () {
    let promiseList = [];
    for (page = 1; page <= 20; page++) {
        promiseList.push(
            new Promise(function (resolve, reject) {
                loadPage(url + page).then(function ($) {
                    Promise.all(grapData($)).then(function () {
                        console.log('All articles data in one page is ready ');
                        resolve(true);
                    });
                });
            })
        );
    }
    Promise.all(promiseList).then(function () {
        saveToExcel();
    });
});
// saveToExcel()
// 获取总页数
function setToatalPage($) {
    let pageTags = $('.desNavigation.cf a');
    totalPage = Number(pageTags.eq(pageTags.length - 2).text());
    totalPage = totalPage === 0 ? 1 : totalPage;
    console.log('totalPage:', totalPage)
}
// 获取每一篇文章的链接
function getArticleUrlList($) {
    let list = $('.result .youji-ul a.pic');
    let articleUrlList = [];
    list.each(function (i, ele) {
        articleUrlList.push(context + $(ele).attr('href'))
    });
    return articleUrlList;
}
// 获取每一篇文章的数据
function grapData($) {
    let articleUrlList = getArticleUrlList($);
    let promiseList = [];
    for (let i = 0; i < articleUrlList.length; i++) {
        promiseList.push(
            new Promise(function (resolve, reject) {
                loadPage(articleUrlList[i]).then(function ($) {
                    // console.log(d)
                    let model = {
                        title: $('.ctd_head h2').text().trim(),
                        content: $('.ctd_content').text(),
                        like: 0,
                        view: 0,
                        comments: 0,
                        link: articleUrlList[i]
                    }
                    let articleId = articleUrlList[i].substring(articleUrlList[i].lastIndexOf('/') + 1, articleUrlList[i].indexOf('.html'));
                    let GetBusinessDataUrl = `http://you.ctrip.com/TravelSite/Home/GetBusinessData?random=${Math.random()}
                    &arList[0].RetCode=0
                    &arList[0].Html=${articleId}
                    &arList[1].RetCode=1
                    &arList[1].Html=${articleId}`
                    loadPage(GetBusinessDataUrl).then(function (json) {
                        let data = JSON.parse(json[1]['Html'])['Html'];
                        model.like = data.LikeCount;
                        model.view = data.VisitCount;
                        model.comments = data.CommentCount;
                        excelList.push([model.title, model.content, model.like, model.view, model.comments, model.link]);
                        console.log('excelList:', excelList.length);
                        resolve(true);
                    });
                });
            }))
    }
    return promiseList;
}
// 保存到excel
function saveToExcel() {
    // var buffer = xlsx.build([{ name: "mySheetName", data: excelList.unshift(['title', 'content', 'like', 'view', 'comments', 'link']) }]);
    var buffer = xlsx.build([{ name: "mySheetName", data: excelList}]);
    app.get('/Excel', function (req, res) {
        res.setHeader('Content-Type', 'application/vnd.openxmlformats');
        res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
        res.end(buffer, 'binary');
    });
    app.listen(3000);
    console.log('Listening on port 3000');
    opn('http://localhost:3000/Excel')

}
