const http = require("http");
const cheerio = require('cheerio');
const nodeExcel = require('excel-export');
let $;
let page = 1;
const url = `http://you.ctrip.com/searchsite/travels/?query=%E5%BE%B7%E5%9B%BD%E4%B9%A1%E6%9D%91&isAnswered=&isRecommended=&publishDate=&PageNo=`
const context = 'http://you.ctrip.com';
let totalPAge = 0;
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
    for (page = 1; page <= 1; page++) {
        promiseList.push(
            new Promise(function (resolve, reject) {
                loadPage(url + page).then(function ($) {
                    Promise.all(grapData($)).then(function () {
                        console.log('All article data in one page is ready ');
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
// 获取总页数
function setToatalPage($) {
    let pageTags = $('.desNavigation.cf a');
    totalPAge = Number(pageTags.eq(pageTags.length - 2).text());
    console.log(totalPAge)
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
                        excelList.push([model.title, model.content, model.like, model.view, model.comments]);
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
    let conf = {};
    conf.stylesXmlFile = "styles.xml";
    conf.name = "mysheet";
    conf.cols = [{
        caption: 'title',
        type: 'string',
        width: 28.7109375
    }, {
        caption: 'content',
        type: 'string',
        width: 300
    }, {
        caption: 'like',
        type: 'number¸'
    }, {
        caption: 'view',
        type: 'number'
    }, {
        caption: 'comments',
        type: 'number'
    }];
    console.log('excelList:', excelList.length);
    conf.rows = excelList;
    // const result = nodeExcel.execute(conf);
}