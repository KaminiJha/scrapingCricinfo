const cheerio = require("cheerio");
// const writeXlsxFile = require('write-excel-file/node');
const excel=require('excel4node');
const request = require('request');
const fs = require('fs');
const path = require('path');

const matchId = "ipl-2020-21-1210595";
request(`https://www.espncricinfo.com/series/${matchId}/match-results`, fetchScoreCard);

let scoreCardUrl = []

function fetchScoreCard(err, res, html) {
    const $ = cheerio.load(html);
    const fetchedScoredCardUrl = $(`[data-hover=Scorecard]`);
    for (let i = 0; i < fetchedScoredCardUrl.length; i++) {
        let url = "https://www.espncricinfo.com" + $(fetchedScoredCardUrl[i]).attr('href');
        scoreCardUrl[i] = url;
    }
    for (let i = 0; i < scoreCardUrl.length; i++) {

        request(scoreCardUrl[i], fetchContentOfMatch.bind(this, scoreCardUrl[i]));
    }
}

function fetchContentOfMatch(url,err, res, html) {
    //extracting match id 
    const urlSplit = url.split('/');
    const matchCompleteName = urlSplit[urlSplit.length - 2].split('-')
    const matchId = matchCompleteName[matchCompleteName.length - 1]
    //create new work book
    let wb=new excel.Workbook();

    let filename=path.join(`workbook`,matchId+".xlsx");
    //cheerio loading html
    const $ = cheerio.load(html);
    //innings fetching table
    const innings = $("div.Collapsible");
    let obj = [];
    for (let i = 0; i < innings.length; i++) {
        //creating worksheet
        let wsbatsman=wb.addWorksheet(`batsman-inning-${i+1}`);
        //       adding header
        wsbatsman.cell(1,1).string("Player name");
        wsbatsman.cell(1,2).string("runs");
        wsbatsman.cell(1,3).string("balls");
        wsbatsman.cell(1,4).string("fours");
        wsbatsman.cell(1,5).string("sixs");
        wsbatsman.cell(1,6).string("sr");
        //batsman table
        let batsmanRows = $(innings[i]).find('.table.batsman tbody tr')
        for (let j = 0; j < batsmanRows.length; j++) {
            let tds = $(batsmanRows[j]).find('td');
            if (tds.length == 8) {
              
                wsbatsman.cell(j+2,1).string($(tds[0]).text())
                wsbatsman.cell(j+2,2).string($(tds[2]).text())
                wsbatsman.cell(j+2,3).string($(tds[3]).text())
                wsbatsman.cell(j+2,4).string($(tds[5]).text())
                wsbatsman.cell(j+2,5).string($(tds[6]).text())
                wsbatsman.cell(j+2,6).string($(tds[7]).text())
                // let playerName = $(tds[0]).text();
                // let runs = $(tds[2]).text();
                // let balls = $(tds[3]).text();
                // let fours = $(tds[5]).text();
                // let sixs = $(tds[6]).text();
                // let sr = s = $(tds[7]).text();
                // obj.push({
                //     playerName,
                //     runs,
                //     balls,
                //     fours,
                //     sixs,
                //     sr,
                // })
            }

        }

        //bowling table
        let wsbowlers=wb.addWorksheet(`bowler-inning-${i+1}`);
        //headers
        wsbowlers.cell(1,1).string("Player name");
        wsbowlers.cell(1,2).string("o");
        wsbowlers.cell(1,3).string("m");
        wsbowlers.cell(1,4).string("r");
        wsbowlers.cell(1,5).string("w");
        wsbowlers.cell(1,6).string("econ");
        wsbowlers.cell(1,7).string("wd");
        wsbowlers.cell(1,8).string("nb");

        let bowlingRows = $(innings[i]).find('.table.bowler tbody tr');
        for (let j = 0; j < bowlingRows.length; j++) {
            let tds = $(bowlingRows[j]).find('td');
            if (tds.length == 11) {
                wsbowlers.cell(j+2,1).string($(tds[0]).text());
                wsbowlers.cell(j+2,2).string($(tds[1]).text());
                wsbowlers.cell(j+2,3).string($(tds[2]).text());
                wsbowlers.cell(j+2,4).string($(tds[3]).text());
                wsbowlers.cell(j+2,5).string($(tds[4]).text());
                wsbowlers.cell(j+2,6).string($(tds[5]).text());
                wsbowlers.cell(j+2,7).string($(tds[9]).text());
                wsbowlers.cell(j+2,8).string($(tds[10]).text());


                // let playerName = $(tds[0]).text();
                // let o = $(tds[1]).text();
                // let m = $(tds[2]).text()
                // let r = $(tds[3]).text()
                // let w = $(tds[4]).text();
                // let econ = $(tds[5]).text();
                // let wd = $(tds[9]).text();
                // let nb = $(tds[10]).text();
                // obj.push({
                //     playerName,
                //     o,
                //     m,
                //     r,
                //     w,
                //     econ,
                //     wd,
                //     nb
                // })
            }
        }
    }
    // fs.writeFileSync(filename,);
    fs.writeFileSync(filename,"")
    wb.write(filename)
    // fs.writeFileSync(`abc${c++}.json`, JSON.stringify(obj))
}
