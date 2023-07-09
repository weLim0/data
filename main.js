import { createWriteStream as _createWriteStream, existsSync, mkdir, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { PDFDocument, rgb } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit'
import express from 'express';
import multer from 'multer';
import cors from 'cors';
import expressSession from "express-session";
import axios from 'axios';
import { makeZip } from './makeZip.js';
import { read, utils } from 'xlsx'
import jsontoxlsx from 'jsonrawtoxlsx'
let ck;
function addCommasToNumber(number) {
    // 숫자를 문자열로 변환하고, 천 단위마다 쉼표를 추가합니다.
    const strNumber = String(number);
    const parts = strNumber.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');

    // 소수점 아래가 있으면 합쳐줍니다.
    const formattedNumber = parts.join('.');
    return formattedNumber;
}
function splitText(text, n) {
    const sentences = [];
    let remainingText = text;

    while (remainingText.length > 0) {
        let sentence = '';
        let index = remainingText.indexOf('\n');

        if (index !== -1 && index <= n) {
            sentence = remainingText.slice(0, index + 1);
            remainingText = remainingText.slice(index + 1);
        } else if (remainingText.length > n) {
            sentence = remainingText.slice(0, n);
            remainingText = remainingText.slice(n);
        } else {
            sentence = remainingText;
            remainingText = '';
        }

        sentences.push(sentence);
    }

    return sentences.map((e) => e.replace(" \n", "").replace("\n", "").replace("\n ", "").trim()).join('\n').slice(0, 789);
}

function 대출(업태,매출) {
    if(업태.includes("도매") || 업태.includes("소매")) {
        return 매출*0.35
    }
    if(업태.includes("제조") || 업태.includes("정보")) {
        return 매출*0.5
    }
    if(업태.includes("서비스")) {
        return 매출*0.25
    }
    if(업태.includes("건설")) {
        return 매출*0.2
    }
}

function 대출2(업태,매출) {
    if(업태.includes("도매") || 업태.includes("소매")) {
        return 매출*0.30
    }
    if(업태.includes("제조") || 업태.includes("정보")) {
        return 매출*0.4
    }
    if(업태.includes("서비스")) {
        return 매출*0.25
    }
    if(업태.includes("건설")) {
        return 매출*0.2
    }
}

function 대출3(업태,매출) {
    if(업태.includes("도매") || 업태.includes("소매")) {
        return 매출*0.40
    }
    if(업태.includes("제조") || 업태.includes("정보")) {
        return 매출*0.65
    }
    if(업태.includes("서비스")) {
        return 매출*0.35
    }
    if(업태.includes("건설")) {
        return 매출*0.35
    }
}

async function modifyPdf(resJson, id) {
    let 예상대출 = Math.floor(대출(resJson.업태,resJson.매출액)) > 1000000000 ? 1000000000 : Math.floor(대출(resJson.업태,resJson.매출액))
    const pdfDoc = await PDFDocument.load(readFileSync('./보고서-빈칸.pdf').buffer)
    pdfDoc.registerFontkit(fontkit);
    const light = await pdfDoc.embedFont(readFileSync("./fonts/NanumGothicLight.ttf"))
    const bold = await pdfDoc.embedFont(readFileSync("./fonts/NanumGothicBold.ttf"))

    const pages = pdfDoc.getPages()
    const airesult1 = pages[1]
    const airesult2 = pages[2]
    const createlab = pages[3]
    const airesult3 = pages[5]
    const final = pages[7]
    const { width, height } = airesult1.getSize()

    //1p 왼쪽 표 x좌표 공식 x: width / 2 - bold.widthOfTextAtSize(addCommasToNumber(resJson.당기순이익), 24) - 105
    //1p 왼쪽 표 y좌표 공식 y: height / 2 + 122 - (60.5 * n)
    //기업명
    airesult1.drawText(resJson.기업명, {
        x: width / 2 - bold.widthOfTextAtSize(resJson.기업명, 48) - 45 - 8 - 9 - 9 - 4 - 4,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //대표자
    airesult1.drawText(resJson.대표자, {
        x: width / 2 - bold.widthOfTextAtSize(resJson.대표자, 48) - 45 - 8 - 9 - 9 - 4 - 4,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //사업자등록번호
    airesult1.drawText(resJson.사업자등록번호, {
        x: width / 2 - bold.widthOfTextAtSize(resJson.사업자등록번호, 48) - 45 - 8 - 9 - 9 - 4 - 4,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //업태
    airesult1.drawText(resJson.업태, {
        x: width / 2 - bold.widthOfTextAtSize(resJson.업태, 48) - 45 - 8 - 9 - 9 - 4 - 4,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125 - 125 + 10,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //사업연도
    airesult1.drawText(resJson.사업연도, {
        x: width / 2 - bold.widthOfTextAtSize(resJson.사업연도, 48) - 45 - 8 - 9 - 9 - 4 - 4,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125 - 125 - 125 + 20,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //매출액
    airesult1.drawText(addCommasToNumber(resJson.매출액), {
        x: width / 2 - bold.widthOfTextAtSize(addCommasToNumber(resJson.매출액), 48) - 45 - 8 - 9 - 9 - 4 - 4 + 1220,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //당기순이익
    airesult1.drawText(addCommasToNumber(resJson.당기순이익), {
        x: width / 2 - bold.widthOfTextAtSize(addCommasToNumber(resJson.당기순이익), 48) - 45 - 8 - 9 - 9 - 4 - 4 + 1220,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //추정세금
    airesult1.drawText(addCommasToNumber(resJson.세금), {
        x: width / 2 - bold.widthOfTextAtSize(addCommasToNumber(resJson.세금), 48) - 45 - 8 - 9 - 9 - 4 - 4 + 1220,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //자본금
    airesult1.drawText(addCommasToNumber(resJson.자본총액), {
        x: width / 2 - bold.widthOfTextAtSize(addCommasToNumber(resJson.자본총액), 48) - 45 - 8 - 9 - 9 - 4 - 4 + 1220,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125 - 125 + 10,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //이익잉여금
    airesult1.drawText(addCommasToNumber(resJson.이익잉여금), {
        x: width / 2 - bold.widthOfTextAtSize(addCommasToNumber(resJson.이익잉여금), 48) - 45 - 8 - 9 - 9 - 4 - 4 + 1220,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125 - 125 - 125 + 20,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //비유동부채
    airesult1.drawText(addCommasToNumber(resJson.비유동부채), {
        x: width / 2 - bold.widthOfTextAtSize(addCommasToNumber(resJson.비유동부채), 48) - 45 - 8 - 9 - 9 - 4 - 4 + 1220,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125 - 125 - 125 - 125 +29,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //부채비율
    airesult1.drawText(resJson.부채비율 + "%", {
        x: width / 2 - bold.widthOfTextAtSize(resJson.부채비율 + "%", 48) - 45 - 8 - 9 - 9 - 4 - 4 + 1220,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125 - 125 - 125 - 125 + 27 - 125 + 25 - 10,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //차입금/매출
    airesult1.drawText(Number(resJson.장기차입금_매출액).toFixed(2) + "%", {
        x: width / 2 - bold.widthOfTextAtSize(Number(resJson.장기차입금_매출액).toFixed(2) + "%", 48) - 45 - 8 - 9 - 9 - 4 - 4 + 1220,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 125 - 125 - 125 - 125 - 125 + 27 - 125 - 100,
        size: 48,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //AI추천 컨설팅[1]
    let 컨설팅 = ''
    컨설팅 += `당사의 매출은 ${addCommasToNumber(resJson.매출액)}원이며 '비유동부채'는 ${addCommasToNumber(resJson.비유동부채)}원 이다.\n\n'비유동부채의 대출이' 운전자금이 아니라고 가정한 대출가능 금액을 추산해봤습니다.`
    airesult2.drawText(splitText(컨설팅, 52), {
        x: width / 2 - 30,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    let max대출 = addCommasToNumber(Math.floor(대출2(resJson.업태,resJson.매출액)) > 330000000 ? 330000000 : Math.floor(대출2(resJson.업태,resJson.매출액)))
    airesult2.drawText(`매출액 ${addCommasToNumber(resJson.매출액)}원의 예상 최대 정책자금(운전자금)은의 `, {
        x: width / 2 - 30,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 200 - 50,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult2.drawText(`최고한도는 ${addCommasToNumber(예상대출)}원이며, 1회 최대 대출 한도는 ${max대출}원`, {
        x: width / 2 - 30,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 200 - 50 - 33 - 5 - 5,
        size: 33,
        font: bold,
        color: rgb(1, 0, 0)
    })
    airesult2.drawText("이다.", {
        x: width / 2 - 30 + bold.widthOfTextAtSize(`최고한도는 ${addCommasToNumber(예상대출)}원이며, 1회 최대 대출 한도는 ${max대출}원`, 33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 200 - 50 - 33 - 5 - 5,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult2.drawText(`연구소, 각종 인증 등을 할 경우 상승시킬 수 있는 최대 정책자금 한도는`, {
        x: width / 2 - 30,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 200 - 50 - 33 - 5 - 5 - 33 - 5 - 5,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult2.drawText(addCommasToNumber(Math.floor(대출3(resJson.업태,resJson.매출액)))+"원이다.", {
        x: width / 2 - 30 ,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 200 - 50 - 33 - 5 - 5 - 33 - 5 - 5 - 33 - 5 - 5,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult2.drawText("1회 최대 대출 한도 금액은 "+addCommasToNumber(Math.floor(Number(max대출.replaceAll(",",""))*1.25))+"원이다.", {
        x: width / 2 - 30 ,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 200 - 50 - 33 - 5 - 5 - 33 - 5 - 5 - 33 - 5 - 5 - 33 - 5 - 5,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    //연구소설립효과
    let 절약
    let createlaboratory = ''
    createlaboratory += `당사의 당기순이익 ${addCommasToNumber(resJson.당기순이익)}원의 추정 납부 세금은`
    createlab.drawText(createlaboratory, {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    createlab.drawText(`${addCommasToNumber(resJson.세금)}원이라고 가정한다.`, {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    createlab.drawText('- 연구소 설립 후 연구요원 입명시 ', {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 33 - 5,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    createlab.drawText('인건비 총액의 25%가 세액 공제', {
        x: width / 2 + 45 + bold.widthOfTextAtSize('- 연구소 설립 후 연구요원 입명시 ',33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 33 - 5,
        size: 33,
        font: bold,
        color: rgb(1, 0, 0)
    })
    createlab.drawText('된다', {
        x: width / 2 + 45 + bold.widthOfTextAtSize('- 연구소 설립 후 연구요원 입명시 '+'인건비 총액의 25%가 세액 공제',33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 33 - 5,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    if (Number(resJson.세금) - 17500000 < Math.floor(resJson.당기순이익 * 0.07)) {
        if (Number(resJson.세금) - 17500000 < 0) {
            let createlaboratorySim = `'1명'의 연구원 임금을 각 1년 '35,000,000원으로 책정했을 경우\n\n35,000,000원 x 2명 x 25% = 17,500,000원의 세액이 공제 된다.\n\n당사의 최종 납부 세금은 ${addCommasToNumber(resJson.세금)}원 - 17,500,000원 = ${addCommasToNumber(Number(resJson.세금) - 17500000 < 0 ? 0 : Number(resJson.세금) - 17500000)}원${Number(resJson.세금) - 17500000 < 0 ? `\n\n(${addCommasToNumber(Number(resJson.세금) - 17500000)}원 이월가능)` : ""}이다.\n\n하지만, 최저한세로 인해 ${addCommasToNumber(resJson.당기순이익)}원 x 7% = `
            createlab.drawText(splitText(createlaboratorySim, 51), {
                x: width / 2 + 45,
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250,
                size: 33,
                font: bold,
                color: rgb(0, 0, 0)
            })

            createlab.drawText(`${addCommasToNumber(Math.floor(resJson.당기순이익 * 0.07))}원을`, {
                x: width / 2 + 45 + bold.widthOfTextAtSize(`하지만, 최저한세로 인해 ${addCommasToNumber(resJson.당기순이익)}원 x 7% = `, 33),
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*5),
                size: 33,
                font: bold,
                color: rgb(1, 0, 0)
            })

            createlab.drawText(`최종납부`, {
                x: width / 2 + 45,
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (40*6) + 3.5,
                size: 33,
                font: bold,
                color: rgb(1, 0, 0)
            })

            createlab.drawText(`하게된다.`, {
                x: width / 2 + 45 + bold.widthOfTextAtSize('최종납부', 33),
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (40*6) + 3.5,
                size: 33,
                font: bold,
                color: rgb(0, 0, 0)
            })
            절약 = Math.floor(resJson.당기순이익 * 0.07)
            createlab.drawText(`즉, 당초 ${addCommasToNumber(resJson.세금)}원 - ${addCommasToNumber(절약)}원 = `, {
                x: width / 2 + 45,
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (40*7),
                size: 33,
                font: bold,
                color: rgb(0, 0, 0)
            })

            createlab.drawText(`${addCommasToNumber((resJson.세금 - 절약))}원을 절약 가능`, {
                x: width / 2 + 45 + bold.widthOfTextAtSize(`즉, 당초 ${addCommasToNumber(resJson.세금)}원 - ${addCommasToNumber(절약)}원 =  `, 33),
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (40*7),
                size: 33,
                font: bold,
                color: rgb(1, 0, 0)
            })
            createlab.drawText(`하다.`, {
                x: width / 2 + 45,
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (40*8),
                size: 33,
                font: bold,
                color: rgb(0, 0, 0)
            })
        } else {
            let createlaboratorySim = `'1명'의 연구원 임금을 각 1년 '35,000,000원으로 책정했을 경우\n\n35,000,000원 x 2명 x 25% = 17,500,000원의 세액이 공제 된다.\n\n당사의 최종 납부 세금은 \n\n${addCommasToNumber(resJson.세금)}원 - 17,500,000원 = ${addCommasToNumber(Number(resJson.세금) - 17500000 < 0 ? 0 : Number(resJson.세금) - 17500000)}원${Number(resJson.세금) - 17500000 < 0 ? `\n\n(${addCommasToNumber(Number(resJson.세금) - 17500000)}원 이월가능)` : ""}이다.\n\n하지만, 최저한세로 인해 ${addCommasToNumber(resJson.당기순이익)}원 x 7% = `
            createlab.drawText(splitText(createlaboratorySim, 51), {
                x: width / 2 + 45,
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250,
                size: 33,
                font: bold,
                color: rgb(0, 0, 0)
            })

            createlab.drawText(`${addCommasToNumber(Math.floor(resJson.당기순이익 * 0.07))}원을`, {
                x: width / 2 + 45 + bold.widthOfTextAtSize(`하지만, 최저한세로 인해 ${addCommasToNumber(resJson.당기순이익)}원 x 7% = `, 33),
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*4),
                size: 33,
                font: bold,
                color: rgb(1, 0, 0)
            })

            createlab.drawText(`최종납부`, {
                x: width / 2 + 45,
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*4),
                size: 33,
                font: bold,
                color: rgb(1, 0, 0)
            })

            createlab.drawText(`하게된다.`, {
                x: width / 2 + 45 + bold.widthOfTextAtSize('최종납부', 33),
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*4),
                size: 33,
                font: bold,
                color: rgb(0, 0, 0)
            })
            절약 = Math.floor(resJson.당기순이익 * 0.07)
            createlab.drawText(`즉, 당초 ${addCommasToNumber(resJson.세금)}원 - ${addCommasToNumber(절약)}원 = `, {
                x: width / 2 + 45,
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*7),
                size: 33,
                font: bold,
                color: rgb(0, 0, 0)
            })

            createlab.drawText(` ${addCommasToNumber((resJson.세금 - 절약))}원을 절약 가능`, {
                x: width / 2 + 45 + bold.widthOfTextAtSize(`즉, 당초 ${addCommasToNumber(resJson.세금)}원 - ${addCommasToNumber(절약)}원 =  `, 33),
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*7),
                size: 33,
                font: bold,
                color: rgb(1, 0, 0)
            })
            createlab.drawText(`하다.`, {
                x: width / 2 + 45,
                y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*8),
                size: 33,
                font: bold,
                color: rgb(0, 0, 0)
            })
        }


    } else {
        let createlaboratorySim = `'1명'의 연구원 임금을 각 1년 '35,000,000원으로 책정했을 경우\n\n35,000,000원 x 2명 x 25% = 17,500,000원의 세액이 공제 된다.`
        createlab.drawText(splitText(createlaboratorySim, 51), {
            x: width / 2 + 45,
            y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250,
            size: 33,
            font: bold,
            color: rgb(0, 0, 0)
        })

        createlab.drawText("당사의 최종 납부 세금은", {
            x: width / 2 + 45,
            y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*3),
            size: 33,
            font: bold,
            color: rgb(1, 0, 0)
        })

        createlab.drawText(`${addCommasToNumber(resJson.세금)}원 - 17,500,000원 = ${addCommasToNumber(Number(resJson.세금) - 17500000 < 0 ? 0 : Number(resJson.세금) - 17500000)}원${Number(resJson.세금) - 17500000 < 0 ? `\n(${addCommasToNumber(Number(resJson.세금) - 17500000)}원 이월가능)` : ""}이다.`, {
            x: width / 2 + 45,
            y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*4),
            size: 33,
            font: bold,
            color: rgb(1, 0, 0)
        })
        절약 = Number(resJson.세금) - 17500000 < 0 ? 0 : Number(resJson.세금) - 17500000
        createlab.drawText(`즉, 당초 ${addCommasToNumber(resJson.세금)}원 - ${addCommasToNumber(절약)}원 = `, {
            x: width / 2 + 45,
            y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*5),
            size: 33,
            font: bold,
            color: rgb(0, 0, 0)
        })

        createlab.drawText(` ${addCommasToNumber((resJson.세금 - 절약))}원을 절약 가능`, {
            x: width / 2 + 45 + light.widthOfTextAtSize(`즉, 당초 ${addCommasToNumber(resJson.세금)}원 - ${addCommasToNumber(절약)}원 =  `, 33),
            y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*5),
            size: 33,
            font: bold,
            color: rgb(1, 0, 0)
        })
        createlab.drawText(`하다.`, {
            x: width / 2 + 45,
            y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 33 - 5 - 5 - 250 - (38*6),
            size: 33,
            font: bold,
            color: rgb(0, 0, 0)
        })
    }

    //AI 추천 컨설팅 [3]
    let 벤처 = `당사의 추정 납부 세금은 ${addCommasToNumber(resJson.세금)}원이라고 가정한다.`
    airesult3.drawText(벤처, {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult3.drawText(`${addCommasToNumber(resJson.세금)}원 x 50% = `, {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 38,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult3.drawText(`${addCommasToNumber(Math.floor(resJson.세금 * 0.5))}원으로 감면`, {
        x: width / 2 + 45 + bold.widthOfTextAtSize(`${addCommasToNumber(resJson.세금)}원 x 50% = `, 33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 38,
        size: 33,
        font: bold,
        color: rgb(1, 0, 0)
    })
    airesult3.drawText(`된다.`, {
        x: width / 2 + 45 + bold.widthOfTextAtSize(`${addCommasToNumber(resJson.세금)}원 x 50% = ` + `${addCommasToNumber(Math.floor(resJson.세금 * 0.5))}원으로 감면`, 33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 38,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })

    let 대출한도 = (Math.floor(resJson.매출액 * 0.4) > 500000000 ? 500000000 : Math.floor(resJson.매출액 * 0.4))
    let 메인비즈 = `당사의 1회 대출 최고 한도를 ${addCommasToNumber(Math.floor(resJson.매출액 * 0.4) > 500000000 ? 500000000 : Math.floor(resJson.매출액 * 0.4))}원이라고 가정한다`
    airesult3.drawText(메인비즈, {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 210,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult3.drawText(`${addCommasToNumber(Math.floor(resJson.매출액 * 0.4) > 500000000 ? 500000000 : Math.floor(resJson.매출액 * 0.4))}원 X 125% = `, {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 210 - 38,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult3.drawText(`${addCommasToNumber(Math.floor(대출한도 * 1.25))}원으로 증가`, {
        x: width / 2 + 45 + bold.widthOfTextAtSize(`${addCommasToNumber(Math.floor(resJson.매출액 * 0.4) > 500000000 ? 500000000 : Math.floor(resJson.매출액 * 0.4))}원 X 125% = `, 33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 210 - 38,
        size: 33,
        font: bold,
        color: rgb(1, 0, 0)
    })
    airesult3.drawText(`된다.`, {
        x: width / 2 + 45 + bold.widthOfTextAtSize(`${addCommasToNumber(Math.floor(resJson.매출액 * 0.4) > 500000000 ? 500000000 : Math.floor(resJson.매출액 * 0.4))}원 X 125% = ` + `${addCommasToNumber(Math.floor(대출한도 * 1.25))}원으로 증가`, 33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 210 - 38,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })

    let 이노비즈 = `당사의 1회 대출 최고 한도를 ${addCommasToNumber(Math.floor(대출한도 * 1.25))}원으로 가정한다.`
    airesult3.drawText(이노비즈, {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 210 - 210,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult3.drawText(`${addCommasToNumber(Math.floor(대출한도 * 1.25))}원 X 125% = `, {
        x: width / 2 + 45,
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 210 - 210 - 38,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })
    airesult3.drawText(`${addCommasToNumber(Math.floor(Math.floor(대출한도 * 1.25) * 1.25))}원으로 증가`, {
        x: width / 2 + 45 + bold.widthOfTextAtSize(`${addCommasToNumber(Math.floor(대출한도 * 1.25))}원 X 125% = `, 33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 210 - 210 - 38,
        size: 33,
        font: bold,
        color: rgb(1, 0, 0)
    })
    airesult3.drawText(`된다.`, {
        x: width / 2 + 55 + bold.widthOfTextAtSize(`${addCommasToNumber(Math.floor(대출한도 * 1.25))}원 X 125% = ` + `${addCommasToNumber(Math.floor(Math.floor(대출한도 * 1.25) * 1.25))}원으로 증가`, 33),
        y: height / 2 + 112 + 100 + 15 + 60 + 100 - 40 - 5 - 70 - 20 - 5 - 5 - 210 - 210 - 38,
        size: 33,
        font: bold,
        color: rgb(0, 0, 0)
    })


    //최종혜택
    //정책자금 한도 - 당초
    final.drawText(addCommasToNumber(예상대출), {
        x: width / 2 - (bold.widthOfTextAtSize(addCommasToNumber(예상대출), 44) / 2) - 544,
        y: height / 2 + 200,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //정책자금 한도 - 연구소설립
    final.drawText(addCommasToNumber(Math.floor(예상대출 * 1.2)), {
        x: width / 2 - (bold.widthOfTextAtSize(addCommasToNumber(Math.floor(예상대출 * 1.2)), 44) / 2) - 544 + 390,
        y: height / 2 + 200,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //정책자금 한도 - 기업인증
    final.drawText(addCommasToNumber(Math.floor(예상대출 * 1.25)), {
        x: width / 2 - (bold.widthOfTextAtSize(addCommasToNumber(Math.floor(예상대출 * 1.25)), 44) / 2) - 544 + 390 + 380,
        y: height / 2 + 200,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //정책자금 한도 - 혜택합계
    final.drawText(addCommasToNumber((Math.floor(예상대출 * 1.2) - 예상대출) + (Math.floor(예상대출 * 1.25) - 예상대출)), {
        x: width / 2 - (bold.widthOfTextAtSize(addCommasToNumber((Math.floor(예상대출 * 1.2) - 예상대출) + (Math.floor(예상대출 * 1.25) - 예상대출)), 44) / 2) - 544 + 390 + 390 + 390,
        y: height / 2 + 200,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //정책자금 이자 - 혜택합계
    final.drawText(addCommasToNumber(Math.floor(예상대출 * 0.02)), {
        x: width / 2- (bold.widthOfTextAtSize(addCommasToNumber(Math.floor(예상대출 * 0.02)), 44) / 2) - 544 + 390 + 390 + 390,
        y: height / 2 + 70,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })


    //세금혜택 - 당초
    final.drawText(addCommasToNumber(resJson.세금), {
        x: width / 2 - (bold.widthOfTextAtSize(addCommasToNumber(resJson.세금), 44) / 2) - 544 -5,
        y: height / 2 - 70 + 5,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })
    let 이월 = Number(resJson.세금) - 17500000 < 0 ? Number(resJson.세금) - 17500000 : 0

    //세금혜택 - 연구소설립
    final.drawText(addCommasToNumber((resJson.세금 - 절약) + Math.abs(이월)), {
        x: width / 2 + 11.5 - (bold.widthOfTextAtSize(addCommasToNumber((resJson.세금 - 절약) + Math.abs(이월)), 44) / 2) - 544 + 390 -5,
        y: height / 2 - 70 + 5,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //세금혜택 - 혜택합계
    final.drawText(addCommasToNumber((((resJson.세금 - 절약) + Math.abs(이월)) + Math.floor(resJson.세금 * 0.5))), {
        x: width / 2 + 11.5 - (bold.widthOfTextAtSize(addCommasToNumber((((resJson.세금 - 절약) + Math.abs(이월)) + Math.floor(resJson.세금 * 0.5))), 44) / 2) - 544 + 390 + 390 + 390 -5,
        y: height / 2 - 70 + 5,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })

    //세금혜택 - 기업인증
    final.drawText(addCommasToNumber(Math.floor(resJson.세금 * 0.5)), {
        x: width / 2 + 11.5 - (bold.widthOfTextAtSize(addCommasToNumber(Math.floor(resJson.세금 * 0.5)), 44) / 2) - 544 + 390 + 380 -5,
        y: height / 2 - 70 + 5,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })
    //`${addCommasToNumber((resJson.세금 - Math.floor(resJson.당기순이익 * 0.07)))}원 감소되며, 내년에 ${addCommasToNumber((resJson.세금 - Math.floor(resJson.당기순이익 * 0.07)) + Math.abs(이월))}원 추가 감면됩니다.`

    //정책자금 한도
    final.drawText(`최소 ${addCommasToNumber((Math.floor(예상대출 * 1.2) - 예상대출) + (Math.floor(예상대출 * 1.25) - 예상대출))}원 증가 `, {
        x: width / 2 - 845 + 7.5 + 55.5,
        y: height / 2 - 264 - 46.5,
        size: 44,
        font: bold,
        color: rgb(1, 0, 0)
    })
    final.drawText(`됩니다.`, {
        x: width / 2 - 845 + 7.5 + 55.5 + bold.widthOfTextAtSize(`최소 ${addCommasToNumber((Math.floor(예상대출 * 1.2) - 예상대출) + (Math.floor(예상대출 * 1.25) - 예상대출))}원 증가 `, 44),
        y: height / 2 - 264 - 46.5,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })


    //정책자금이자
    final.drawText(`최소 ${addCommasToNumber(Math.floor(예상대출 * 0.02))}원 감소`, {
        x: width / 2 - 845 + 7.5 + 55.5,
        y: height / 2 - 264 - 46.5 - 77,
        size: 44,
        font: bold,
        color: rgb(1, 0, 0)
    })

    final.drawText(`됩니다.`, {
        x: width / 2 + - 845 + 7.5 + 55.5 + bold.widthOfTextAtSize(`최소 ${addCommasToNumber(Math.floor(예상대출 * 0.02))}원 감소`, 44),
        y: height / 2 - 264 - 46.5 - 77,
        size: 44,
        font: bold,
        color: rgb(0, 0, 0)
    })


    //세금납부
    if (Number(resJson.세금) - 17500000 < 0) {
        let rd = Number(resJson.세금) - 17500000 < Math.floor(resJson.당기순이익 * 0.07) ? Math.floor(resJson.당기순이익 * 0.07) : resJson.세금 - 절약
        final.drawText(`${addCommasToNumber((rd))}원 감소되며, 내년에 ${addCommasToNumber((resJson.세금 - 절약) + Math.abs(이월))}원 추가 감면`, {
            x: width / 2 - 845 + 7.5 + 55.5,
            y: height / 2 - 264 - 46.5 - 77 - 74,
            size: 44,
            font: bold,
            color: rgb(1, 0, 0)
        })

        final.drawText(`됩니다.`, {
            x: width / 2 - 845 + 7.5 + 55.5 + bold.widthOfTextAtSize(`${addCommasToNumber((resJson.세금 - 절약))}원 감소되며, 내년에 ${addCommasToNumber((resJson.세금 - 절약) + Math.abs(이월))}원 추가 감면`, 44),
            y: height / 2 -  264 - 46.5 - 77 - 74,
            size: 44,
            font: bold,
            color: rgb(0, 0, 0)
        })
    } else {
        let rd = Number(resJson.세금) - 17500000 < Math.floor(resJson.당기순이익 * 0.07) ? Math.floor(resJson.당기순이익 * 0.07) : resJson.세금 - 절약
        final.drawText(`${addCommasToNumber(rd)}원 감소`, {
            x: width / 2 - 845 + 7.5 + 55.5,
            y: height / 2 - 264 - 46.5 - 77 - 74,
            size: 44,
            font: bold,
            color: rgb(1, 0, 0)
        })

        final.drawText(`됩니다.`, {
            x: width / 2 - 845 + 7.5 + 55.5 + bold.widthOfTextAtSize(`${addCommasToNumber((resJson.세금 - 절약))}원 감소`, 44),
            y: height / 2 -  264 - 46.5 - 77 - 74,
            size: 44,
            font: bold,
            color: rgb(0, 0, 0)
        })
    }
    const pdfBytes = await pdfDoc.save()
    function mkdir( dirPath ) {
        const isExists = existsSync( dirPath );
        if( !isExists ) {
            mkdirSync( dirPath, { recursive: true } );
        }
    }
    mkdir("data/pdfData/"+id)
    writeFileSync('data/pdfData/' + id + '/' + resJson.기업명 + "_" + resJson.대표자 + '.pdf', pdfBytes)
    return true
}

const getCookie = async () => {
    const url = 'https://new.cretop.com/httpService/request.json';

    const headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6,zh;q=0.5',
        'cache-control': 'no-cache',
        'content-type': 'application/json',
        'pragma': 'no-cache',
        'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'Referer': 'https://new.cretop.com/ET/FI/ETFI110M1?h=1688379752019',
        'Referrer-Policy': 'strict-origin-when-cross-origin'
    };

    const data = {
        "header": {
          "trxCd": "PLIL1401R",
          "sysCd": "",
          "chlType": "02",
          "userId": "",
          "screenId": "PLIS010M1",
          "menuId": "01W0000737",
          "langCd": "ko",
          "bzno": null,
          "conoPid": null,
          "kedcd": null,
          "indCd": null,
          "franMngNo": null,
          "ctrNo": null,
          "bzcCd": null,
          "infoOfrStpgeYn": null,
          "pageNum": 0,
          "pageCount": 0,
          "pndNo": null
        },
        "PLIL1401R": {
          "kipId": "LOV6005",
          "pwd": "sugar123!@#",
          "autoLoginYn": "N",
          "sgnVal": null,
          "webEdevId": "iRYXMbxSgoINre7rZ6WLsiliOMgibIzK",
          "appEdevId": "",
          "sessEdYn": "N"
        }
      }

    try {
        const response = await axios.default.post(url, data, { headers });
        return response.headers['set-cookie'][1]
    } catch (error) {
        console.error(error);
    }
}

async function getCert(id) {
    const options = {
        method: 'POST',
        url: 'https://new.cretop.com/httpService/request.json',
        headers: {
            accept: 'application/json, text/plain, */*',
            'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6,zh;q=0.5',
            'cache-control': 'no-cache',
            'content-type': 'application/json',
            pragma: 'no-cache',
            'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            cookie: ck,
            Referer: 'https://new.cretop.com/ET/GN/ETGN090M1?h=1688378684358',
            'Referrer-Policy': 'strict-origin-when-cross-origin',
        },
        data: {
            "header": {
                "trxCd": "ETBR0807R",
                "sysCd": "",
                "chlType": "02",
                "screenId": "ETBR080SH",
                "menuId": "01W0000719",
                "langCd": "ko",
                "bzno": null,
                "conoPid": null,
                "kedcd": id,
                "indCd": null,
                "franMngNo": null,
                "ctrNo": null,
                "bzcCd": null,
                "infoOfrStpgeYn": null,
                "pageNum": 0,
                "pageCount": 0,
                "pndNo": null
            },
            "ETBR0807R": {
                "kedcd": id
            }
        }
    };

    try {
        const response = await axios(options);
        let cerfs = response.data.SETBR08007OT.setbr080abot
        let res = {
            메인비즈: cerfs.mainbizHdYn == "N" ? "미인증" : "인증",
            벤처: cerfs.venpHdYn == "N" ? "미인증" : "인증",
            이노비즈: cerfs.innobizHdYn == "N" ? "미인증" : "인증",
            연구개발전담부서: cerfs.rsrchDvlExrsOrgHdYn == "N" ? "미인증" : "인증",
            부설연구소: cerfs.lapatenpHdYn == "N" ? "미인증" : "인증"
        }
        return res;
    } catch (error) {
        return {
            메인비즈: "미확인",
            벤처: "미확인",
            이노비즈: "미확인",
            연구개발전담부서: "미확인",
            부설연구소: "미확인"
        }
    }
}


async function getUpTae(id) {
    const options = {
        method: 'POST',
        url: 'https://new.cretop.com/httpService/request.json',
        headers: {
            accept: 'application/json, text/plain, */*',
            'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6,zh;q=0.5',
            'cache-control': 'no-cache',
            'content-type': 'application/json',
            pragma: 'no-cache',
            'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            cookie: ck,
            Referer: 'https://new.cretop.com/ET/GN/ETGN090M1?h=1688378684358',
            'Referrer-Policy': 'strict-origin-when-cross-origin',
        },
        data: {
            "header": {
              "trxCd": "ETBR080KR",
              "sysCd": "",
              "chlType": "02",
              "userId": "LOV6005",
              "screenId": "ETBR080SU",
              "menuId": "01W0000732",
              "langCd": "ko",
              "bzno": null,
              "conoPid": null,
              "kedcd": id,
              "indCd": null,
              "franMngNo": null,
              "ctrNo": null,
              "bzcCd": null,
              "infoOfrStpgeYn": null,
              "pageNum": 0,
              "pageCount": 0,
              "pndNo": null
            },
            "ETBR080KR": {
              "kedcd": id
            }
          }
    };

    try {
        const response = await axios(options);
        return response.data.SETBR08020OT.setbr080arot.bzcCdNm1;
    } catch (error) {
        return "X"
    }
}
function numberFormat(x) {
    return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
async function getEcoDatas(id) {
    const options = {
        method: 'POST',
        url: 'https://new.cretop.com/httpService/request.json',
        headers: {
            accept: 'application/json, text/plain, */*',
            'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6,zh;q=0.5',
            'cache-control': 'no-cache',
            'content-type': 'application/json',
            pragma: 'no-cache',
            'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            cookie: ck,
            Referer: 'https://new.cretop.com/ET/GN/ETGN090M1?h=1688378684358',
            'Referrer-Policy': 'strict-origin-when-cross-origin',
        },
        data: {
            "header": {
                "trxCd": "ETFI1122R",
                "sysCd": "",
                "chlType": "02",
                "screenId": "ETFI112S2",
                "menuId": "01W0000777",
                "langCd": "ko",
                "bzno": null,
                "conoPid": null,
                "kedcd": id,
                "indCd": null,
                "franMngNo": null,
                "ctrNo": null,
                "bzcCd": null,
                "infoOfrStpgeYn": null,
                "pageNum": 0,
                "pageCount": 0,
                "pndNo": null
            },
            "ETFI1122R": {
                "kedcd": id,
                "acctCcd": "Y",
                "acctDt": "20221231",
                "fsCcd": "1",
                "fsCls": "2",
                "chk": "1",
                "smryYn": "N",
                "srchCls": "3"
            }
        }
    };

    try {
        const response = await axios(options);
        return {"비유동부채":numberFormat(response.data.SETFI11202OT.setfi112acotList.find((e) => e.accNm == "   비유동부채(*)").val5 == null ? "X" : response.data.SETFI11202OT.setfi112acotList.find((e) => e.accNm == "   비유동부채(*)").val5),"장기차입금":numberFormat(response.data.SETFI11202OT.setfi112acotList.find((e) => e.accNm == "      장기차입금(*)").val5 == null ? "X" : response.data.SETFI11202OT.setfi112acotList.find((e) => e.accNm == "      장기차입금(*)").val5),"자본금":numberFormat(response.data.SETFI11202OT.setfi112acotList.find((e) => e.accNm == "자본(*)").val5 == null ? "X" : response.data.SETFI11202OT.setfi112acotList.find((e) => e.accNm == "자본(*)").val5),"이익잉여금":numberFormat(response.data.SETFI11202OT.setfi112acotList.find((e) => e.accNm == "   이익잉여금(*)").val5 == null ? "X" : response.data.SETFI11202OT.setfi112acotList.find((e) => e.accNm == "   이익잉여금(*)").val5)}
    } catch (error) {
        console.log(error)
        return false;
    }
}
function convertPhoneNumber(phoneNumber) {
    const prefix = phoneNumber.substring(0, 3);
    const remaining = phoneNumber.substring(3);

    let convertedPrefix = '';

    if (prefix === '011') {
        const areaCode = remaining.substring(1, 4);
        const exchangeCode = remaining.substring(4, 7);

        if (areaCode >= 200 && areaCode <= 499) {
            convertedPrefix = '5';
        } else if (areaCode >= 500 && areaCode <= 899) {
            convertedPrefix = '3';
        } else if (areaCode >= 1700 && areaCode <= 1799) {
            convertedPrefix = remaining.substring(0, 2);
        } else if (areaCode >= 9500 && areaCode <= 9999) {
            if (remaining.startsWith('9')) {
                convertedPrefix = '8';
            } else {
                convertedPrefix = '0';
            }
        } else if (areaCode >= 9000 && areaCode <= 9499) {
            convertedPrefix = '0';
        }
    } else if (prefix === '017') {
        const areaCode = remaining.substring(1, 4);

        if (areaCode >= 200 && areaCode <= 499) {
            convertedPrefix = '6';
        } else if (areaCode >= 500 && areaCode <= 899) {
            convertedPrefix = '4';
        }
    } else if (prefix === '016') {
        const areaCode = remaining.substring(1, 4);

        if (areaCode >= 200 && areaCode <= 499) {
            convertedPrefix = '3';
        } else if (areaCode >= 500 && areaCode <= 899) {
            convertedPrefix = '2';
        } else if (areaCode >= 9000 && areaCode <= 9499 && areaCode !== '710' && areaCode !== '719') {
            if (remaining.startsWith('9')) {
                convertedPrefix = '7';
            } else {
                convertedPrefix = '0';
            }
        } else if (areaCode >= 9500 && areaCode <= 9999) {
            convertedPrefix = '0';
        }
    } else if (prefix === '018') {
        const areaCode = remaining.substring(1, 4);

        if (areaCode >= 200 && areaCode <= 499) {
            convertedPrefix = '4';
        } else if (areaCode >= 500 && areaCode <= 899) {
            convertedPrefix = '6';
        }
    } else if (prefix === '019') {
        const areaCode = remaining.substring(1, 4);

        if (areaCode >= 200 && areaCode <= 499) {
            convertedPrefix = '2';
        } else if (areaCode >= 500 && areaCode <= 899) {
            convertedPrefix = '5';
        } else if (areaCode >= 9000 && areaCode <= 9499) {
            if (remaining.startsWith('9')) {
                convertedPrefix = '8';
            }
        } else if (areaCode >= 9500 && areaCode <= 9999) {
            if (remaining.startsWith('9')) {
                convertedPrefix = '7';
            }
        }
    }
    return `010-${convertedPrefix.replace('-', '')}${remaining.replace('-', '')}`;
}
async function getCompanyInfo(id) {
    const options = {
        method: 'POST',
        url: 'https://new.cretop.com/httpService/request.json',
        headers: {
            accept: 'application/json, text/plain, */*',
            'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6,zh;q=0.5',
            'cache-control': 'no-cache',
            'content-type': 'application/json',
            pragma: 'no-cache',
            'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            cookie: ck,
            Referer: 'https://new.cretop.com/ET/GN/ETGN090M1?h=1688378684358',
            'Referrer-Policy': 'strict-origin-when-cross-origin',
        },
        data: {
            "header": {
                "trxCd": "ETGN0911R",
                "sysCd": "",
                "chlType": "02",
                "screenId": "ETGN091SA",
                "menuId": "01W0000737",
                "langCd": "ko",
                "bzno": null,
                "conoPid": null,
                "kedcd": id,
                "indCd": null,
                "franMngNo": null,
                "ctrNo": null,
                "bzcCd": null,
                "infoOfrStpgeYn": null,
                "pageNum": 0,
                "pageCount": 0,
                "pndNo": null
            },
            "ETGN0911R": {
                "kedcd": id
            }
        },
    };

    try {
        const response = await axios(options);
        if (response.data.SETGN09101OT.setgn091aaot == null) return false
        let dump = response.data.SETGN09101OT
        if (dump.setgn091abotList.length == 0 || Number(dump.setgn091abotList[0].acctDt.slice(0, 4)) <= 2019) return false;
        let res = {
            매출액: dump.setgn091abotList[0].sam,
            영업이익: dump.setgn091abotList[0].bzpf,
            당기순이익: dump.setgn091abotList[0].npf,
            종업원수: dump.setgn091aaot.employee == null ? "X" : dump.setgn091aaot.employee + "명",
            이메일: dump.setgn091aaot.email == null ? "X" : dump.setgn091aaot.email,
            주소: dump.setgn091aaot.rdnmAddr
        }
        return res;
        // Handle the response data here
    } catch (error) {
        return {}
        // Handle the error here
    }
}
const regex = /<!.{2}>/g;

function removeTag(array, regex) {
    return array.map(obj => {
        const updatedObj = {};
        for (let key in obj) {
            if (typeof obj[key] === 'string') {
                updatedObj[key] = obj[key].replace(regex, '');
            } else {
                updatedObj[key] = obj[key];
            }
        }
        return updatedObj;
    });
}

async function getCompany(company, ceo, phone) {
    ck = await getCookie();
    let name = company + " " + ceo;
    const url = 'https://new.cretop.com/httpService/request.json';
    const headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6,zh;q=0.5',
        'content-type': 'application/json',
        'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        cookie: ck,
        'Referer': 'https://new.cretop.com/PL/IS/PLIS030M1?h=1686060519607',
        'Referrer-Policy': 'strict-origin-when-cross-origin'
    };

    const data = {
        "header": {
            "trxCd": "PLIS0301R",
            "chlType": "02",
            "screenId": "PLIS030M1",
            "menuId": "01W0000990",
            "langCd": "ko",
            "pageNum": 0,
            "pageCount": 0
        },
        "PLIS0301R": {
            "srchKeyWord": name,
            "keyWordHiltYn1": "Y",
            "brBznoIcldYn": "Y",
            "ordFldNm2": "pndDt",
            "ordMtd2": 1,
            "keyWordHiltYn2": "Y",
            "ordFldNm3": "RANK",
            "ordMtd3": 0,
            "keyWordHiltYn3": "N",
            "ordFldNm4": "RANK",
            "ordMtd4": 0,
            "keyWordHiltYn4": "Y",
            "pageNo": 1,
            "pageCn": 5,
            "excpPcVal1": "kedcd,bzno,cono,ksic10BzcCd,enpStatNm",
            "excpPcVal2": "",
            "excpPcVal3": "",
            "excpPcVal4": "reperNm,bzno,franTelNo,ctrNo,addr"
        }
    };

    try {
        const response = await axios.post(url, data, { headers });
        let tar = ""
        let res = removeTag(response.data.SPLIS03001OT.btCtt1.filter(obj => obj.enpStatNm === '정상'), regex)
        if (res[0] == undefined) return false;
        if (res.length > 1) {
            for (var i in res) {
                let res2 = await getCompanyInfo(res[i].kedcd);
                if (res2 == false) continue;
                tar = res[0];
                break;
            }
        } else {
            tar = res[0];
        }
        if (tar == "") return false
        let dump = await getCompanyInfo(tar.kedcd)
        let eco = await getEcoDatas(tar.kedcd)
        let cerfs = await getCert(tar.kedcd)
        if(eco == false) return false
        return {
            기업형태: tar.ipoNm,
            업태: await getUpTae(tar.kedcd),
            업체명: tar.enpNm,
            사업자등록번호:tar.bzno,
            대표자: tar.reperNm,
            휴대폰번호: phone == undefined ? "X" : convertPhoneNumber(phone),
            주소: dump.주소 == undefined ? "X" : dump.주소,
            종업원수: dump.종업원수 == undefined ? "X" : dump.종업원수,
            이메일: dump.이메일 == undefined ? "X" : dump.이메일,
            매출액: dump.매출액 == undefined ? 0 : dump.매출액,
            영업이익: dump.영업이익 == undefined ? 0 : dump.영업이익,
            자본금:eco.자본금,
            당기순이익: dump.당기순이익 == undefined ? 0 : dump.당기순이익,
            이익잉여금: eco.이익잉여금,
            장기차입금:eco.장기차입금,
            비유동부채:eco.비유동부채,
            벤처: cerfs.벤처,
            연구개발전담부서: cerfs.연구개발전담부서,
            부설연구소: cerfs.부설연구소,
            메인비즈: cerfs.메인비즈,
            이노비즈: cerfs.이노비즈,
        };
        // 여기서 응답 데이터를 사용할 수 있습니다.
    } catch (error) {
        return false
    }
}



const app = express();
app.use(cors());
const storage = multer.memoryStorage();
const upload = multer({
    storage,
    limits: {
        files: 1
    }
});
app.use(express.json());

app.use(express.urlencoded({
    extended: true
}));

app.use(
    expressSession({
        secret: "cretop",
        resave: true,
        saveUninitialized: true,
    })
);


app.set('view engine', 'ejs');
app.use(express.static('data'));

app.post('/uploadExcel', upload.array('xlsx', 1), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send('업로드된 파일이 없습니다.');
    }
    try {
        let data;
        for (const file of req.files) {
            const fileBuffer = file.buffer;
            const workbook = read(fileBuffer)
            const sheetName = workbook.SheetNames[0]
            const json = utils.sheet_to_json(workbook.Sheets[sheetName])
            data = json
        }
        res.status(200).json(data);
    } catch (e) {
        res.status(500);
    }
});

app.get('/', (req, res) => {
    res.render('main', {
        user: req.session.login
    });
})

app.post('/makePdf', async (req, res) => {
    let {
        id,
        data
    } = req.body
    await modifyPdf(data, id)
    res.json({"status":true})
})
app.post('/makeExcel', async (req, res) => {
    let {
        id,
        data
    } = req.body
    writeFileSync("data/excelData/" + id + ".xlsx", jsontoxlsx(data), 'binary')
    res.json({"status":true})
})

app.post('/makeZip', async (req, res) => {
    let {
        id
    } = req.body
    let data = await makeZip(id)
    if(data) res.json({"status":true})
    else res.json({"status":false})
})

app.get('/getCompany', async (req, res) => {
    let {
        company,
        ceo,
        phone
    } = req.query
    let data = await getCompany(company,ceo,phone)
    res.json(data)
})

app.get('/login', async (req, res) => {
    let {
        pw
    } = req.query
    if (pw === "test1234") {
        req.session.login = true
        res.redirect('/')
    } else {
        res.redirect('/')
    }
})

app.listen(3001, () => {
    console.log('서버가 3001번 포트에서 실행 중입니다.');
});