const express = require('express');
const { promisify } = require('util');
const app = express();


app.use(express.static('./pages/public'))

const nunjucks = require('nunjucks');
nunjucks.configure('./pages/', {
    express: app
});



const docId = '1Tq_elTOWO644zZQ1OiimdoTnMQH8B0xSOnjezOuQ64o';



const { GoogleSpreadsheet } = require('google-spreadsheet');
const { ALPN_ENABLED } = require('constants');

var datas = [];
let totalfaults = 0 ;




app.get('/', async (request, response) => {

    await  loading();
   

    return response.render("index.html",{datas});
})
 


app.get('/tarefa', async (request, response) => {

    const doc = new GoogleSpreadsheet(docId);

    // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth({
        client_email: "tarefa-google@quickstart-1610739987253.iam.gserviceaccount.com",
        private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDEcaJLmv99DWVh\nSXiNR6biFwrSG1fQ2nqKNjy2DRdLppVCWU2L3sHWD0AJa5IApHpgBCrKUWKXUcOO\nQIjS9aMHXo5LJiChnPGtZntCXN3BKOUFIe8+mxhUGBzNGwSwlTOzbcSTcUhLhfmU\nUkEL1rDF+UkB23j+UaH3UUZewN1dFT3EkEa8GFRs8b/Wupw8S44CW+bTLGolccEA\nlEYsq9OZnLOxcmZIiWOF9afQggvtDgG7NcGhOz3u5GJw3QR1/FG6OmPTDdyH8B1K\ncnq4kzBVRwhaSZnxhaGKvxNM19vBt14Z5ZGZ2VOmrXD37C6p+nmHMxRP7hkOvN3i\nQXBoAbRxAgMBAAECggEABgNManAGHffJAJ9VF034J7d411GK8JOfaJecaB4idmhU\n7UD6hKt+12SEG0W1pFtke4flH2g6UlNoXvROu9ZU9SbJyDcUjJ3XL+2RHEjnaMAt\nsmiFgC8TIY/TYdvP2u/WM0nK2JCBG/6v0wBpiUk7A/RLbckf/PjWslFEjCXvIKg2\nDwI6peqdsNdGr6lHmVoA33foZSYExY1zNuTqr5RsGWsBDGflGapb1VcUV9QE40yX\nRj3LgsCOPsROGkvbGrhnisw8E+Ea0rm+kyRpa5lCdWhCKUE8ue7vjpEDmvLlir6P\nX6BqvPFO+apf58A/i/nRRgJbFC9VY5kr6Z+OuD4W+QKBgQDjSAUwT8/J8A4bfeB+\nuSD6xzDuRokBobE7RSSTcc1G1IAH8RbJB8D6Q/3jmTl3uI9NnEz6S6aHVM5rJz1l\nO2QDJwUSP1qMLenF72mUF0iBxMqeEewsXhMLZ+ZYkCZCpEk1VcrUEDaj9QOWcmK6\nIvZ+rbic7TnbKoj0CnvfkEJ0MwKBgQDdRBmNV/VifEFfkcYL0hWeBLQMk/zAHV/y\nKa5Ivo5ygFArMEkLkgQzL7rOTzq0c5q0x7zo9JVaCc7mXYl+Oumxc9fKiyB26eeM\nE3F5pMc6mlk/dyZMwbApuZ4A7mZ/IHrn6aaP7rPMJ+uA6xnAjwJJ/ojgN8Wm7dSE\nqyxhpmIwywKBgHDzx/Bcmc2oCbrL8hfIdYVsHPst/sTa0LO+BxFnyzbaQM6xmDtM\nKTG3PKQx8Ad5p25QsUjq89Xp5bQHClIXE/slFzYcWim0X6vI8dVxRM2JOZEZIyBh\nmGFgv29gJEOWVfO1sVl2vVD6YVARhNMwsQP/3fHPS6OKHgn6c9mFXiFVAoGAe8PF\nzyvuFAKQxpZRgvcmJFdZJtf4PrWvn1L1K7d7Ekz3itDdat1oAAGoqhHjMmCfnpNC\n9cMpb02hL3YOnE7zvNChWafsptc7Lz0I8hPbZMpFNZy+DZ0hnpU27iprppxSYzps\ncoIAjCegMWJP60eS7jSz90b7Bd5uSy88Cfr5XXUCgYAzyuIJtk8ZDkn2Jelvqodg\n1TpZ55JQU0v+KEx12xyTuscp1h9zuBkh03FOCat/wgKHucCmAwASXCFUJ99R2b82\nYyEvxwFwNlbHsfdb7yYy1Dy1X7HWWtioONRoaoquPQOLGA1sR4ikCp6UgfFCGNSa\nP43mRUw+9u9lvgPn4R1Umw==\n-----END PRIVATE KEY-----\n",
    });

    await doc.loadInfo(); // loads document properties and worksheets
    console.log(doc.title);
    await doc.updateProperties({ title: 'Engenharia de Software – Desafio [Pablo]' });
    const sheet = doc.sheetsByIndex[0]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
    await sheet.loadCells('A2:H27'); // loads a range of cells

    let total = 26;
    totalfaults = Number(sheet.getCell(1, 0).value.split(":")[1]);
    for (i = 3; i <= total; i++) {


        let cellSituation = sheet.getCell(i, 6);
        let cellFinal = sheet.getCell(i, 7);
        let faults = sheet.getCell(i, 2).value;
        let p1 = sheet.getCell(i, 3).value;
        let p2 = sheet.getCell(i, 4).value;
        let p3 = sheet.getCell(i, 5).value;



        cellSituation.value = situation(p1, p2, p3, faults)
        cellFinal.value = gradeFinal(p1, p2, p3, faults,);
        
        await sheet.saveUpdatedCells();
    }

   

    await loading()

   
        return response.render("index.html",{datas});
   

    

});

async function  loading(){
    const doc = new GoogleSpreadsheet(docId);

    // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth({
        client_email: "tarefa-google@quickstart-1610739987253.iam.gserviceaccount.com",
        private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDEcaJLmv99DWVh\nSXiNR6biFwrSG1fQ2nqKNjy2DRdLppVCWU2L3sHWD0AJa5IApHpgBCrKUWKXUcOO\nQIjS9aMHXo5LJiChnPGtZntCXN3BKOUFIe8+mxhUGBzNGwSwlTOzbcSTcUhLhfmU\nUkEL1rDF+UkB23j+UaH3UUZewN1dFT3EkEa8GFRs8b/Wupw8S44CW+bTLGolccEA\nlEYsq9OZnLOxcmZIiWOF9afQggvtDgG7NcGhOz3u5GJw3QR1/FG6OmPTDdyH8B1K\ncnq4kzBVRwhaSZnxhaGKvxNM19vBt14Z5ZGZ2VOmrXD37C6p+nmHMxRP7hkOvN3i\nQXBoAbRxAgMBAAECggEABgNManAGHffJAJ9VF034J7d411GK8JOfaJecaB4idmhU\n7UD6hKt+12SEG0W1pFtke4flH2g6UlNoXvROu9ZU9SbJyDcUjJ3XL+2RHEjnaMAt\nsmiFgC8TIY/TYdvP2u/WM0nK2JCBG/6v0wBpiUk7A/RLbckf/PjWslFEjCXvIKg2\nDwI6peqdsNdGr6lHmVoA33foZSYExY1zNuTqr5RsGWsBDGflGapb1VcUV9QE40yX\nRj3LgsCOPsROGkvbGrhnisw8E+Ea0rm+kyRpa5lCdWhCKUE8ue7vjpEDmvLlir6P\nX6BqvPFO+apf58A/i/nRRgJbFC9VY5kr6Z+OuD4W+QKBgQDjSAUwT8/J8A4bfeB+\nuSD6xzDuRokBobE7RSSTcc1G1IAH8RbJB8D6Q/3jmTl3uI9NnEz6S6aHVM5rJz1l\nO2QDJwUSP1qMLenF72mUF0iBxMqeEewsXhMLZ+ZYkCZCpEk1VcrUEDaj9QOWcmK6\nIvZ+rbic7TnbKoj0CnvfkEJ0MwKBgQDdRBmNV/VifEFfkcYL0hWeBLQMk/zAHV/y\nKa5Ivo5ygFArMEkLkgQzL7rOTzq0c5q0x7zo9JVaCc7mXYl+Oumxc9fKiyB26eeM\nE3F5pMc6mlk/dyZMwbApuZ4A7mZ/IHrn6aaP7rPMJ+uA6xnAjwJJ/ojgN8Wm7dSE\nqyxhpmIwywKBgHDzx/Bcmc2oCbrL8hfIdYVsHPst/sTa0LO+BxFnyzbaQM6xmDtM\nKTG3PKQx8Ad5p25QsUjq89Xp5bQHClIXE/slFzYcWim0X6vI8dVxRM2JOZEZIyBh\nmGFgv29gJEOWVfO1sVl2vVD6YVARhNMwsQP/3fHPS6OKHgn6c9mFXiFVAoGAe8PF\nzyvuFAKQxpZRgvcmJFdZJtf4PrWvn1L1K7d7Ekz3itDdat1oAAGoqhHjMmCfnpNC\n9cMpb02hL3YOnE7zvNChWafsptc7Lz0I8hPbZMpFNZy+DZ0hnpU27iprppxSYzps\ncoIAjCegMWJP60eS7jSz90b7Bd5uSy88Cfr5XXUCgYAzyuIJtk8ZDkn2Jelvqodg\n1TpZ55JQU0v+KEx12xyTuscp1h9zuBkh03FOCat/wgKHucCmAwASXCFUJ99R2b82\nYyEvxwFwNlbHsfdb7yYy1Dy1X7HWWtioONRoaoquPQOLGA1sR4ikCp6UgfFCGNSa\nP43mRUw+9u9lvgPn4R1Umw==\n-----END PRIVATE KEY-----\n",
    });

    await doc.loadInfo(); // loads document properties and worksheets
    await doc.updateProperties({ title: 'Engenharia de Software – Desafio [Pablo]' });
    const sheet = doc.sheetsByIndex[0]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
    await sheet.loadCells('A2:H27'); // loads a range of cells
  

    const rows = await sheet.getRows(); // can pass in { limit, offset }
    const totalLine = rows.length 

    datas = [];

     for( i = 3 ; i <= totalLine ; i ++ ){
     
        datas.push({"matricula":sheet.getCell(i, 0).value,
        "nome":sheet.getCell(i, 1).value,
        "faltas":sheet.getCell(i, 2).value,
        "p1":sheet.getCell(i, 3).value,
        "p2":sheet.getCell(i, 4).value,  
        "p3":sheet.getCell(i, 5).value,
        "situacao":sheet.getCell(i, 6).value,
        "nota_final":sheet.getCell(i, 7).value});
     }
}


  function situation(p1, p2, p3, faults) {
   
    if (faults > (totalfaults / 4)) {
        return "Reprovado por falta";
    }

    let average = (p1 + p2 + p3) / 3;

    if (average >= 70) {
        return "Aprovado"
    } else if (average <= 50 && average < 70) {
        return "Prova final"
    } else {
        return "Reprovado por Nota"
    }

}

function gradeFinal (p1, p2, p3, faults){
    
    if (faults > (totalfaults / 4)) {
        return "";
    } 

    let average = (p1 + p2 + p3) / 3;

    if(average >= 50 && average < 70 ){
       let  grade =  (70 - average );
        return  Math.round( average + ( grade * 2)) ;
    } else if (average < 50){
        return " ";
    }else{
        return 0;
    }

}


 


app.listen(3333, () => {
    console.log('connected server');
});