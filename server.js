const PDFDocument = require('pdfkit');
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const filesystem = require('fs').promises;
const createfile = require('fs');

async function generatePDF(rows) {
  const doc = new PDFDocument({ size: 'A4' });
  doc.font('Helvetica')
  doc.image('./logo.png', {width:350,align: 'center'})
  doc.lineGap(15)
  doc.fontSize(16).text('Relatório de ronda noturna', {align: 'center', underline: true});
  
  /* DEFININDO AS DATAS PARA BUSCAS */
  let searchSettings = { timeZone: 'UTC', timeZoneName: 'short' };
  /* DIA INÍCIO */
  let dia_inicio = new Date(2023, 4, 07);
  let dataUTC_inicio = dia_inicio.toLocaleDateString(searchSettings)
  /* DIA FIM */
  dia_fim = new Date(dia_inicio);
  dia_fim.setDate(dia_fim.getDate() + 8)
  let dataUTC_fim = dia_fim.toLocaleDateString(searchSettings)

  console.log('Período: ', dataUTC_inicio, dataUTC_fim) // RANGE DE DATAS

  /* CHAMADA DAS DATAS STRING PARA REALIZAR BUSCAS NA PLANILHA */
  function indexReportInicio() {
   return dataUTC_inicio;
  }
  function indexReportFim() {
    return dataUTC_fim;
   }

   /* BUSCA DO INDEX INICIAL E FINAL */
  let findInicio = rows.findIndex(row => row[0] == indexReportInicio() )
  console.log('início: ', findInicio) 
  
  let findFim = rows.findIndex(row => row[0] == indexReportFim() )
  console.log('fim: ', findFim)
  
  
  /* ITERAÇÃO DOS DADOS */
  for(findInicio; findInicio < findFim; findInicio++){
    const result = rows.filter((element, index) => index === findInicio);
    if (result !== undefined) {
      /* PARA CADA ELEMENTO, INSERE OS VALORES ABAIXO */
      result.forEach((row, index) => {
        doc.moveDown();
        doc.fontSize(14).text(`Data: ${row[0]}`, {align: 'center'});
        doc.fontSize(12).text(`Hora de início: ${row[2]}`);
        doc.fontSize(12).text(`Hora de término: ${row[3]}`);
        doc.fontSize(12).text(`TA's Margem Direita: ${row[4]}`);
        doc.fontSize(12).text(`TA's Margem Esquerda: ${row[5]}`);
        doc.fontSize(12).text(`Inspeção visual dos painéis dos TA's: ${row[6]}`);
        doc.fontSize(12).text(`Inspeção visual dos painéis das ELF's: ${row[7]}`);
        doc.fontSize(12).text(`Inspeção de todos elementos nas salas QCM's: ${row[8]}`);
        doc.fontSize(12).text(`Inspeção de todos elementos nas salas GMG's: ${row[9]}`);
        doc.fontSize(12).text(`Inspeção de todos elementos nas estruturas dos TA's: ${row[10]}`);
        doc.fontSize(12).text(`Alguma anomalia e/ou caça-perigo foi encontrado/registrado?: ${row[11].length ? row[11] : 'Não'}`);
        doc.fontSize(12).text(`Observações: ${doc.moveDown(), row[12].length ? row[12] : 'Não houveram fatos relevantes durante as rondas'}`, {
          align: 'justify',
          width: 500
        });
        doc.fontSize(12).text(`Inspeção visual dos painéis dos TA's: ${row[14]}`);
      });
    } else {
      console.log('nao encontrado')
    }
    doc.addPage();
  }
  doc.pipe(createfile.createWriteStream(`filepath`));
  doc.end();
}


// DEFINE O SCOPES
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];

// O arquivo token.json armazena o acesso do usuário e recarrega tokens e é criado
//automaticamente quando o fluxo de autorização completa pela primeira vez.
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'client.json');

async function loadSavedCredentialsIfExist() {
  try {
    const content = await filesystem.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

async function saveCredentials(client) {
  const content = await filesystem.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await filesystem.writeFile(TOKEN_PATH, payload);
}

/* Carrega, requisita ou autoriza para chamar API */
async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

async function listMajors(auth) {
  const sheets = google.sheets({version: 'v4', auth});
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: 'gheetsID',
    range: 'sheetname!B:R',
  });
  const rows = res.data.values;
  if (!rows || rows.length === 0) {
    console.log('No data found.');
    return;
  }

  // Gera PDF com dados das linhas
  await generatePDF(rows);
}

authorize().then(listMajors).catch(console.error);
