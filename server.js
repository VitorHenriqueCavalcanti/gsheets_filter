const PDFDocument = require('pdfkit');
const path = require('path');
const process = require('process');
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');
const fs = require('fs').promises;
const fss = require('fs');


// Cria o arquivo PDF
async function generatePDF(rows) {
  const doc = new PDFDocument({size: 'A4'});
  doc.pipe(fss.createWriteStream('output.pdf'));
  doc.font('Helvetica')
  doc.lineGap(15)
  doc.image('./logo.png', {width:350,align: 'center', lineGap: 10})
  doc.fontSize(16).text('Relatório de ronda noturna', {align: 'center', underline: true});

  let day = 05;
  let month = '02';
  let year = 2023;
  let days = parseInt(day) + 7

  for(day; day < days; day++){
    let dateSearch = `${day < 10 ? `${'0'+day}`: day}/${month}/${year}`;
    console.log(dateSearch)
    const result = rows.filter(row => row[0] == (dateSearch));
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
      doc.fontSize(12).text(`Observações: ${doc.moveDown(), row[12]}`, {
        align: 'justify',
        width: 500
      });
      doc.fontSize(12).text(`Inspeção visual dos painéis dos TA's: ${row[14]}`);
    });
    doc.addPage();
  }
    doc.end();
}


// Define o SCOPES
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];

// O arquivo token.json armazena o acesso do usuário e recarrega tokens e é criado
//automaticamente quando o fluxo de autorização completa pela primeira vez.
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'client.json');

/**.
 * @return {Promise<OAuth2Client|null>}
 */

async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */

async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Carrega ou requisita ou autoriza para chamar APIs
 *
 */
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
    spreadsheetId: 'idDaPlanilhas',
    range: 'Relatório de ronda noturna!B:R',
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