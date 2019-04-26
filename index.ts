const fs = require('fs');
const find = require('findit');
const path = require('upath');
const crypt = require('crypto');
const emlformat = require('eml-format');
const stringify = require('csv-stringify');
const tm = require('text-miner');

interface IEmail {
  file: string,
  messageId: string;
  dateSent: Date;
  subject: string;
  fromEmail: string;
  fromName: string;
  toEmails: string[];
  toNames: string[];
  ccEmails: string[];
  ccNames: string[];
  bccEmails: string[];
  bccNames: string[];
  text: string;
  textHash: string;
  attachments: string[];
}

const dirInput = '/Users/ralfbecher/Documents/Daten/maildir';
const dirOutput = './output';
const csvFileMeta = 'maildir.csv';
const csvFileTerms = 'mailterms.csv';
const csvFilePersons = 'mailpersons.csv';
const outputFieldsMeta = ['File', 'ID', 'Date', 'Subject', 'FromEmail', 'FromName', 'ToEmail', 'ToName', 'CCEmails', 'CCNames', 'BCCEmails', 'BCCNames', 'Text', 'TextLength', 'Hash', 'Attachments'];
const outputFieldsTerms = ['Hash', 'Term', 'Count'];
const outputFieldsPersons = ['ID', 'Name', 'Email', 'Role'];
const files:Set<string> = new Set([]);
const hashes:Set<string> = new Set([]);

if (!fs.existsSync(dirOutput)) {
  fs.mkdirSync(dirOutput);
}

let outputMeta = fs.createWriteStream(dirOutput + '/' + csvFileMeta, { flags: 'w' });
let outputTerms = fs.createWriteStream(dirOutput + '/' + csvFileTerms, { flags: 'w' });
let outputPersons = fs.createWriteStream(dirOutput + '/' + csvFilePersons, { flags: 'w' });

const stringifierMeta = stringify({
  delimiter: ',',
  header: true,
  quoted: true,
  columns: outputFieldsMeta
});

const stringifierTerms = stringify({
  delimiter: ',',
  header: true,
  quoted: true,
  columns: outputFieldsTerms
});

const stringifierPersons = stringify({
  delimiter: ',',
  header: true,
  quoted: true,
  columns: outputFieldsPersons
});

function startPipe() {
  stringifierMeta.pipe(outputMeta);
  stringifierTerms.pipe(outputTerms);
  stringifierPersons.pipe(outputPersons);  
}
startPipe();

function extractNames(emails: string): string[] {
  let names = [];
  emails = cleanUpEmailsAndNames(emails);
  if (emails) {
    let arrEmails = emails.split(',');
    names = arrEmails.map((e: string) => emlformat.getEmailAddress(e).name).filter((e: string) => e && e.length > 0);
  }
  return names;
}

async function mineText(hash: string, text: string): Promise<boolean> {
  return new Promise((resolve) => {
    if (!hashes.has(hash)) {
      hashes.add(hash);
      text = text.replace(/[:#_\\\/\t\(\)\'\"<>\[\]\*\$]/g,' ');
      let corpus = new tm.Corpus([text]);
      let cleansedCorpus = corpus.clean()
        .trim()
        .removeInterpunctuation()
        .removeInvalidCharacters()
        .removeDigits()
        .removeNewlines()
        .clean()
        .toLower()
        .removeWords(tm.STOPWORDS.EN);
      let terms = new tm.DocumentTermMatrix(cleansedCorpus);
      let p: Promise<boolean>[] = [];
      terms.vocabulary.forEach((e: string, i: number) => {
        p.push(writeRow(stringifierTerms, [hash, e, terms.data[0][i]]));
      });
      Promise.all(p)
      .then(() => resolve(true));
    } else {
      resolve(true);
    }
  });
}

function cleanUpEmailsAndNames(str: string): string {
  return str ? str.replace(/[\s'"]+/g, '') : '';
}

async function transposePersonsMails(id: string, mails: string[], names: string[], role: string): Promise<boolean> {
  return new Promise((resolve) => {
    if ((names.length === 0 && mails.length === 0)) {
      resolve(true);
    } else {
      let p: Promise<boolean>[] = [];
      if (mails.length === names.length) {
        names.forEach((e: string, i: number) => {
          p.push(writeRow(stringifierPersons, [id, e, mails[i], role]));
        });
      } else {
        mails.forEach((e: string) => {
          p.push(writeRow(stringifierPersons, [id, '', e, role]));
        });
      }
      return Promise.all(p)
        .then(() => resolve(true));
    }
  });
}

async function writeRow(stream: any, row: string[]): Promise<boolean> {
  return new Promise((resolve) => {
    stream.write(row, () => { resolve(true); });
    // let res = stream.write(row);
    // if (res) {
    //   resolve(true);
    // } else {
    //   console.log("\n\nDRAIN...");
    //   stream.pause();
    //   stream.once('drain', () => {
    //     stream.resume();
    //     console.log("\n\nRESUME...");
    //     resolve(true);
    //   });
    // }
  });
}

async function convertEml(fileName: string, eml: string): Promise<IEmail> {
  return new Promise((resolve, reject) => {
    let email: IEmail = {
      file: '',
      messageId: '',
      dateSent: new Date(),
      subject: '',
      fromEmail: '',
      fromName: '',
      toEmails: [],
      toNames: [],
      ccEmails: [],
      ccNames: [],
      bccEmails: [],
      bccNames: [],
      text: '',
      textHash: '',
      attachments: []
    };

    if (eml.substr(0,11) === 'Message-ID:') {
      try {
        emlformat.read(eml, (error: any, data: any) => {
          if (error) return console.log(error);

          email.file = fileName;
          email.messageId = data.headers['Message-ID'] || '';
          email.dateSent = data.date;

          if (data.text) {
            email.textHash = require('crypto').createHash('md5').update(data.text).digest('base64');
            email.text = data.text;
          } else {
            email.textHash = '0000'; // placeholder for empty text#
            email.text = '';
          }

          if (data.attachments && data.attachments.length > 0) {
            email.attachments = data.attachments.map((e: { name: string; }) => e.name);
          }
          email.fromEmail = data.from && data.from.email ? data.from.email : '';
          if (data.headers['X-From']) {
            email.fromName = emlformat.getEmailAddress(data.headers['X-From']).name;
          }
          email.fromEmail = cleanUpEmailsAndNames(email.fromEmail);
          email.fromName = cleanUpEmailsAndNames(email.fromName && email.fromName.length > 0 ? email.fromName : data.headers['X-From'] || '');
          email.toEmails = data.to && data.to.email ? cleanUpEmailsAndNames(data.to.email).split(',') : [];
          email.toNames = extractNames(data.headers['X-To']);
          email.ccEmails = data.headers['Cc'] ? data.headers['Cc'].split(',') : [];
          email.ccNames = extractNames(data.headers['X-cc']);
          email.bccEmails = data.headers['Bcc'] ? data.headers['Bcc'].split(',') : [];
          email.bccNames = extractNames(data.headers['X-bcc']);

          resolve(email);          
        });  
      } catch (error) {
        console.error(error);
        resolve(email);
      }
    } else {
      resolve(email);
    }
  });
}

async function processEmlFile(emlFile: string): Promise<boolean> {
  return new Promise(async (resolve) => {
    let eml = fs.readFileSync(emlFile, 'utf-8');
    if (eml.substr(0,11) === 'Message-ID:') {
      console.log('Processing ' + emlFile);
      let email: IEmail = await convertEml(emlFile.substr(dirInput.length), eml);
      let dateSent = '';
      try {
        dateSent = email.dateSent.toISOString();
      } catch {
        // do nothing
      }
      let csvRowMeta: string[] = [
        email.file,
        email.messageId,
        dateSent,
        email.subject || '',
        email.fromEmail,
        email.fromName,
        email.toEmails.join(','),
        email.toNames.join(','),
        email.ccEmails.join(','),
        email.ccNames.join(','),
        email.bccEmails.join(','),
        email.bccNames.join(','),
        email.text,
        email.text.length.toString(),
        email.textHash,
        email.attachments.join(',')
      ];

      try {
        await writeRow(stringifierMeta, csvRowMeta);
        await mineText(email.textHash, email.text);
        await transposePersonsMails(email.messageId, [email.fromEmail], [email.fromName], 'From');
        await transposePersonsMails(email.messageId, email.toEmails, email.toNames, 'To');
        await transposePersonsMails(email.messageId, email.ccEmails, email.ccNames, 'Cc');
        await transposePersonsMails(email.messageId, email.bccEmails, email.bccNames, 'Bcc');
        return resolve(true);
      }
      catch (error) {
        console.error(error);
        process.exit(1);
      }
    } else {
      resolve(true);
    }  
  });
}

// async function waitWrapper(file: string): Promise<boolean> {
//   await processEmlFile(file);
//   return new Promise((resolve) => { resolve(true) });
// }

var finder = find(dirInput);
console.log("Collecting EML files...");
finder.on('file', (file: any) => {
  files.add(path.normalize(file));
});

finder.on('end', () => {
  let p: Promise<boolean>[] = [];
  for (let file of files.values()) {
    p.push(processEmlFile(file));
  }
  Promise.all(p)
  .then(() => {
    stringifierMeta.end();
    stringifierTerms.end();
    stringifierPersons.end();
    console.log("Processing finished.");
  });
});
