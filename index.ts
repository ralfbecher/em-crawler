const fs = require('fs');
const find = require('findit');
const path = require('upath');
const crypt = require('crypto');
const simpleParser = require('mailparser').simpleParser;
const stringify = require('csv-stringify');
const tm = require('text-miner');
const nGram = require('n-gram')
const htmlTags = require('html-tags');
const voidHtmlTags = require('html-tags/void');
const cssProperties = require('known-css-properties').all;

const additionalWords = ['http', 'https', 'www', 'cc', 'nbsp', 'gif', 'image', 'images', 'spacer', 'href', 'arial', 'helvetica', 'verdana', 
  'bgcolor', 'ffffff', 'mime', 'mailto', 'cgi', 'bin'];
const removeWords: string[] = tm.STOPWORDS.EN.concat(additionalWords).concat(htmlTags).concat(voidHtmlTags).concat(cssProperties);

interface IEmail {
  file: string,
  messageId: string;
  dateSent: Date;
  subject: string;
  fromEmail: string;
  fromName: string;
  toEmails: string[];
  ccEmails: string[];
  bccEmails: string[];
  textHash: string;
  text: string;
  textLength: number;
  attachments: string[];
}

const dirInput = '/Users/ralfbecher/Documents/Daten/maildir';
// const dirInput = '/Users/ralfbecher/Documents/Daten/test';
const dirOutput = './output';
const csvFileMeta = 'maildir.csv';
const csvFileTexts = 'mailtexts.csv';
const csvFileTerms = 'mailterms.csv';
const csvFileNGrams = 'mailngrams.csv';
const csvFilePersons = 'mailpersons.csv';
const outputFieldsMeta = ['File', 'ID', 'Date', 'Subject', 'FromEmail', 'FromName', 'ToEmails', 'CCEmails', 'BCCEmails', 'TextLength', 'Hash', 'Attachments'];
const outputFieldsTexts = ['Hash', 'CleansedText'];
const outputFieldsTerms = ['Hash', 'Term', 'Count'];
const outputFieldsNGrams = ['Hash', 'nGram', 'Type', 'Count'];
const outputFieldsPersons = ['ID', 'Email', 'Role'];
const files:Set<string> = new Set([]);
const hashes:Set<string> = new Set([]);

if (!fs.existsSync(dirOutput)) {
  fs.mkdirSync(dirOutput);
}

let outputMeta = fs.createWriteStream(dirOutput + '/' + csvFileMeta, { flags: 'w' });
let outputTexts = fs.createWriteStream(dirOutput + '/' + csvFileTexts, { flags: 'w' });
let outputTerms = fs.createWriteStream(dirOutput + '/' + csvFileTerms, { flags: 'w' });
let outputNGrams = fs.createWriteStream(dirOutput + '/' + csvFileNGrams, { flags: 'w' });
let outputPersons = fs.createWriteStream(dirOutput + '/' + csvFilePersons, { flags: 'w' });

const stringifierMeta = stringify({
  delimiter: ',',
  header: true,
  quoted: true,
  columns: outputFieldsMeta
});

const stringifierTexts = stringify({
  delimiter: ',',
  header: true,
  quoted: true,
  columns: outputFieldsTexts
});

const stringifierTerms = stringify({
  delimiter: ',',
  header: true,
  quoted: true,
  columns: outputFieldsTerms
});

const stringifierNGrams = stringify({
  delimiter: ',',
  header: true,
  quoted: true,
  columns: outputFieldsNGrams
});

const stringifierPersons = stringify({
  delimiter: ',',
  header: true,
  quoted: true,
  columns: outputFieldsPersons
});

function startPipe(): void {
  stringifierMeta.pipe(outputMeta);
  stringifierTexts.pipe(outputTexts);
  stringifierTerms.pipe(outputTerms);
  stringifierNGrams.pipe(outputNGrams);
  stringifierPersons.pipe(outputPersons);  
}

function endPipe(): void {
  stringifierMeta.end();
  stringifierTexts.end();
  stringifierTerms.end();
  stringifierNGrams.end();
  stringifierPersons.end();
}

function nGramAggr(nGrams: string[][]): string[][] {
  let res: string[][] = [];
  let gramMap: Map<string, number[]> = new Map();
  nGrams.forEach((e: string[], i: number) => {
    let t = e.length;
    // filter one character words
    if (t === e.filter(w => w.length > 1).length) { 
        let k = e.join(' ');
      if (!gramMap.has(k)) {
        gramMap.set(k, [t, 1]);
      } else {
        let n: number[] = (gramMap.get(k) as number[]);
        gramMap.set(k, [n[0], n[1] +1]);
      }
    }
  });
   
  for (let e of gramMap.entries()) {
    res.push([e[0], e[1][0].toString(), e[1][1].toString()]);
  }
  return res;
}

async function mineText(hash: string, text: string): Promise<boolean> {
  return new Promise(async (resolve) => {
    if (!hashes.has(hash)) {
      hashes.add(hash);
      text = tm.expandContractions(text);
      text = text.replace(/([^\x20-\uD7FF\uE000-\uFFFC\u{10000}-\u{10FFFF}])/ug,' ').replace(/[\x7F:#_\\\/\t\(\)\[\]\*\$\+\|'"{}<>=~&%`´]/g,' ');
      let corpus = new tm.Corpus([text]);
      let cleansedCorpus = corpus.clean()
        .trim()
        .removeInterpunctuation()
        .removeInvalidCharacters()
        .removeDigits()
        .removeNewlines()
        .clean()
        .toLower()
        .removeWords(removeWords)
        .stem('Porter');
      let docs = cleansedCorpus.documents;
      await writeRow(stringifierTexts, [hash, docs]);
      let list = docs[0].text ? docs[0].text.split(' ') : [];
      let biGrams = nGram.bigram(list);
      // let triGrams = nGram.trigram(list);
      let terms = new tm.DocumentTermMatrix(cleansedCorpus);
      terms.vocabulary.forEach(async (e: string, i: number) => {
        await writeRow(stringifierTerms, [hash, e, terms.data[0][i]]);
      });
      // nGramAggr(biGrams.concat(triGrams)).forEach(async (e: string[]) => {
      nGramAggr(biGrams).forEach(async (e: string[]) => {
        await writeRow(stringifierNGrams, [hash, e[0], e[1], e[2]]);
      });
      resolve(true);
    } else {
      resolve(true);
    }
  });
}

function cleanUpEmailsAndNames(str: string): string {
  return str ? str.trim().replace(/[\r\n|\r|\n'"]+/g, '') : '';
}

function switchNameParts(str: string): string {
  if (str && str.indexOf('@') === -1 && str.indexOf(',') !== -1) {
    let parts = str.split(',');
    if (parts.length === 2) {
      str = parts[1].trim() + ' ' + parts[0].trim();
    }
  }
  return str;
}

async function transposeMails(id: string, mails: string[], role: string): Promise<boolean> {
  return new Promise((resolve) => {
    if (mails.length === 0) {
      resolve(true);
    } else {
      mails.forEach(async (e: string) => {
        await writeRow(stringifierPersons, [id, e, role]);
      });
      resolve(true);
    }
  });
}

async function writeRow(stream: any, row: string[]): Promise<boolean> {
  return new Promise(async (resolve) => {
    await stream.write(row);
    resolve(true);
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

function headers2obj(headers: Map<string, string>): object {
  let res: any = {};
  for (let e of headers.entries()) {
    res[e[0]] = e[1];
  }
  return res;
}

function getNameFromEmailAddress(email: string): string {
  let result = '';
  let regex = /^(.*?)(\s*\<(.*?)\>)$/g;
  let match = regex.exec(email);
  if (match) {
    let name = match[1].replace(/"/g, "").trim();
    if (name && name.length) {
      result = name;
    }
  }
  return result;
}

async function convertEml(fileName: string, eml: string): Promise<IEmail> {
  return new Promise(async (resolve, reject) => {
    let email: IEmail = {
      file: '',
      messageId: '',
      dateSent: new Date(),
      subject: '',
      fromEmail: '',
      fromName: '',
      toEmails: [],
      ccEmails: [],
      bccEmails: [],
      textHash: '',
      text: '',
      textLength: 0,
      attachments: []
    };

    try {
      let data: any = await simpleParser(eml);
      let headers: any = headers2obj(data.headers);

      email.file = fileName;
      email.messageId = data.messageId || '';
      email.dateSent = data.date;
      email.subject = data.subject;

      if (data.text) {
        email.textHash = require('crypto').createHash('md5').update(data.text).digest('base64');
        email.text = data.text;
        email.textLength = data.text.lengh;
      } else {
        email.textHash = '0000'; // placeholder for empty text
      }

      email.fromEmail = cleanUpEmailsAndNames(data.from && data.from.text ? data.from.text : '');
      email.fromName = switchNameParts(getNameFromEmailAddress(cleanUpEmailsAndNames(headers['x-from'] || '')));

      email.toEmails = data.to && data.to.text ? cleanUpEmailsAndNames(data.to.text).split(',').map((e: string) => e.trim()) : [];
      email.ccEmails = headers['cc'] ? headers['cc'].text.split(',').map((e: string) => e.trim()) : [];
      email.bccEmails = headers['bcc'] ? headers['bcc'].text.split(',').map((e: string) => e.trim()) : [];
      if (email.ccEmails.length && email.bccEmails.length && email.ccEmails.join(',') === email.bccEmails.join(',')) {
        email.bccEmails = [];
      }

      if (data.attachments && data.attachments.length > 0) {
        email.attachments = data.attachments
          .filter((e: { contentDisposition: string }) => e.contentDisposition === 'attachment')
          .map((e: { filename: string; }) => e.filename);
      }
      resolve(email);          
    } catch (error) {
      console.error(error);
      resolve(email);
    }
  });
}

async function processEmlFile(emlFile: string): Promise<boolean> {
  return new Promise(async (resolve) => {
    let eml = fs.readFileSync(emlFile, 'utf-8');
    if (eml.substr(0,11) === 'Message-ID:' || eml.substr(0,14) === 'Received: from') {
      console.log('Processing ' + emlFile);
      let email: IEmail = await convertEml(emlFile.substr(dirInput.length), eml);
      let dateSent = '';
      try {
        dateSent = email.dateSent.toISOString();
      } catch {
        // do nothing
      }
      try {
        
        let csvRowMeta: string[] = [
          email.file,
          email.messageId,
          dateSent,
          email.subject || '',
          email.fromEmail,
          email.fromName,
          email.toEmails.join(','),
          email.ccEmails.join(','),
          email.bccEmails.join(','),
          email.textHash,
          email.textLength.toString(),
          email.attachments.join(',')
        ];
  
        await writeRow(stringifierMeta, csvRowMeta);
        await mineText(email.textHash, email.text);
        // await transposeMails(email.messageId, [email.fromEmail], 'From');
        await transposeMails(email.messageId, email.toEmails, 'To');
        await transposeMails(email.messageId, email.ccEmails, 'Cc');
        await transposeMails(email.messageId, email.bccEmails, 'Bcc');
        resolve(true);
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

var finder = find(dirInput);
console.log("Collecting EML files...");
finder.on('file', (file: any) => {
  files.add(path.normalize(file));
});

async function processAll(): Promise<boolean> {
  console.log("Collecting EML files finished...");
  return new Promise(async (resolve) => {
    startPipe();
    for (let file of files.values()) {
      await processEmlFile(file);
    }
    endPipe();
    console.log("Processing finished.");
    resolve(true);
  });
}

finder.on('end', () => {
  processAll();
});
