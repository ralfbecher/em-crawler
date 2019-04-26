"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __values = (this && this.__values) || function (o) {
    var m = typeof Symbol === "function" && o[Symbol.iterator], i = 0;
    if (m) return m.call(o);
    return {
        next: function () {
            if (o && i >= o.length) o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
};
var fs = require('fs');
var find = require('findit');
var path = require('upath');
var crypt = require('crypto');
var emlformat = require('eml-format');
var stringify = require('csv-stringify');
var tm = require('text-miner');
var dirInput = '/Users/ralfbecher/Documents/Daten/maildir';
var dirOutput = './output';
var csvFileMeta = 'maildir.csv';
var csvFileTerms = 'mailterms.csv';
var csvFilePersons = 'mailpersons.csv';
var outputFieldsMeta = ['File', 'ID', 'Date', 'Subject', 'FromEmail', 'FromName', 'ToEmail', 'ToName', 'CCEmails', 'CCNames', 'BCCEmails', 'BCCNames', 'Text', 'TextLength', 'Hash', 'Attachments'];
var outputFieldsTerms = ['Hash', 'Term', 'Count'];
var outputFieldsPersons = ['ID', 'Name', 'Email', 'Role'];
var files = new Set([]);
var hashes = new Set([]);
if (!fs.existsSync(dirOutput)) {
    fs.mkdirSync(dirOutput);
}
var outputMeta = fs.createWriteStream(dirOutput + '/' + csvFileMeta, { flags: 'w' });
var outputTerms = fs.createWriteStream(dirOutput + '/' + csvFileTerms, { flags: 'w' });
var outputPersons = fs.createWriteStream(dirOutput + '/' + csvFilePersons, { flags: 'w' });
var stringifierMeta = stringify({
    delimiter: ',',
    header: true,
    quoted: true,
    columns: outputFieldsMeta
});
var stringifierTerms = stringify({
    delimiter: ',',
    header: true,
    quoted: true,
    columns: outputFieldsTerms
});
var stringifierPersons = stringify({
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
function extractNames(emails) {
    var names = [];
    emails = cleanUpEmailsAndNames(emails);
    if (emails) {
        var arrEmails = emails.split(',');
        names = arrEmails.map(function (e) { return emlformat.getEmailAddress(e).name; }).filter(function (e) { return e && e.length > 0; });
    }
    return names;
}
function mineText(hash, text) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) {
                    if (!hashes.has(hash)) {
                        hashes.add(hash);
                        text = text.replace(/[:#_\\\/\t\(\)\'\"<>\[\]\*\$]/g, ' ');
                        var corpus = new tm.Corpus([text]);
                        var cleansedCorpus = corpus.clean()
                            .trim()
                            .removeInterpunctuation()
                            .removeInvalidCharacters()
                            .removeDigits()
                            .removeNewlines()
                            .clean()
                            .toLower()
                            .removeWords(tm.STOPWORDS.EN);
                        var terms_1 = new tm.DocumentTermMatrix(cleansedCorpus);
                        var p_1 = [];
                        terms_1.vocabulary.forEach(function (e, i) {
                            p_1.push(writeRow(stringifierTerms, [hash, e, terms_1.data[0][i]]));
                        });
                        Promise.all(p_1)
                            .then(function () { return resolve(true); });
                    }
                    else {
                        resolve(true);
                    }
                })];
        });
    });
}
function cleanUpEmailsAndNames(str) {
    return str ? str.replace(/[\s'"]+/g, '') : '';
}
function transposePersonsMails(id, mails, names, role) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) {
                    if ((names.length === 0 && mails.length === 0)) {
                        resolve(true);
                    }
                    else {
                        var p_2 = [];
                        if (mails.length === names.length) {
                            names.forEach(function (e, i) {
                                p_2.push(writeRow(stringifierPersons, [id, e, mails[i], role]));
                            });
                        }
                        else {
                            mails.forEach(function (e) {
                                p_2.push(writeRow(stringifierPersons, [id, '', e, role]));
                            });
                        }
                        return Promise.all(p_2)
                            .then(function () { return resolve(true); });
                    }
                })];
        });
    });
}
function writeRow(stream, row) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) {
                    stream.write(row, function () { resolve(true); });
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
                })];
        });
    });
}
function convertEml(fileName, eml) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve, reject) {
                    var email = {
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
                    if (eml.substr(0, 11) === 'Message-ID:') {
                        try {
                            emlformat.read(eml, function (error, data) {
                                if (error)
                                    return console.log(error);
                                email.file = fileName;
                                email.messageId = data.headers['Message-ID'] || '';
                                email.dateSent = data.date;
                                if (data.text) {
                                    email.textHash = require('crypto').createHash('md5').update(data.text).digest('base64');
                                    email.text = data.text;
                                }
                                else {
                                    email.textHash = '0000'; // placeholder for empty text#
                                    email.text = '';
                                }
                                if (data.attachments && data.attachments.length > 0) {
                                    email.attachments = data.attachments.map(function (e) { return e.name; });
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
                        }
                        catch (error) {
                            console.error(error);
                            resolve(email);
                        }
                    }
                    else {
                        resolve(email);
                    }
                })];
        });
    });
}
function processEmlFile(emlFile) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) { return __awaiter(_this, void 0, void 0, function () {
                    var eml, email, dateSent, csvRowMeta, error_1;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                eml = fs.readFileSync(emlFile, 'utf-8');
                                if (!(eml.substr(0, 11) === 'Message-ID:')) return [3 /*break*/, 11];
                                console.log('Processing ' + emlFile);
                                return [4 /*yield*/, convertEml(emlFile.substr(dirInput.length), eml)];
                            case 1:
                                email = _a.sent();
                                dateSent = '';
                                try {
                                    dateSent = email.dateSent.toISOString();
                                }
                                catch (_b) {
                                    // do nothing
                                }
                                csvRowMeta = [
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
                                _a.label = 2;
                            case 2:
                                _a.trys.push([2, 9, , 10]);
                                return [4 /*yield*/, writeRow(stringifierMeta, csvRowMeta)];
                            case 3:
                                _a.sent();
                                return [4 /*yield*/, mineText(email.textHash, email.text)];
                            case 4:
                                _a.sent();
                                return [4 /*yield*/, transposePersonsMails(email.messageId, [email.fromEmail], [email.fromName], 'From')];
                            case 5:
                                _a.sent();
                                return [4 /*yield*/, transposePersonsMails(email.messageId, email.toEmails, email.toNames, 'To')];
                            case 6:
                                _a.sent();
                                return [4 /*yield*/, transposePersonsMails(email.messageId, email.ccEmails, email.ccNames, 'Cc')];
                            case 7:
                                _a.sent();
                                return [4 /*yield*/, transposePersonsMails(email.messageId, email.bccEmails, email.bccNames, 'Bcc')];
                            case 8:
                                _a.sent();
                                return [2 /*return*/, resolve(true)];
                            case 9:
                                error_1 = _a.sent();
                                console.error(error_1);
                                process.exit(1);
                                return [3 /*break*/, 10];
                            case 10: return [3 /*break*/, 12];
                            case 11:
                                resolve(true);
                                _a.label = 12;
                            case 12: return [2 /*return*/];
                        }
                    });
                }); })];
        });
    });
}
// async function waitWrapper(file: string): Promise<boolean> {
//   await processEmlFile(file);
//   return new Promise((resolve) => { resolve(true) });
// }
var finder = find(dirInput);
console.log("Collecting EML files...");
finder.on('file', function (file) {
    files.add(path.normalize(file));
});
finder.on('end', function () {
    var e_1, _a;
    var p = [];
    try {
        for (var _b = __values(files.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
            var file = _c.value;
            p.push(processEmlFile(file));
        }
    }
    catch (e_1_1) { e_1 = { error: e_1_1 }; }
    finally {
        try {
            if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
        }
        finally { if (e_1) throw e_1.error; }
    }
    Promise.all(p)
        .then(function () {
        stringifierMeta.end();
        stringifierTerms.end();
        stringifierPersons.end();
        console.log("Processing finished.");
    });
});
//# sourceMappingURL=index.js.map