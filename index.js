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
var simpleParser = require('mailparser').simpleParser;
var stringify = require('csv-stringify');
var tm = require('text-miner');
var nGram = require('n-gram');
var dirInput = '/Users/ralfbecher/Documents/Daten/maildir';
// const dirInput = '/Users/ralfbecher/Documents/Daten/test';
var dirOutput = './output';
var csvFileMeta = 'maildir.csv';
var csvFileTerms = 'mailterms.csv';
var csvFileNGrams = 'mailngrams.csv';
var csvFilePersons = 'mailpersons.csv';
var outputFieldsMeta = ['File', 'ID', 'Date', 'Subject', 'FromEmail', 'FromName', 'ToEmails', 'CCEmails', 'BCCEmails', 'Text', 'TextLength', 'Hash', 'Attachments'];
var outputFieldsTerms = ['Hash', 'Term', 'Count'];
var outputFieldsNGrams = ['Hash', 'nGram', 'Type', 'Count'];
var outputFieldsPersons = ['ID', 'Email', 'Role'];
var files = new Set([]);
var hashes = new Set([]);
if (!fs.existsSync(dirOutput)) {
    fs.mkdirSync(dirOutput);
}
var outputMeta = fs.createWriteStream(dirOutput + '/' + csvFileMeta, { flags: 'w' });
var outputTerms = fs.createWriteStream(dirOutput + '/' + csvFileTerms, { flags: 'w' });
var outputNGrams = fs.createWriteStream(dirOutput + '/' + csvFileNGrams, { flags: 'w' });
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
var stringifierNGrams = stringify({
    delimiter: ',',
    header: true,
    quoted: true,
    columns: outputFieldsNGrams
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
    stringifierNGrams.pipe(outputNGrams);
    stringifierPersons.pipe(outputPersons);
}
startPipe();
function nGramAggr(biGrams, triGrams) {
    var e_1, _a;
    var res = [];
    var gramMap = new Map();
    var nGram = biGrams.concat(triGrams);
    nGram.forEach(function (e, i) {
        var t = e.length;
        var k = e.join(' ');
        if (!gramMap.has(k)) {
            gramMap.set(k, [t, 1]);
        }
        else {
            var n = gramMap.get(k);
            gramMap.set(k, [n[0], n[1] + 1]);
        }
    });
    try {
        for (var _b = __values(gramMap.entries()), _c = _b.next(); !_c.done; _c = _b.next()) {
            var e = _c.value;
            res.push([e[0], e[1][0].toString(), e[1][1].toString()]);
        }
    }
    catch (e_1_1) { e_1 = { error: e_1_1 }; }
    finally {
        try {
            if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
        }
        finally { if (e_1) throw e_1.error; }
    }
    return res;
}
function mineText(hash, text) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) {
                    if (!hashes.has(hash)) {
                        hashes.add(hash);
                        text = text.replace(/([^\x20-\uD7FF\uE000-\uFFFC\u{10000}-\u{10FFFF}])/ug, ' ').replace(/[:#_\\\/\t\(\)\[\]\*\$\+"{}<>=~&%`´]/g, ' ');
                        var corpus = new tm.Corpus([text]);
                        var cleansedCorpus = corpus.clean()
                            .trim()
                            .removeInterpunctuation()
                            .removeInvalidCharacters()
                            .removeDigits()
                            .removeNewlines()
                            .clean()
                            .toLower()
                            .removeWords(tm.STOPWORDS.EN)
                            .stem('Porter');
                        var docs = cleansedCorpus.documents;
                        var list = docs[0].text ? docs[0].text.split(' ') : [];
                        var biGrams = nGram.bigram(list);
                        var triGrams = nGram.trigram(list);
                        var terms_1 = new tm.DocumentTermMatrix(cleansedCorpus);
                        terms_1.vocabulary.forEach(function (e, i) { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0: return [4 /*yield*/, writeRow(stringifierTerms, [hash, e, terms_1.data[0][i]])];
                                    case 1:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); });
                        nGramAggr(biGrams, triGrams).forEach(function (e) { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0: return [4 /*yield*/, writeRow(stringifierNGrams, [hash, e[0], e[1], e[2]])];
                                    case 1:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); });
                        resolve(docs[0].text);
                    }
                    else {
                        resolve('');
                    }
                })];
        });
    });
}
function cleanUpEmailsAndNames(str) {
    return str ? str.trim().replace(/[\r\n|\r|\n'"]+/g, '') : '';
}
function switchNameParts(str) {
    if (str && str.indexOf('@') === -1 && str.indexOf(',') !== -1) {
        var parts = str.split(',');
        if (parts.length === 2) {
            str = parts[1].trim() + ' ' + parts[0].trim();
        }
    }
    return str;
}
function transposeMails(id, mails, role) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) {
                    if (mails.length === 0) {
                        resolve(true);
                    }
                    else {
                        mails.forEach(function (e) { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0: return [4 /*yield*/, writeRow(stringifierPersons, [id, e, role])];
                                    case 1:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); });
                        resolve(true);
                    }
                })];
        });
    });
}
function writeRow(stream, row) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) { return __awaiter(_this, void 0, void 0, function () {
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0: return [4 /*yield*/, stream.write(row)];
                            case 1:
                                _a.sent();
                                resolve(true);
                                return [2 /*return*/];
                        }
                    });
                }); })];
        });
    });
}
function headers2obj(headers) {
    var e_2, _a;
    var res = {};
    try {
        for (var _b = __values(headers.entries()), _c = _b.next(); !_c.done; _c = _b.next()) {
            var e = _c.value;
            res[e[0]] = e[1];
        }
    }
    catch (e_2_1) { e_2 = { error: e_2_1 }; }
    finally {
        try {
            if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
        }
        finally { if (e_2) throw e_2.error; }
    }
    return res;
}
function getNameFromEmailAddress(email) {
    var result = '';
    var regex = /^(.*?)(\s*\<(.*?)\>)$/g;
    var match = regex.exec(email);
    if (match) {
        var name_1 = match[1].replace(/"/g, "").trim();
        if (name_1 && name_1.length) {
            result = name_1;
        }
    }
    return result;
}
function convertEml(fileName, eml) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                    var email, data, headers, error_1;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                email = {
                                    file: '',
                                    messageId: '',
                                    dateSent: new Date(),
                                    subject: '',
                                    fromEmail: '',
                                    fromName: '',
                                    toEmails: [],
                                    ccEmails: [],
                                    bccEmails: [],
                                    text: '',
                                    textHash: '',
                                    attachments: []
                                };
                                _a.label = 1;
                            case 1:
                                _a.trys.push([1, 3, , 4]);
                                return [4 /*yield*/, simpleParser(eml)];
                            case 2:
                                data = _a.sent();
                                headers = headers2obj(data.headers);
                                email.file = fileName;
                                email.messageId = data.messageId || '';
                                email.dateSent = data.date;
                                email.subject = data.subject;
                                if (data.text) {
                                    email.textHash = require('crypto').createHash('md5').update(data.text).digest('base64');
                                    email.text = data.text;
                                }
                                else {
                                    email.textHash = '0000'; // placeholder for empty text#
                                    email.text = '';
                                }
                                email.fromEmail = cleanUpEmailsAndNames(data.from && data.from.text ? data.from.text : '');
                                email.fromName = switchNameParts(getNameFromEmailAddress(cleanUpEmailsAndNames(headers['x-from'] || '')));
                                email.toEmails = data.to && data.to.text ? cleanUpEmailsAndNames(data.to.text).split(',').map(function (e) { return e.trim(); }) : [];
                                email.ccEmails = headers['cc'] ? headers['cc'].text.split(',').map(function (e) { return e.trim(); }) : [];
                                email.bccEmails = headers['bcc'] ? headers['bcc'].text.split(',').map(function (e) { return e.trim(); }) : [];
                                if (data.attachments && data.attachments.length > 0) {
                                    email.attachments = data.attachments
                                        .filter(function (e) { return e.contentDisposition === 'attachment'; })
                                        .map(function (e) { return e.filename; });
                                }
                                resolve(email);
                                return [3 /*break*/, 4];
                            case 3:
                                error_1 = _a.sent();
                                console.error(error_1);
                                resolve(email);
                                return [3 /*break*/, 4];
                            case 4: return [2 /*return*/];
                        }
                    });
                }); })];
        });
    });
}
function processEmlFile(emlFile) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) { return __awaiter(_this, void 0, void 0, function () {
                    var eml, email, dateSent, cleanText, csvRowMeta, error_2;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                eml = fs.readFileSync(emlFile, 'utf-8');
                                if (!(eml.substr(0, 11) === 'Message-ID:' || eml.substr(0, 14) === 'Received: from')) return [3 /*break*/, 10];
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
                                _a.label = 2;
                            case 2:
                                _a.trys.push([2, 8, , 9]);
                                return [4 /*yield*/, mineText(email.textHash, email.text)];
                            case 3:
                                cleanText = _a.sent();
                                csvRowMeta = [
                                    email.file,
                                    email.messageId,
                                    dateSent,
                                    email.subject || '',
                                    email.fromEmail,
                                    email.fromName,
                                    email.toEmails.join(','),
                                    email.ccEmails.join(','),
                                    email.bccEmails.join(','),
                                    cleanText,
                                    email.text.length.toString(),
                                    email.textHash,
                                    email.attachments.join(',')
                                ];
                                return [4 /*yield*/, writeRow(stringifierMeta, csvRowMeta)];
                            case 4:
                                _a.sent();
                                // await transposeMails(email.messageId, [email.fromEmail], 'From');
                                return [4 /*yield*/, transposeMails(email.messageId, email.toEmails, 'To')];
                            case 5:
                                // await transposeMails(email.messageId, [email.fromEmail], 'From');
                                _a.sent();
                                return [4 /*yield*/, transposeMails(email.messageId, email.ccEmails, 'Cc')];
                            case 6:
                                _a.sent();
                                return [4 /*yield*/, transposeMails(email.messageId, email.bccEmails, 'Bcc')];
                            case 7:
                                _a.sent();
                                resolve(true);
                                return [3 /*break*/, 9];
                            case 8:
                                error_2 = _a.sent();
                                console.error(error_2);
                                process.exit(1);
                                return [3 /*break*/, 9];
                            case 9: return [3 /*break*/, 11];
                            case 10:
                                resolve(true);
                                _a.label = 11;
                            case 11: return [2 /*return*/];
                        }
                    });
                }); })];
        });
    });
}
var finder = find(dirInput);
console.log("Collecting EML files...");
finder.on('file', function (file) {
    files.add(path.normalize(file));
});
function processAll() {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            console.log("Collecting EML files finished...");
            return [2 /*return*/, new Promise(function (resolve) { return __awaiter(_this, void 0, void 0, function () {
                    var e_3, _a, _b, _c, file, e_3_1;
                    return __generator(this, function (_d) {
                        switch (_d.label) {
                            case 0:
                                _d.trys.push([0, 5, 6, 7]);
                                _b = __values(files.values()), _c = _b.next();
                                _d.label = 1;
                            case 1:
                                if (!!_c.done) return [3 /*break*/, 4];
                                file = _c.value;
                                return [4 /*yield*/, processEmlFile(file)];
                            case 2:
                                _d.sent();
                                _d.label = 3;
                            case 3:
                                _c = _b.next();
                                return [3 /*break*/, 1];
                            case 4: return [3 /*break*/, 7];
                            case 5:
                                e_3_1 = _d.sent();
                                e_3 = { error: e_3_1 };
                                return [3 /*break*/, 7];
                            case 6:
                                try {
                                    if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
                                }
                                finally { if (e_3) throw e_3.error; }
                                return [7 /*endfinally*/];
                            case 7:
                                stringifierMeta.end();
                                stringifierTerms.end();
                                stringifierNGrams.end();
                                stringifierPersons.end();
                                console.log("Processing finished.");
                                resolve(true);
                                return [2 /*return*/];
                        }
                    });
                }); })];
        });
    });
}
finder.on('end', function () {
    processAll();
});
//# sourceMappingURL=index.js.map