/* Copyright (C) 2017-2025 Tuumik Systems OÜ */

import ExcelJS from 'exceljs';
import dayjs from 'dayjs';
import utc from 'dayjs/plugin/utc.js';

dayjs.extend(utc);

export async function xlsx1(req, res) {
  // API key
  const apiKeyFromHeader = req.headers['x-api-key'];
  const apiKey = process.env.API_KEY;
  if (apiKeyFromHeader !== apiKey) return res.status(403).json({ message: 'Forbidden: Invalid API Key' });
  // /API key

  // translation strings
  const ws1Name = 'Tasks';
  const ws2Name = 'People';
  const trGeneralTitle = 'Annex to Invoice No';
  const trHighlightChanges = 'Highlight changes:';
  const trTopDuration = 'DURATION';
  const trTopRate = 'RATE';
  const trTopSum = 'SUM';
  const trTopSourceData = 'SOURCE DATA';
  const trRowRef = '#';
  const trDate = 'Date';
  const trClient = 'Client';
  const trProject = 'Project';
  const trUser = 'User';
  const trTask = 'Task';
  const trDuration = 'Duration';
  const trRate = 'Rate';
  const trSum = 'Sum';
  const trDurOrig = 'Original';
  const trMinutes = 'Minutes';
  const trApplied = 'Applied';
  const trChange = 'Change';
  const trUserMulti = 'User X';
  const trRowMulti = 'Row X';
  const trAdd = '+/-';
  const trRateOrig = 'Original';
  const trPeopleOrigName = 'Original Name';
  const trPeopleAppliedName = 'Applied Name';
  const trPeopleOrigRate = 'Original Rate';
  const trPeopleAppliedRate = 'Applied Rate';
  const trPeopleDurMulti = 'Duration X';
  const trPeopleRateMulti = 'Rate X';
  const trSumOrig = 'Original';
  const trClientProject = 'Client/project:';
  const trClientProjectAny = 'Any';
  const trPeriod = 'Period:';
  const trTotalDurTitle = 'Total duration:';
  const trTotalSumTitle = 'Total sum:';
  // /translation strings

  const data = req.body;

  const wb = new ExcelJS.Workbook();
  const ws1 = wb.addWorksheet(ws1Name);
  const ws2 = wb.addWorksheet(ws2Name);
  const times = data.times;

  // convert dates (js timestamp) into excel timestamp
  const times2 = times.map(time => {
    const timestampUnix = Math.floor(time.date / 1000);
    const timestampExcel = (timestampUnix / 86400) + 25569;
    // 86400 - seconds in a day
    // 25569 - days between 1 Jan 1970 (Unix) and 1 Jan 1900 (Excel)
    const timeDoc = time;
    timeDoc.timestampExcel = timestampExcel;
    return timeDoc;
  });
  // /convert dates (js timestamp) into excel timestamp

  // owners
  const uniqueOwners = [];
  const ownerIds = [];
  times2.forEach(time => {
    if (!ownerIds.includes(time.owner)) {
      uniqueOwners.push({ _id: time.owner, name: time.ownerName });
      ownerIds.push(time.owner);
    }
  });
  // /owners

  // general font
  const font = { name: 'Arial', family: 2 };
  for (let i = 1; i < 60; i += 1) {
    ws1.getColumn(i).font = font;
    ws2.getColumn(i).font = font;
  }
  // /general font

  // general style
  const genBorder = {
    top: { style: 'thin', color: { argb: 'FF9F9F9F' } },
    bottom: { style: 'thin', color: { argb: 'FF9F9F9F' } },
  };
  const modFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFAF1A2' },
  };
  const modBorder = {
    top: { style: 'thin', color: { argb: 'FF9F9F9F' } },
    left: { style: 'thin', color: { argb: 'FF9F9F9F' } },
    bottom: { style: 'thin', color: { argb: 'FF9F9F9F' } },
    right: { style: 'thin', color: { argb: 'FF9F9F9F' } },
  };
  const modRateOrigFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFB4C7DC' },
  };
  const currencyFormat = '#,##0.00 [$€-425];[RED]-#,##0.00 [$€-425]';
  const durFormat = '[h]"h" mm"m"';
  // /general style

  // divider style
  const divider1Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF77BC65' },
  };
  const divider2Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF77BC65' },
  };
  const divider3Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF77BC65' },
  };
  const divider4Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF77BC65' },
  };
  const divider5Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF77BC65' },
  };
  // /divider style

  // columns
  const colRowRef = 1;
  const colDate = 2;
  const colClient = 3;
  const colProject = 4;
  const colUser = 5;
  const colTask = 6;
  const colDuration = 7;
  const colRate = 8;
  const colSum = 9;
  const colDivider1 = 10;
  const colDurOrigMinutes = 11;
  const colDurOrigDisplay = 12;
  const colDurAppliedMinutes = 13;
  const colDurAppliedDisplay = 14;
  const colDurChange = 15;
  const colDurUserMulti = 16;
  const colDurRowMulti = 17;
  const colDurAdd = 18;
  const colDivider2 = 19;
  const colRateOrig = 20;
  const colRateApplied = 21;
  const colRateChange = 22;
  const colRateUserMulti = 23;
  const colRateRowMulti = 24;
  const colRateAdd = 25;
  const colDivider3 = 26;
  const colSumOrig = 27;
  const colSumApplied = 28;
  const colSumChange = 29;
  const colDivider4 = 30;
  const colSrcDate = 31;
  const colSrcClient = 32;
  const colSrcProject = 33;
  const colSrcUser = 34;
  const colSrcTask = 35;
  const colSrcDur = 36;
  const colSrcRate = 37;
  const colSrcSum = 38;
  const colDivider5 = 39;
  const colSrcIntCom = 40;
  const colRowRefLetter = 'A';
  const colDateLetter = 'B';
  const colClientLetter = 'C';
  const colProjectLetter = 'D';
  const colUserLetter = 'E';
  const colTaskLetter = 'F';
  const colDurLetter = 'G';
  const colRateLetter = 'H';
  const colSumLetter = 'I';
  const colDurOrigMinutesLetter = 'K';
  const colDurOrigDisplayLetter = 'L';
  const colDurAppliedMinutesLetter = 'M';
  const colDurAppliedDisplayLetter = 'N';
  const colDurChangeLetter = 'O';
  const colDurRowMultiLetter = 'Q';
  const colDurAddLetter = 'R';
  const colRateOrigLetter = 'T';
  const colRateAppliedLetter = 'U';
  const colRateChangeLetter = 'V';
  const colRateUserMultiLetter = 'W';
  const colRateRowMultiLetter = 'X';
  const colRateAddLetter = 'Y';
  const colSumOrigLetter = 'AA';
  const colSumAppliedLetter = 'AB';
  const colSumChangeLetter = 'AC';
  const colSrcDateLetter = 'AE';
  const colSrcClientLetter = 'AF';
  const colSrcProjectLetter = 'AG';
  const colSrcUserLetter = 'AH';
  const colSrcTaskLetter = 'AI';
  const colSrcDurLetter = 'AJ';
  const colSrcRateLetter = 'AK';
  const colSrcSumLetter = 'AL';
  const mainCols = [1, 2, 3, 4, 5, 6, 7, 8, 9];
  const mainColsLeft = [1, 2, 3, 4, 5, 6];
  const mainColsRight = [7, 8, 9];
  const durEditingCols = [11, 12, 13, 14, 15, 16, 17, 18];
  const rateEditingCols = [20, 21, 22, 23, 24, 25];
  const sumEditingCols = [27, 28, 29];
  const srcCols = [31, 32, 33, 34, 35, 36, 37, 38];
  const srcColsLeft = [31, 32, 33, 34, 35];
  const srcColsRight = [36, 37, 38];
  // /columns

  // column widths
  const rowRefWidth = 8;
  const dateWidth = 12;
  const clientWidth = 22;
  const projectWidth = 40;
  const userWidth = 22;
  const durationWidth = 12;
  const rateWidth = 12;
  const sumWidth = 12;
  ws1.getColumn(colRowRef).width = rowRefWidth;
  ws1.getColumn(colDate).width = dateWidth;
  ws1.getColumn(colClient).width = clientWidth;
  ws1.getColumn(colProject).width = projectWidth;
  ws1.getColumn(colUser).width = userWidth;
  ws1.getColumn(colDuration).width = durationWidth;
  ws1.getColumn(colRate).width = rateWidth;
  ws1.getColumn(colSum).width = sumWidth;
  ws1.getColumn(colDivider1).width = 2;
  ws1.getColumn(colDurOrigMinutes).width = 8;
  ws1.getColumn(colDurOrigDisplay).width = 12;
  ws1.getColumn(colDurAppliedMinutes).width = 8;
  ws1.getColumn(colDurAppliedDisplay).width = 12;
  ws1.getColumn(colDurChange).width = 8;
  ws1.getColumn(colDurUserMulti).width = 8;
  ws1.getColumn(colDurRowMulti).width = 8;
  ws1.getColumn(colDurAdd).width = 8;
  ws1.getColumn(colDivider2).width = 2;
  ws1.getColumn(colRateOrig).width = 12;
  ws1.getColumn(colRateApplied).width = 12;
  ws1.getColumn(colRateChange).width = 12;
  ws1.getColumn(colRateUserMulti).width = 8;
  ws1.getColumn(colRateRowMulti).width = 8;
  ws1.getColumn(colRateAdd).width = 8;
  ws1.getColumn(colDivider3).width = 2;
  ws1.getColumn(colSumApplied).width = sumWidth;
  ws1.getColumn(colSumChange).width = sumWidth;
  ws1.getColumn(colDivider4).width = 2;
  ws1.getColumn(colSrcDate).width = dateWidth;
  ws1.getColumn(colSrcClient).width = clientWidth;
  ws1.getColumn(colSrcProject).width = projectWidth;
  ws1.getColumn(colSumOrig).width = sumWidth;
  ws1.getColumn(colSrcUser).width = userWidth;
  ws1.getColumn(colSrcTask).width = 70;
  ws1.getColumn(colSrcDur).width = durationWidth;
  ws1.getColumn(colSrcRate).width = rateWidth;
  ws1.getColumn(colSrcSum).width = sumWidth;
  ws1.getColumn(colDivider5).width = 2;
  ws1.getColumn(colSrcIntCom).width = 25;
  // /column widths

  // set dynamic width for task column
  const sf = data.exportOptions.showFields;
  let widthRemaining = 160; // columns 1-8 should always be this total width
  if (sf.rowRef) widthRemaining = widthRemaining - rowRefWidth;
  if (sf.date) widthRemaining = widthRemaining - dateWidth;
  if (sf.client) widthRemaining = widthRemaining - clientWidth;
  if (sf.project) widthRemaining = widthRemaining - projectWidth;
  if (sf.user) widthRemaining = widthRemaining - userWidth;
  if (sf.duration) widthRemaining = widthRemaining - durationWidth;
  if (sf.rate) widthRemaining = widthRemaining - rateWidth;
  if (sf.sum) widthRemaining = widthRemaining - sumWidth;
  ws1.getColumn(colTask).width = widthRemaining;
  // /set dynamic width for task column

  // hide columns
  if (!sf.rowRef) ws1.getColumn(colRowRef).hidden = true;
  if (!sf.date) ws1.getColumn(colDate).hidden = true;
  if (!sf.client) ws1.getColumn(colClient).hidden = true;
  if (!sf.project) ws1.getColumn(colProject).hidden = true;
  if (!sf.user) ws1.getColumn(colUser).hidden = true;
  if (!sf.task) ws1.getColumn(colTask).hidden = true;
  if (!sf.duration) ws1.getColumn(colDuration).hidden = true;
  if (!sf.rate) ws1.getColumn(colRate).hidden = true;
  if (!sf.sum) ws1.getColumn(colSum).hidden = true;
  ws1.getColumn(colDurOrigMinutes).hidden = true;
  ws1.getColumn(colDurAppliedMinutes).hidden = true;
  // /hide columns

  let currentRow = 0;

  // top logo
  const logoRow = ws1.addRow([]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  logoRow.height = 55;
  const imageId1 = wb.addImage({
    filename: 'src/assets/top.png',
    extension: 'png',
  });
  ws1.addImage(imageId1, {
    tl: { col: 0, row: 0 },
    ext: { width: 50, height: 50 },
    editAs: 'oneCell',
  });
  // /top logo

  // top info 1: document title
  const titleRowContent = [];
  titleRowContent[1] = trGeneralTitle;
  const titleRow = ws1.addRow(titleRowContent);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  const titleRowFont = { name: 'Arial', family: 2, size: 20, bold: true };
  titleRow.getCell(1).font = titleRowFont;
  titleRow.height = 30;
  // /top info 1: document title

  // tenant name row
  const tenantRow = ws1.addRow([data.tenant.name]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  const tenantRowFont = { name: 'Arial', family: 2, bold: true };
  tenantRow.getCell(1).font = tenantRowFont;
  // /tenant name row

  // spacer row
  ws1.addRow([]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  // /spacer row

  // top info 2: client and project title
  ws1.addRow([trClientProject]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  // /top info 2: client and project title

  // top info 2: client and project list
  if (data.meta.clients) {
    const cpList = [];
    data.meta.clients.forEach(client => {
      client.projects.forEach(project => {
        cpList.push({ client: client.name, project: project.name });
      });
    });
    cpList.forEach((cliPro, index) => {
      const cpContent = [];
      cpContent[1] = `${cliPro.client} - ${cliPro.project}`;
      const cpRow = ws1.addRow(cpContent);
      currentRow = currentRow + 1;
      ws1.mergeCells(`A${currentRow}:I${currentRow}`);
      const cpRowFont = { name: 'Arial', family: 2, bold: true };
      cpRow.getCell(1).font = cpRowFont;
    });
  } else {
    const cpContent = [];
    cpContent[1] = trClientProjectAny;
    const cpRow = ws1.addRow(cpContent);
    currentRow = currentRow + 1;
    ws1.mergeCells(`A${currentRow}:I${currentRow}`);
    const cpRowFont = { name: 'Arial', family: 2, bold: true };
    cpRow.getCell(1).font = cpRowFont;
  }
  // /top info 2: client and project list

  // spacer row
  ws1.addRow([]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  // /spacer row

  // top info 3: period title
  ws1.addRow([trPeriod]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  // /top info 3: period title

  // top info 3: period string
  const startStr = data.meta.period.start ? dayjs.utc(data.meta.period.start).format(data.tenant.dateFormat) : '*';
  const endStr = data.meta.period.end ? dayjs.utc(data.meta.period.end).format(data.tenant.dateFormat) : '*';
  const periodStr = `${startStr} - ${endStr}`;
  const periodRowContent = [];
  periodRowContent[1] = periodStr;
  const periodRow = ws1.addRow(periodRowContent);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  const periodRowFont = { name: 'Arial', family: 2, bold: true };
  periodRow.getCell(1).font = periodRowFont;
  // /top info 3: period string

  // diff toggle
  const diffToggleLabel = ws1.getCell('K1');
  diffToggleLabel.value = trHighlightChanges;
  ws1.mergeCells(`K1:N1`);
  const diffToggleTarget = ws1.getCell('O1');
  diffToggleTarget.value = 'YES';
  const toggleLabelFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFEEEEEE' },
  };
  const toggleLabelAlignment = { vertical: 'middle', horizontal: 'center' };
  diffToggleLabel.fill = toggleLabelFill;
  diffToggleLabel.alignment = toggleLabelAlignment;
  const toggleTargetFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFDDDDDD' },
  };
  const toggleTargetAlignment = { vertical: 'middle', horizontal: 'center' };
  diffToggleTarget.fill = toggleTargetFill;
  diffToggleTarget.alignment = toggleTargetAlignment;
  const diffToggle = '$O$1';
  // /diff toggle

  const topInfoRowCount = currentRow;

  // top edit field headers
  const fhRowContent = [];
  fhRowContent[colDurOrigMinutes] = trTopDuration;
  fhRowContent[colRateOrig] = trTopRate;
  fhRowContent[colSumOrig] = trTopSum;
  fhRowContent[colSrcDate] = trTopSourceData;
  const fhRow = ws1.addRow(fhRowContent);
  currentRow = currentRow + 1;
  const eFont = { name: 'Arial', family: 2, bold: true };
  const eFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFDDDDDD' },
  };
  const eAlignment = { vertical: 'middle', horizontal: 'center' };
  durEditingCols.forEach(x => {
    fhRow.getCell(x).font = eFont;
    fhRow.getCell(x).fill = eFill;
    fhRow.getCell(x).alignment = eAlignment;
  });
  rateEditingCols.forEach(x => {
    fhRow.getCell(x).font = eFont;
    fhRow.getCell(x).fill = eFill;
    fhRow.getCell(x).alignment = eAlignment;
  });
  sumEditingCols.forEach(x => {
    fhRow.getCell(x).font = eFont;
    fhRow.getCell(x).fill = eFill;
    fhRow.getCell(x).alignment = eAlignment;
  });
  srcCols.forEach(x => {
    fhRow.getCell(x).font = eFont;
    fhRow.getCell(x).fill = eFill;
    fhRow.getCell(x).alignment = eAlignment;
  });
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  ws1.mergeCells(`${colDurOrigMinutesLetter}${currentRow}:${colDurAddLetter}${currentRow}`);
  ws1.mergeCells(`${colRateOrigLetter}${currentRow}:${colRateAddLetter}${currentRow}`);
  ws1.mergeCells(`${colSumOrigLetter}${currentRow}:${colSumChangeLetter}${currentRow}`);
  ws1.mergeCells(`${colSrcDateLetter}${currentRow}:${colSrcSumLetter}${currentRow}`);
  fhRow.getCell(colDivider1).fill = divider1Fill;
  fhRow.getCell(colDivider2).fill = divider2Fill;
  fhRow.getCell(colDivider3).fill = divider3Fill;
  fhRow.getCell(colDivider4).fill = divider4Fill;
  fhRow.getCell(colDivider5).fill = divider5Fill;
  // /top edit field headers

  // times heading 1
  const th1Content = [];
  th1Content[colRowRef] = trRowRef;
  th1Content[colDate] = trDate;
  th1Content[colClient] = trClient;
  th1Content[colProject] = trProject;
  th1Content[colUser] = trUser;
  th1Content[colTask] = trTask;
  th1Content[colDuration] = trDuration;
  th1Content[colRate] = trRate;
  th1Content[colSum] = trSum;
  th1Content[colDurOrigMinutes] = trDurOrig;
  th1Content[colDurOrigDisplay] = trDurOrig;
  th1Content[colDurAppliedMinutes] = trApplied;
  th1Content[colDurAppliedDisplay] = trApplied;
  th1Content[colDurChange] = trChange;
  th1Content[colDurUserMulti] = trUserMulti;
  th1Content[colDurRowMulti] = trRowMulti;
  th1Content[colDurAdd] = trAdd;
  th1Content[colRateOrig] = trRateOrig;
  th1Content[colRateApplied] = trApplied;
  th1Content[colRateChange] = trChange;
  th1Content[colRateUserMulti] = trUserMulti;
  th1Content[colRateRowMulti] = trRowMulti;
  th1Content[colRateAdd] = trAdd;
  th1Content[colSumOrig] = trSumOrig;
  th1Content[colSumApplied] = trApplied;
  th1Content[colSumChange] = trChange;
  th1Content[colSrcDate] = trDate;
  th1Content[colSrcClient] = trClient;
  th1Content[colSrcProject] = trProject;
  th1Content[colSrcUser] = trUser;
  th1Content[colSrcTask] = trTask;
  th1Content[colSrcDur] = trDuration;
  th1Content[colSrcRate] = trRate;
  th1Content[colSrcSum] = trSum;
  const th1Row = ws1.addRow(th1Content);
  currentRow = currentRow + 1;
  const th1RowNumber = currentRow;
  const h1Font = { name: 'Arial', family: 2, bold: true };
  const h1Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC9C9C9' } };
  const h1Alignment1 = { vertical: 'top', horizontal: 'left' };
  const h1Alignment2 = { vertical: 'top', horizontal: 'right' };
  mainColsLeft.forEach(x => {
    th1Row.getCell(x).font = h1Font;
    th1Row.getCell(x).fill = h1Fill;
    th1Row.getCell(x).alignment = h1Alignment1;
  });
  mainColsRight.forEach(x => {
    th1Row.getCell(x).font = h1Font;
    th1Row.getCell(x).fill = h1Fill;
    th1Row.getCell(x).alignment = h1Alignment2;
  });
  const h1Fill2 = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFC9C9C9' },
  };
  durEditingCols.forEach(x => {
    th1Row.getCell(x).font = h1Font;
    th1Row.getCell(x).fill = h1Fill2;
    th1Row.getCell(x).alignment = h1Alignment2;
  });
  rateEditingCols.forEach(x => {
    th1Row.getCell(x).font = h1Font;
    th1Row.getCell(x).fill = h1Fill2;
    th1Row.getCell(x).alignment = h1Alignment2;
  });
  sumEditingCols.forEach(x => {
    th1Row.getCell(x).font = h1Font;
    th1Row.getCell(x).fill = h1Fill2;
    th1Row.getCell(x).alignment = h1Alignment2;
  });
  srcColsLeft.forEach(x => {
    th1Row.getCell(x).font = h1Font;
    th1Row.getCell(x).fill = h1Fill;
    th1Row.getCell(x).alignment = h1Alignment1;
  });
  srcColsRight.forEach(x => {
    th1Row.getCell(x).font = h1Font;
    th1Row.getCell(x).fill = h1Fill;
    th1Row.getCell(x).alignment = h1Alignment2;
  });
  th1Row.getCell(colDivider1).fill = divider1Fill;
  th1Row.getCell(colDivider2).fill = divider2Fill;
  th1Row.getCell(colDivider3).fill = divider3Fill;
  th1Row.getCell(colDivider4).fill = divider4Fill;
  th1Row.getCell(colDivider5).fill = divider5Fill;
  // /times heading 1

  // times heading 2
  const th2Content = [];
  th2Content[colRate] = `${data.tenant.currency.str}/h`;
  th2Content[colSum] = `${data.tenant.currency.str}`;
  th2Content[colDurChange] = trMinutes;
  th2Content[colDurAdd] = trMinutes;
  th2Content[colRateOrig] = `${data.tenant.currency.str}/h`;
  th2Content[colRateApplied] = `${data.tenant.currency.str}/h`;
  th2Content[colRateChange] = `${data.tenant.currency.str}/h`;
  th2Content[colRateAdd] = `${data.tenant.currency.str}/h`;
  th2Content[colSumOrig] = `${data.tenant.currency.str}`;
  th2Content[colSumApplied] = `${data.tenant.currency.str}`;
  th2Content[colSumChange] = `${data.tenant.currency.str}`;
  th2Content[colSrcRate] = `${data.tenant.currency.str}/h`;
  th2Content[colSrcSum] = `${data.tenant.currency.str}`;
  const th2Row = ws1.addRow(th2Content);
  currentRow = currentRow + 1;
  const th2RowNumber = currentRow;
  const h2Font = { name: 'Arial', family: 2, bold: true, size: 9, color: { argb: 'FF777777' } };
  const h2Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFC9C9C9' },
  };
  const h2Alignment1 = { vertical: 'top', horizontal: 'left' };
  const h2Alignment2 = { vertical: 'top', horizontal: 'right' };
  mainColsLeft.forEach(x => {
    th2Row.getCell(x).font = h2Font;
    th2Row.getCell(x).fill = h2Fill;
    th2Row.getCell(x).alignment = h2Alignment1;
  });
  mainColsRight.forEach(x => {
    th2Row.getCell(x).font = h2Font;
    th2Row.getCell(x).fill = h2Fill;
    th2Row.getCell(x).alignment = h2Alignment2;
  });
  const h2Fill2 = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFC9C9C9' },
  };
  durEditingCols.forEach(x => {
    th2Row.getCell(x).font = h2Font;
    th2Row.getCell(x).fill = h2Fill2;
    th2Row.getCell(x).alignment = h2Alignment2;
  });
  rateEditingCols.forEach(x => {
    th2Row.getCell(x).font = h2Font;
    th2Row.getCell(x).fill = h2Fill2;
    th2Row.getCell(x).alignment = h2Alignment2;
  });
  sumEditingCols.forEach(x => {
    th2Row.getCell(x).font = h2Font;
    th2Row.getCell(x).fill = h2Fill2;
    th2Row.getCell(x).alignment = h2Alignment2;
  });
  srcColsLeft.forEach(x => {
    th2Row.getCell(x).font = h2Font;
    th2Row.getCell(x).fill = h2Fill;
    th2Row.getCell(x).alignment = h2Alignment1;
  });
  srcColsRight.forEach(x => {
    th2Row.getCell(x).font = h2Font;
    th2Row.getCell(x).fill = h2Fill;
    th2Row.getCell(x).alignment = h2Alignment2;
  });
  th2Row.getCell(colDivider1).fill = divider1Fill;
  th2Row.getCell(colDivider2).fill = divider2Fill;
  th2Row.getCell(colDivider3).fill = divider3Fill;
  th2Row.getCell(colDivider4).fill = divider4Fill;
  th2Row.getCell(colDivider5).fill = divider5Fill;
  // /times heading 2

  // times content
  times2.forEach((time, index) => {
    const content = [];
    const ownerIndex = uniqueOwners.findIndex(
      user => user._id === time.owner,
    );
    content[colRowRef] = '';
    content[colDate] = time.timestampExcel;
    content[colClient] = time.clientName;
    content[colProject] = time.projectName;
    content[colUser] = '';
    content[colTask] = time.taskDesc;
    content[colDuration] = '';
    content[colRate] = '';
    content[colSum] = '';
    content[colDurOrigMinutes] = time.endMinute - time.startMinute;
    content[colDurOrigDisplay] = '';
    content[colDurAppliedMinutes] = '';
    content[colDurAppliedDisplay] = '';
    content[colDurChange] = '';
    content[colDurUserMulti] = '';
    content[colDurRowMulti] = 1;
    content[colDurAdd] = 0;
    content[colRateOrig] = '';
    content[colRateApplied] = '';
    content[colRateChange] = '';
    content[colRateUserMulti] = '';
    content[colRateRowMulti] = 1;
    content[colRateAdd] = 0;
    content[colSumOrig] = '';
    content[colSumApplied] = '';
    content[colSumChange] = '';
    content[colSrcDate] = time.timestampExcel;
    content[colSrcClient] = time.clientName;
    content[colSrcProject] = time.projectName;
    content[colSrcUser] = time.ownerName;
    content[colSrcTask] = time.taskDesc;
    content[colSrcDur] = '';
    content[colSrcRate] = '';
    content[colSrcSum] = '';
    content[colSrcIntCom] = time.intCom || '';
    const row = ws1.addRow(content);
    currentRow = currentRow + 1;
    row.getCell(colRowRef).value = {
      formula: `ROW()-${th2RowNumber}`,
    };
    row.getCell(colUser).value = {
      formula: `${ws2Name}!B${ownerIndex + 2}`,
    };
    row.getCell(
      colDate,
    ).numFmt = data.tenant.dateFormat;
    row.getCell(colDuration).value = {
      formula: `${colDurAppliedMinutesLetter}${currentRow}/1440`,
    };
    row.getCell(colDuration).numFmt = durFormat;
    row.getCell(colDuration).font = {
      name: 'Arial',
      family: 2,
      bold: true,
    };
    row.getCell(colRate).value = {
      formula: `${colRateAppliedLetter}${currentRow}`,
    };
    row.getCell(colRate).numFmt = currencyFormat;
    row.getCell(colSum).value = {
      formula: `ROUND(${colDurAppliedMinutesLetter}${currentRow}*${colRateLetter}${currentRow}/60, 2)`,
    };
    row.getCell(colSum).numFmt = currencyFormat;
    row.getCell(colDurOrigDisplay).value = {
      formula: `${colDurOrigMinutesLetter}${currentRow}/1440`,
    };
    row.getCell(colDurOrigDisplay).numFmt = durFormat;
    row.getCell(colDurAppliedMinutes).value = {
      formula: `ROUND(${colDurOrigMinutesLetter}${currentRow}*${ws2Name}!E${ownerIndex + 2}*${colDurRowMultiLetter}${currentRow}+${colDurAddLetter}${currentRow}, 0)`,
    };
    row.getCell(colDurAppliedDisplay).value = {
      formula: `${colDurAppliedMinutesLetter}${currentRow}/1440`,
    };
    row.getCell(colDurAppliedDisplay).numFmt = durFormat;
    row.getCell(colDurChange).value = {
      formula: `${colDurAppliedMinutesLetter}${currentRow}-${colDurOrigMinutesLetter}${currentRow}`,
    };
    row.getCell(colDurUserMulti).value = {
      formula: `${ws2Name}!E${ownerIndex + 2}`,
    };
    row.getCell(colRateOrig).value = {
      formula: `${ws2Name}!C${ownerIndex + 2}`,
    };
    row.getCell(colRateOrig).numFmt = currencyFormat;
    row.getCell(colRateApplied).value = {
      formula: `${ws2Name}!C${ownerIndex + 2}*${ws2Name}!F${ownerIndex + 2}*${colRateRowMultiLetter}${currentRow}+${colRateAddLetter}${currentRow}`,
    };
    row.getCell(colRateApplied).numFmt = currencyFormat;
    row.getCell(colRateChange).value = {
      formula: `${colRateAppliedLetter}${currentRow}-${colRateOrigLetter}${currentRow}`,
    };
    row.getCell(colRateChange).numFmt = currencyFormat;
    row.getCell(colRateUserMulti).value = {
      formula: `${ws2Name}!F${ownerIndex + 2}`,
    };
    row.getCell(colSumOrig).value = {
      formula: `ROUND(${colDurOrigMinutesLetter}${currentRow}*${colSrcRateLetter}${currentRow}/60, 2)`,
    };
    row.getCell(colSumOrig).numFmt = currencyFormat;
    row.getCell(colSumApplied).value = {
      formula: `ROUND(${colDurAppliedMinutesLetter}${currentRow}*${colRateLetter}${currentRow}/60, 2)`,
    };
    row.getCell(colSumApplied).numFmt = currencyFormat;
    row.getCell(colSumChange).value = {
      formula: `${colSumAppliedLetter}${currentRow}-${colSumOrigLetter}${currentRow}`,
    };
    row.getCell(colSumChange).numFmt = currencyFormat;
    row.getCell(
      colSrcDate,
    ).numFmt = data.tenant.dateFormat;
    row.getCell(colSrcDur).value = {
      formula: `${colDurOrigMinutesLetter}${currentRow}/1440`,
    };
    row.getCell(colSrcDur).numFmt = durFormat;
    row.getCell(colSrcDur).font = {
      name: 'Arial',
      family: 2,
      bold: true,
    };
    row.getCell(colSrcRate).value = {
      formula: `${colRateOrigLetter}${currentRow}`,
    };
    row.getCell(colSrcRate).numFmt = currencyFormat;
    row.getCell(colSrcSum).value = {
      formula: `ROUND(${colDurOrigMinutesLetter}${currentRow}*${colSrcRateLetter}${currentRow}/60, 2)`,
    };
    row.getCell(colSrcSum).numFmt = currencyFormat;
    const tAlignment1 = { vertical: 'top', horizontal: 'left' };
    const tAlignment2 = { vertical: 'top', horizontal: 'right' };
    mainColsLeft.forEach(x => {
      row.getCell(x).alignment = tAlignment1;
      row.getCell(x).border = genBorder;
    });
    mainColsRight.forEach(x => {
      row.getCell(x).alignment = tAlignment2;
      row.getCell(x).border = genBorder;
    });
    durEditingCols.forEach(x => {
      row.getCell(x).alignment = tAlignment2;
      row.getCell(x).border = genBorder;
    });
    rateEditingCols.forEach(x => {
      row.getCell(x).alignment = tAlignment2;
      row.getCell(x).border = genBorder;
    });
    sumEditingCols.forEach(x => {
      row.getCell(x).alignment = tAlignment2;
      row.getCell(x).border = genBorder;
    });
    srcCols.forEach(x => {
      row.getCell(x).alignment = tAlignment1;
      row.getCell(x).border = genBorder;
    });
    srcColsLeft.forEach(x => {
      row.getCell(x).alignment = tAlignment1;
      row.getCell(x).border = genBorder;
    });
    srcColsRight.forEach(x => {
      row.getCell(x).alignment = tAlignment2;
      row.getCell(x).border = genBorder;
    });
    row.getCell(colDurRowMulti).fill = modFill;
    row.getCell(colDurAdd).fill = modFill;
    row.getCell(colRateRowMulti).fill = modFill;
    row.getCell(colRateAdd).fill = modFill;
    row.getCell(colDurRowMulti).border = modBorder;
    row.getCell(colDurAdd).border = modBorder;
    row.getCell(colRateRowMulti).border = modBorder;
    row.getCell(colRateAdd).border = modBorder;
    row.getCell(colDivider1).fill = divider1Fill;
    row.getCell(colDivider2).fill = divider2Fill;
    row.getCell(colDivider3).fill = divider3Fill;
    row.getCell(colDivider4).fill = divider4Fill;
    row.getCell(colDivider5).fill = divider5Fill;
  });
  // /times content

  // times summary
  const tsContent = [];
  tsContent[colSum] = '';
  tsContent[colSrcSum] = '';
  const tsRow = ws1.addRow(tsContent);
  currentRow = currentRow + 1;
  const tsRowNumber = currentRow;
  tsRow.height = 25;
  tsRow.getCell(colDuration).value = {
    formula: `${colDurAppliedDisplayLetter}${tsRowNumber}`,
  };
  tsRow.getCell(colDuration).numFmt = durFormat;
  tsRow.getCell(colSum).value = {
    formula: `${colSumAppliedLetter}${tsRowNumber}`,
  };
  tsRow.getCell(colSum).numFmt = currencyFormat;
  tsRow.getCell(colSrcDur).value = {
    formula: `SUM(${colDurOrigMinutesLetter}${topInfoRowCount + 3}:${colDurOrigMinutesLetter}${topInfoRowCount + times2.length + 2})/1440`,
  };
  tsRow.getCell(colSrcDur).numFmt = durFormat;
  tsRow.getCell(colDurOrigMinutes).value = {
    formula: `SUM(${colDurOrigMinutesLetter}${topInfoRowCount + 3}:${colDurOrigMinutesLetter}${topInfoRowCount + times2.length + 2})`,
  };
  tsRow.getCell(colDurOrigDisplay).value = {
    formula: `SUM(${colDurOrigMinutesLetter}${topInfoRowCount + 3}:${colDurOrigMinutesLetter}${topInfoRowCount + times2.length + 2})/1440`,
  };
  tsRow.getCell(colDurOrigDisplay).numFmt = durFormat;
  tsRow.getCell(colDurAppliedMinutes).value = {
    formula: `SUM(${colDurAppliedMinutesLetter}${topInfoRowCount + 3}:${colDurAppliedMinutesLetter}${topInfoRowCount + times2.length + 2})`,
  };
  tsRow.getCell(colDurAppliedDisplay).value = {
    formula: `SUM(${colDurAppliedMinutesLetter}${topInfoRowCount + 3}:${colDurAppliedMinutesLetter}${topInfoRowCount + times2.length + 2})/1440`,
  };
  tsRow.getCell(colDurAppliedDisplay).numFmt = durFormat;
  tsRow.getCell(colDurChange).value = {
    formula: `SUM(${colDurChangeLetter}${topInfoRowCount + 3}:${colDurChangeLetter}${topInfoRowCount + times2.length + 2})`,
  };
  tsRow.getCell(colSumOrig).value = {
    formula: `SUM(${colSumOrigLetter}${topInfoRowCount + 3}:${colSumOrigLetter}${topInfoRowCount + times2.length + 2})`,
  };
  tsRow.getCell(colSumOrig).numFmt = currencyFormat;
  tsRow.getCell(colSumApplied).value = {
    formula: `SUM(${colSumAppliedLetter}${topInfoRowCount + 3}:${colSumAppliedLetter}${topInfoRowCount + times2.length + 2})`,
  };
  tsRow.getCell(colSumApplied).numFmt = currencyFormat;
  tsRow.getCell(colSumChange).value = {
    formula: `SUM(${colSumChangeLetter}${topInfoRowCount + 3}:${colSumChangeLetter}${topInfoRowCount + times2.length + 2})`,
  };
  tsRow.getCell(colSumChange).numFmt = currencyFormat;
  tsRow.getCell(colSrcSum).value = {
    formula: `SUM(${colSrcSumLetter}${topInfoRowCount + 3}:${colSrcSumLetter}${topInfoRowCount + times2.length + 2})`,
  };
  tsRow.getCell(colSrcSum).numFmt = currencyFormat;
  const tsAlignment1 = { vertical: 'middle', horizontal: 'left' };
  const tsAlignment2 = { vertical: 'middle', horizontal: 'right' };
  mainColsLeft.forEach(x => {
    tsRow.getCell(x).font = h1Font;
    tsRow.getCell(x).fill = h1Fill;
    tsRow.getCell(x).alignment = tsAlignment1;
  });
  mainColsRight.forEach(x => {
    tsRow.getCell(x).font = h1Font;
    tsRow.getCell(x).fill = h1Fill;
    tsRow.getCell(x).alignment = tsAlignment2;
  });
  durEditingCols.forEach(x => {
    tsRow.getCell(x).font = h1Font;
    tsRow.getCell(x).fill = h1Fill;
    tsRow.getCell(x).alignment = tsAlignment2;
  });
  rateEditingCols.forEach(x => {
    tsRow.getCell(x).font = h1Font;
    tsRow.getCell(x).fill = h1Fill;
    tsRow.getCell(x).alignment = tsAlignment2;
  });
  sumEditingCols.forEach(x => {
    tsRow.getCell(x).font = h1Font;
    tsRow.getCell(x).fill = h1Fill;
    tsRow.getCell(x).alignment = tsAlignment2;
  });
  srcColsLeft.forEach(x => {
    tsRow.getCell(x).font = h1Font;
    tsRow.getCell(x).fill = h1Fill;
    tsRow.getCell(x).alignment = tsAlignment1;
  });
  srcColsRight.forEach(x => {
    tsRow.getCell(x).font = h1Font;
    tsRow.getCell(x).fill = h1Fill;
    tsRow.getCell(x).alignment = tsAlignment2;
  });
  const tsChangeFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFB4C7DC' },
  };
  tsRow.getCell(colDurChange).fill = tsChangeFill;
  tsRow.getCell(colSumChange).fill = tsChangeFill;
  tsRow.getCell(colDivider1).fill = divider1Fill;
  tsRow.getCell(colDivider2).fill = divider2Fill;
  tsRow.getCell(colDivider3).fill = divider3Fill;
  tsRow.getCell(colDivider4).fill = divider4Fill;
  tsRow.getCell(colDivider5).fill = divider5Fill;
  // /times summary

  // spacer row
  ws1.addRow([]);
  currentRow = currentRow + 1;
  // /spacer row

  const totalTitleAlignment = { vertical: 'bottom', horizontal: 'right' };
  const totalValueAlignment = { vertical: 'top', horizontal: 'right' };

  // total duration row title
  const totalDurRowTitle = ws1.addRow([trTotalDurTitle]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  totalDurRowTitle.getCell(1).alignment = totalTitleAlignment;
  // /total duration row title

  // total duration row
  const totalDurRowContent = [];
  totalDurRowContent[1] = '';
  const totalDurRow = ws1.addRow(totalDurRowContent);
  currentRow = currentRow + 1;
  totalDurRow.getCell(1).value = {
    formula: `${colDurAppliedDisplayLetter}${tsRowNumber}`,
  };
  totalDurRow.getCell(1).numFmt = durFormat;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  const totalDurRowFont = { name: 'Arial', family: 2, size: 20, bold: true };
  totalDurRow.getCell(1).font = totalDurRowFont;
  totalDurRow.getCell(1).alignment = totalValueAlignment;
  totalDurRow.height = 30;
  // /total duration row

  // spacer row
  ws1.addRow([]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  // /spacer row

  // total sum row title
  const totalSumRowTitle = ws1.addRow([trTotalSumTitle]);
  currentRow = currentRow + 1;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  totalSumRowTitle.getCell(1).alignment = totalTitleAlignment;
  // /total sum row title

  // total sum row
  const totalSumRowContent = [];
  totalSumRowContent[1] = '';
  const totalSumRow = ws1.addRow(totalSumRowContent);
  currentRow = currentRow + 1;
  totalSumRow.getCell(1).value = {
    formula: `${colSumAppliedLetter}${tsRowNumber}`,
  };
  totalSumRow.getCell(1).numFmt = currencyFormat;
  ws1.mergeCells(`A${currentRow}:I${currentRow}`);
  const totalSumRowFont = { name: 'Arial', family: 2, size: 20, bold: true };
  totalSumRow.getCell(1).font = totalSumRowFont;
  totalSumRow.getCell(1).alignment = totalValueAlignment;
  totalSumRow.height = 30;
  // /total sum row

  // ws2 columns
  const ws2ColNameOrig = 1;
  const ws2ColNameApplied = 2;
  const ws2ColRateOrig = 3;
  const ws2ColRateApplied = 4;
  const ws2ColDurMulti = 5;
  const ws2ColRateMulti = 6;
  const ws2ColRateOrigLetter = 'C';
  const ws2ColRateMultiLetter = 'F';
  // /ws2 columns

  // ws2 column widths
  ws2.getColumn(ws2ColNameOrig).width = 32;
  ws2.getColumn(ws2ColNameApplied).width = 32;
  ws2.getColumn(ws2ColRateOrig).width = 14;
  ws2.getColumn(ws2ColRateApplied).width = 14;
  ws2.getColumn(ws2ColDurMulti).width = 14;
  ws2.getColumn(ws2ColRateMulti).width = 14;
  // /ws2 column widths

  // ws2 heading
  const phContent = [];
  phContent[ws2ColNameOrig] = trPeopleOrigName;
  phContent[ws2ColNameApplied] = trPeopleAppliedName;
  phContent[ws2ColRateOrig] = trPeopleOrigRate;
  phContent[ws2ColRateApplied] = trPeopleAppliedRate;
  phContent[ws2ColDurMulti] = trPeopleDurMulti;
  phContent[ws2ColRateMulti] = trPeopleRateMulti;
  const ws2MainColsLeft = [1, 2];
  const ws2MainColsRight = [3, 4, 5, 6];
  const phRow = ws2.addRow(phContent);
  const phFont = { name: 'Arial', family: 2, bold: true };
  const phFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFC9C9C9' },
  };
  const pAlignment1 = { vertical: 'top', horizontal: 'left' };
  const pAlignment2 = { vertical: 'top', horizontal: 'right' };
  ws2MainColsLeft.forEach(x => {
    phRow.getCell(x).font = phFont;
    phRow.getCell(x).fill = phFill;
    phRow.getCell(x).alignment = pAlignment1;
  });
  ws2MainColsRight.forEach(x => {
    phRow.getCell(x).font = phFont;
    phRow.getCell(x).fill = phFill;
    phRow.getCell(x).alignment = pAlignment2;
  });
  // /ws2 heading

  // ws2 content
  uniqueOwners.forEach((owner, index) => {
    const content = [];
    content[ws2ColNameOrig] = owner.name;
    content[ws2ColNameApplied] = owner.name;
    content[ws2ColRateOrig] = 100;
    content[ws2ColRateApplied] = '';
    content[ws2ColDurMulti] = 1;
    content[ws2ColRateMulti] = 1;
    const row = ws2.addRow(content);
    row.getCell(ws2ColRateApplied).value = {
      formula: `${ws2ColRateOrigLetter}${index +
        2}*${ws2ColRateMultiLetter}${index + 2}`,
    };
    row.getCell(ws2ColNameOrig).alignment = { horizontal: 'left' };
    row.getCell(ws2ColNameApplied).alignment = { horizontal: 'left' };
    row.getCell(ws2ColNameApplied).fill = modFill;
    row.getCell(ws2ColNameApplied).border = modBorder;
    row.getCell(ws2ColRateOrig).fill = modRateOrigFill;
    row.getCell(ws2ColRateOrig).border = modBorder;
    row.getCell(ws2ColDurMulti).fill = modFill;
    row.getCell(ws2ColDurMulti).border = modBorder;
    row.getCell(ws2ColRateMulti).fill = modFill;
    row.getCell(ws2ColRateMulti).border = modBorder;
  });
  // /ws2 content

  // conditional formatting: time fields
  const genChangeFill = {
    type: 'pattern',
    pattern: 'solid',
    bgColor: { argb: 'FFB4ED93' },
  };
  const genChangeIncFill = {
    type: 'pattern',
    pattern: 'solid',
    bgColor: { argb: 'FFB4ED93' },
  };
  const genChangeDecFill = {
    type: 'pattern',
    pattern: 'solid',
    bgColor: { argb: 'FFFFD7D7' },
  };
  ws1.addConditionalFormatting({
    ref: `${colDateLetter}${topInfoRowCount + 3}:${colDateLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colDateLetter}${topInfoRowCount + 3}<>${colSrcDateLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeFill },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colClientLetter}${topInfoRowCount + 3}:${colClientLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colClientLetter}${topInfoRowCount + 3}<>${colSrcClientLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeFill },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colProjectLetter}${topInfoRowCount + 3}:${colProjectLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colProjectLetter}${topInfoRowCount + 3}<>${colSrcProjectLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeFill },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colUserLetter}${topInfoRowCount + 3}:${colUserLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colUserLetter}${topInfoRowCount + 3}<>${colSrcUserLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeFill },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colTaskLetter}${topInfoRowCount + 3}:${colTaskLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colTaskLetter}${topInfoRowCount + 3}<>${colSrcTaskLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeFill },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colDurLetter}${topInfoRowCount + 3}:${colDurLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colDurAppliedMinutesLetter}${topInfoRowCount + 3}>${colDurOrigMinutesLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeIncFill },
      },
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colDurAppliedMinutesLetter}${topInfoRowCount + 3}<${colDurOrigMinutesLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeDecFill },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colRateLetter}${topInfoRowCount + 3}:${colRateLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colRateAppliedLetter}${topInfoRowCount + 3}>${colRateOrigLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeIncFill },
      },
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colRateAppliedLetter}${topInfoRowCount + 3}<${colRateOrigLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeDecFill },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colSumLetter}${topInfoRowCount + 3}:${colSumLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colSumLetter}${topInfoRowCount + 3}>${colSrcSumLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeIncFill },
      },
      {
        type: 'expression',
        formulae: [`AND(${diffToggle}="YES", ${colSumLetter}${topInfoRowCount + 3}<${colSrcSumLetter}${topInfoRowCount + 3})`],
        style: { fill: genChangeDecFill },
      },
    ],
  });
  // /conditional formatting: time fields

  // conditional formatting: modifier fields
  const durRateChangeFill1 = {
    type: 'pattern',
    pattern: 'solid',
    bgColor: { argb: 'FFB4ED93' },
  };
  const durRateChangeFill2 = {
    type: 'pattern',
    pattern: 'solid',
    bgColor: { argb: 'FFFFD7D7' },
  };
  ws1.addConditionalFormatting({
    ref: `${colDurChangeLetter}${topInfoRowCount + 3}:${colDurChangeLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['0'],
        style: { fill: durRateChangeFill1 },
      },
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: { fill: durRateChangeFill2 },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colRateChangeLetter}${topInfoRowCount + 3}:${colRateChangeLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['0'],
        style: { fill: durRateChangeFill1 },
      },
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: { fill: durRateChangeFill2 },
      },
    ],
  });
  ws1.addConditionalFormatting({
    ref: `${colSumChangeLetter}${topInfoRowCount + 3}:${colSumChangeLetter}${topInfoRowCount + times2.length + 2}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['0'],
        style: { fill: durRateChangeFill1 },
      },
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: { fill: durRateChangeFill2 },
      },
    ],
  });
  // /conditional formatting: modifier fields

  // header and footer
  ws1.headerFooter.oddHeader = '&Rtuumik.com';
  ws1.headerFooter.evenHeader = '&Rtuumik.com';
  ws1.headerFooter.oddFooter = '&C&P';
  ws1.headerFooter.evenFooter = '&C&P';
  // /header and footer

  // page setup
  ws1.pageSetup.paperSize = 9;
  ws1.pageSetup.orientation = 'landscape';
  ws1.pageSetup.fitToPage = true;
  ws1.pageSetup.fitToWidth = 1;
  ws1.pageSetup.fitToHeight = 0;
  ws1.pageSetup.printTitlesRow = `${th1RowNumber}:${th2RowNumber}`;
  // /page setup

  const buf = await wb.xlsx.writeBuffer();

  // base64
  const base64String = buf.toString('base64');
  const fileBase64 = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64String}`;
  // /base64

  const fileName = 'report.xlsx';
  const responseBody = { exportFiles: [{ fileData: fileBase64, fileName }] };
  return res.status(200).json(responseBody);
}
