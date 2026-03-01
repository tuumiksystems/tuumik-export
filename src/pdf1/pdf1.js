/* Copyright (C) 2017-2025 Tuumik Systems OÜ */

import PdfPrinter from 'pdfmake';
import fs from 'fs';
import { Base64Encode } from 'base64-stream';
import dayjs from 'dayjs';
import utc from 'dayjs/plugin/utc.js';

dayjs.extend(utc);

export async function pdf1(req, res) {
  // API key
  const apiKeyFromHeader = req.headers['x-api-key'];
  const apiKey = process.env.API_KEY;
  if (apiKeyFromHeader !== apiKey) return res.status(403).json({ message: 'Forbidden: Invalid API Key' });
  // /API key

  const data = req.body;

  const docDefinition = { content: [] };

  // images
  const logoBase64Src = fs.readFileSync('src/assets/top.png', 'base64');
  const logoBase64Full = `data:image/png;base64,${logoBase64Src}`;
  docDefinition.images = {
    topLogo: logoBase64Full,
  };
  // /images

  const conf = {
    defaultStyle: {
      fontSize: 8,
      color: '#000000',
    },
    clientProject: {
      fontSize: 12,
      alignment: 'left',
      color: '#000000',
      bold: true,
      italics: false,
      marginTop: 0,
      marginBottom: 0,
    },
    period: {
      fontSize: 8,
      alignment: 'left',
      color: '#000000',
      bold: false,
      italics: false,
      marginTop: 0,
      marginBottom: 10,
    },
    times: {
      borderWidth: 0.6,
      borderColor: '#bfbfbf',
      paddingVert: 2,
      paddingHori: 3,
      paddingVertSummary: 4,
      noEdgePadding: false,
      useVertSummary: true,
      heading: {
        border: 'full',
        color: '#000000',
        fillColor: '#ececec',
        unit: {
          fontSize: 7,
          separator: '\n',
          color: '#7d7d7d',
        },
      },
      content: {
        border: 'full',
        color: '#000000',
        fillColor: '#ffffff',
      },
      summary: {
        border: 'full',
        color: '#000000',
        fillColor: '#ececec',
      },
    },
  };

  const texts = {
    header: 'Time',
    date: 'Date',
    duration: 'Duration',
    owner: 'Person',
    task: 'Task',
    clientProjectAny: 'Any project',
  };

  const sf = data.exportOptions.showFields;
  const times = data.times;

  // top logo
  const topLogo = { image: 'topLogo', height: 20, width: 20, margin: [0, 0, 0, 10] };
  docDefinition.content.push(topLogo);
  // /top logo

  // top info 1: client and project list
  if (data.meta.clients) {
    const cpList = [];
    data.meta.clients.forEach(client => {
      client.projects.forEach(project => {
        cpList.push({ client: client.name, project: project.name });
      });
    });
    cpList.forEach((cliPro, index) => {
      const cpRow = {
        text: `${cliPro.client} - ${cliPro.project}`,
        fontSize: conf.clientProject.fontSize,
        color: conf.clientProject.color,
        bold: conf.clientProject.bold,
        italics: conf.clientProject.italics,
        margin: [0, conf.clientProject.marginTop, 0, conf.clientProject.marginBottom],
      };
      docDefinition.content.push(cpRow);
    });
  } else {
    const cpRow = {
      text: texts.clientProjectAny,
      fontSize: conf.clientProject.fontSize,
      color: conf.clientProject.color,
      bold: conf.clientProject.bold,
      italics: conf.clientProject.italics,
      margin: [0, conf.clientProject.marginTop, 0, conf.clientProject.marginBottom],
    };
    docDefinition.content.push(cpRow);
  }
  // /top info 1: client and project list

  // top info 2: period string
  const startStr = data.meta.period.start ? dayjs.utc(data.meta.period.start).format(data.tenant.dateFormat) : '*';
  const endStr = data.meta.period.end ? dayjs.utc(data.meta.period.end).format(data.tenant.dateFormat) : '*';
  const periodStr = `${startStr} - ${endStr}`;
  const periodRow = {
    text: periodStr,
    fontSize: conf.period.fontSize,
    color: conf.period.color,
    bold: conf.period.bold,
    italics: conf.period.italics,
    margin: [0, conf.period.marginTop, 0, conf.period.marginBottom],
  };
  docDefinition.content.push(periodRow);
  // /top info 2: period string

  // time table
  const timeTableWidths = [];
  if (sf.date) timeTableWidths.push('12%');
  if (sf.user) timeTableWidths.push('20%');
  if (sf.task) timeTableWidths.push('*');
  if (sf.duration) timeTableWidths.push('12%');

  const timeTable = {
    table: {
      headerRows: 1,
      dontBreakRows: true,
      widths: timeTableWidths,
      body: [],
    },
    layout: 'times',
    margin: [0, 4, 0, 0],
  };

  const tHeadingBorder = [false, true, false, true];
  const tContentBorder = [false, true, false, true];
  const tSummaryBorder = [false, true, false, true];

  const tHeaderRow = [];
  if (sf.date) {
    tHeaderRow.push({
      text: texts.date,
      bold: true,
      border: tHeadingBorder,
      color: conf.times.heading.color,
      fillColor: conf.times.heading.fillColor,
    });
  }
  if (sf.user) {
    tHeaderRow.push({
      text: texts.owner,
      bold: true,
      border: tHeadingBorder,
      color: conf.times.heading.color,
      fillColor: conf.times.heading.fillColor,
    });
  }
  if (sf.task) {
    tHeaderRow.push({
      text: texts.task,
      bold: true,
      border: tHeadingBorder,
      color: conf.times.heading.color,
      fillColor: conf.times.heading.fillColor,
    });
  }
  if (sf.duration) {
    tHeaderRow.push({
      text: texts.duration,
      alignment: 'right',
      bold: true,
      border: tHeadingBorder,
      color: conf.times.heading.color,
      fillColor: conf.times.heading.fillColor,
    });
  }
  timeTable.table.body.push(tHeaderRow);

  times.forEach(time => {
    const tRow = [];
    if (sf.date) {
      tRow.push({
        text: displayDate(time),
        border: tContentBorder,
        color: conf.times.content.color,
        fillColor: conf.times.content.fillColor,
      });
    }
    if (sf.user) {
      tRow.push({
        text: time.ownerName,
        border: tContentBorder,
        color: conf.times.content.color,
        fillColor: conf.times.content.fillColor,
      });
    }
    if (sf.task) {
      tRow.push({
        text: time.taskDesc,
        border: tContentBorder,
        color: conf.times.content.color,
        fillColor: conf.times.content.fillColor,
      });
    }
    if (sf.duration) {
      tRow.push({
        text: displayDuration(time),
        alignment: 'right',
        border: tContentBorder,
        color: conf.times.content.color,
        fillColor: conf.times.content.fillColor,
      });
    }
    timeTable.table.body.push(tRow);
  });

  // summary row
  const tsRow = [];
  if (sf.date) {
    tsRow.push({
      text: '',
      border: tSummaryBorder,
      fillColor: conf.times.summary.fillColor,
    });
  }
  if (sf.user) {
    tsRow.push({
      text: '',
      border: tSummaryBorder,
      fillColor: conf.times.summary.fillColor,
    });
  }
  if (sf.task) {
    tsRow.push({
      text: '',
      border: tSummaryBorder,
      fillColor: conf.times.summary.fillColor,
    });
  }
  if (sf.duration) {
    tsRow.push({
      text: durationTotal(),
      alignment: 'right',
      border: tSummaryBorder,
      color: conf.times.summary.color,
      fillColor: conf.times.summary.fillColor,
      bold: true,
    });
  }
  timeTable.table.body.push(tsRow);
  // /summary row

  if (times.length) {
    docDefinition.content.push(timeTable);
  } else {
    docDefinition.content.push({
      text: 'No time entries.',
      margin: [0, 10, 0, 10],
    });
  }
  // /time table

  // helper functions
  function displayDate(time) {
    return dayjs.utc(time.date).format(data.tenant.dateFormat);
  }

  function displayDuration(time) {
    const minutes = time.endMinute - time.startMinute;
    return minutesToDuration(minutes, false);
  }

  function durationTotal() {
    let totalMinutes = 0;
    for (const time of times) {
      totalMinutes += time.endMinute - time.startMinute;
    }
    return minutesToDuration(totalMinutes, false);
  }

  function minutesToDuration(minutesIn, showZeroHours) {
    const minutesInAbs = Math.abs(minutesIn);
    const hoursOut = Math.floor(minutesInAbs / 60);
    const minutesOut = minutesInAbs % 60;
    const sign = minutesIn < 0 ? '-' : '';
    if (!showZeroHours && hoursOut < 1) return `${sign}${minutesOut}m`;
    return `${sign}${hoursOut}h ${minutesOut}m`;
  };
  // /helper functions

  // table layouts
  docDefinition.tableLayouts = {
    times: {
      hLineWidth() {
        return conf.times.borderWidth;
      },
      vLineWidth() {
        return conf.times.borderWidth;
      },
      hLineColor() {
        return conf.times.borderColor;
      },
      vLineColor() {
        return conf.times.borderColor;
      },
      paddingLeft(i) {
        if (i === 0 && conf.times.noEdgePadding) return 0;
        return conf.times.paddingHori;
      },
      paddingRight(i, node) {
        if (i === node.table.widths.length - 1 && conf.times.noEdgePadding) return 0;
        return conf.times.paddingHori;
      },
      paddingTop(i, node) {
        if (i >= node.table.body.length - 1 && conf.times.useVertSummary) return conf.times.paddingVertSummary;
        return conf.times.paddingVert;
      },
      paddingBottom(i, node) {
        if (i >= node.table.body.length - 1 && conf.times.useVertSummary) return conf.times.paddingVertSummary;
        return conf.times.paddingVert;
      },
    },
  };
  // /table layouts

  // default style
  docDefinition.defaultStyle = {
    fontSize: conf.defaultStyle.fontSize,
    color: conf.defaultStyle.color,
  };
  // /default style

  // metadata
  docDefinition.info = {
    title: '',
    author: '',
    creator: 'Tuumik',
    producer: 'Tuumik',
  };
  // /metadata

  // output
  const fonts = {
    Roboto: {
      normal: 'src/assets/fonts/Roboto-Regular.ttf',
      bold: 'src/assets/fonts/Roboto-Medium.ttf',
      italics: 'src/assets/fonts/Roboto-Italic.ttf',
      bolditalics: 'src/assets/fonts/Roboto-MediumItalic.ttf'
    }
  };
  const printer = new PdfPrinter(fonts);
  const pdfDocGenerator = printer.createPdfKitDocument(docDefinition);
  const pdfDoc = pdfDocGenerator;
  // /output

  // base64
  async function getPDFBase64Text(pdfDoc) {
    return new Promise(resolve => {
      const b64 = new Base64Encode();
      let base64String = '';
      const stream = pdfDoc.pipe(b64);
      pdfDoc.end();
      stream.on('data', chunk => { base64String += chunk });
      stream.on('end', () => resolve(base64String));
    })
  }
  
  const base64Text = await getPDFBase64Text(pdfDoc);
  const fileBase64 = `data:application/pdf;base64,${base64Text}`;
  // /base64

  const fileName = 'report.pdf';
  const responseBody = { exportFiles: [{ fileData: fileBase64, fileName }] };
  return res.status(200).json(responseBody);
}