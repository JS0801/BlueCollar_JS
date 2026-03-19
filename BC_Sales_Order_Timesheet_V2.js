/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define([
  'N/search',
  'N/log',
  'N/file',
  'N/record',
  './js_lib'
], function (search, log, file, record, XLSX) {

  function onRequest(context) {
    try {
      if (context.request.method !== 'GET') return;

      var params = context.request.parameters || {};
      var tranid = params.tranid;
      var exportType = params.export || '';

      var tranIds = normalizeTranIds(tranid);
      if (!tranIds.length) {
        context.response.write('No tranid provided.');
        return;
      }

      var empNameArr = getEmployeeList();
      var invoiceData = buildConsolidatedInvoiceData(tranIds);

      var timesheetDataByTran = [];
      for (var i = 0; i < tranIds.length; i++) {
        timesheetDataByTran.push(buildTimesheetData(tranIds[i], empNameArr));
      }

      if (exportType === 'excel' || exportType === 'xlsx') {
        var wb = buildWorkbook(invoiceData, timesheetDataByTran);
        var base64 = XLSX.write(wb, {
          type: 'base64',
          bookType: 'xlsx'
        });

        var outFile = file.create({
          name: 'Weekly_Timesheet_' + new Date().toISOString().slice(0, 10) + '.xlsx',
          fileType: file.Type.EXCEL,
          contents: base64,
          encoding: file.Encoding.BASE_64
        });

        context.response.writeFile(outFile, false);
        return;
      }

      var responseObj = {
        invoice: invoiceData,
        timesheets: timesheetDataByTran
      };

      context.response.writeLine("<#assign ObjDetail=" + JSON.stringify(responseObj) + " />");

    } catch (e) {
      log.error('Suitelet Error', e);
      context.response.write('Error: ' + e.message);
    }
  }

  // --------------------------------------------------------------------------
  // WORKBOOK BUILD
  // --------------------------------------------------------------------------
  function buildWorkbook(invoiceData, timesheetDataByTran) {
    var wb = XLSX.utils.book_new();
    var sheet = createSheetBuilder('Weekly Timesheet');

    addInvoiceSection(sheet, invoiceData);
    sheet.blankRow(2);

    for (var i = 0; i < timesheetDataByTran.length; i++) {
      addTimesheetSection(sheet, timesheetDataByTran[i]);

      if (i !== timesheetDataByTran.length - 1) {
        sheet.blankRow(2);
        sheet.mergeRowValue(sheet.row, 1, 17, '', styles.blackDivider);
        sheet.nextRow();
        sheet.blankRow(2);
      }
    }

    var ws = sheet.toWorksheet();

    ws['!cols'] = [
      { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 18 }, { wch: 12 },
      { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
      { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 16 },
      { wch: 18 }, { wch: 18 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Weekly Timesheet');
    return wb;
  }

  function addInvoiceSection(sheet, invoiceData) {
    var cats = Object.keys(invoiceData.catMap || {}).sort();

    sheet.mergeRowValue(sheet.row, 1, 11, 'DRAFT INVOICE', styles.invoiceTitle);
    sheet.mergeRowValue(sheet.row, 12, 16, '', styles.logoBox);
    sheet.nextRow();

    var metaStart = sheet.row;

    sheet.mergeRowValue(sheet.row, 1, 5, 'ATTN:\n' + stripHtmlBreaks(invoiceData.billAddrHtml), styles.wrapBox);
    sheet.mergeRowValue(sheet.row, 6, 11, 'Invoice Date:\n' + invoiceData.soDate, styles.wrapBox);
    sheet.mergeRowValue(sheet.row, 12, 16, stripHtmlBreaks(invoiceData.subAddrHtml) + '\nABN: ' + invoiceData.subABN, styles.wrapBox);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 6, 11, 'Invoice Number:\nDRAFT', styles.wrapBox);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 6, 11, 'PO Number:\n' + invoiceData.poNumbers.join(', '), styles.wrapBox);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 6, 11, 'Customer Reference:\n' + invoiceData.projectRefs.join(', '), styles.wrapBox);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 1, 16, '', styles.noBorder);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 1, 16, 'Memo:\n' + invoiceData.memos.join(' | '), styles.wrapBox);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 1, 16, '', styles.noBorder);
    sheet.nextRow();

    var hdrRow = sheet.row;
    sheet.mergeRowValue(hdrRow, 1, 8, 'Description', styles.tableHeaderLeft);
    sheet.mergeRowValue(hdrRow, 9, 10, 'Price', styles.tableHeader);
    sheet.mergeRowValue(hdrRow, 11, 12, invoiceData.isAmericas ? 'TAX' : 'GST', styles.tableHeader);
    sheet.mergeRowValue(hdrRow, 13, 14, invoiceData.TAX_LABEL_AMT, styles.tableHeader);
    sheet.mergeRowValue(hdrRow, 15, 16, 'Amount ' + invoiceData.currencyText, styles.tableHeader);
    sheet.nextRow();

    var detailStart = sheet.row;

    for (var i = 0; i < cats.length; i++) {
      var cat = cats[i];
      var rowNum = sheet.row;
      var amt = parseFloat(invoiceData.catMap[cat].amountSum || 0) || 0;
      var txa = Math.abs(parseFloat(invoiceData.catMap[cat].taxAmtSum || 0) || 0);
      var txr = parseFloat(invoiceData.catMap[cat].taxRateMax || 0) || 0;

      sheet.mergeRowValue(rowNum, 1, 8, cat, styles.cellLeft);
      sheet.mergeRowValue(rowNum, 9, 10, amt, styles.currency);
      sheet.mergeRowValue(rowNum, 11, 12, txr / 100, styles.percent);
      sheet.mergeRowValue(rowNum, 13, 14, txa, styles.currency);
      sheet.mergeRowFormula(rowNum, 15, 16, '=I' + rowNum + '+M' + rowNum, styles.currency);
      sheet.nextRow();
    }

    var detailEnd = sheet.row - 1;

    sheet.mergeRowValue(sheet.row, 1, 16, '', styles.noBorder);
    sheet.nextRow();

    var subtotalRow = sheet.row;
    sheet.mergeRowValue(subtotalRow, 1, 12, '', styles.noBorder);
    sheet.mergeRowValue(subtotalRow, 13, 14, 'Subtotal', styles.labelRight);
    sheet.mergeRowFormula(subtotalRow, 15, 16, '=SUM(I' + detailStart + ':I' + detailEnd + ')', styles.currency);
    sheet.nextRow();

    var gstRow = sheet.row;
    sheet.mergeRowValue(gstRow, 1, 12, '', styles.noBorder);
    sheet.mergeRowValue(gstRow, 13, 14, invoiceData.TAX_LABEL_TOTAL, styles.labelRight);
    sheet.mergeRowFormula(gstRow, 15, 16, '=SUM(M' + detailStart + ':M' + detailEnd + ')', styles.currency);
    sheet.nextRow();

    var totalRow = sheet.row;
    sheet.mergeRowValue(totalRow, 1, 12, '', styles.noBorder);
    sheet.mergeRowValue(totalRow, 13, 14, 'TOTAL ' + invoiceData.currencyText, styles.totalLabel);
    sheet.mergeRowFormula(totalRow, 15, 16, '=O' + subtotalRow + '+O' + gstRow, styles.totalCurrency);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 1, 16, '', styles.noBorder);
    sheet.nextRow();

    var bankInfo =
      'Sales Orders: ' + invoiceData.soNumbers.join(', ') + '\n\n' +
      'Customer: ' + invoiceData.customerNames.join(', ') + '\n\n' +
      'Due Date: ' + stripHtmlBreaks(invoiceData.dueDate) + '\n\n' +
      'Payment Terms: ' + stripHtmlBreaks(invoiceData.terms) + '\n\n' +
      'Please email remittance advice to ' + invoiceData.remitEmail + '\n\n' +
      'BANK ACCOUNT DETAILS\n' +
      'Account Name: ' + invoiceData.acctName + '\n' +
      'Bank: ' + invoiceData.bankName + '\n' +
      'BSB: ' + invoiceData.bsb + '\n' +
      'Account: ' + invoiceData.acctNum;

    sheet.mergeRowValue(sheet.row, 1, 16, bankInfo, styles.wrapBox);
    sheet.nextRow();
  }

  function addTimesheetSection(sheet, ts) {
    var x = ts.groupedData || {};
    var h = ts.headerInfo || {};
    var laborKey = x.Labour ? 'Labour' : 'Labor';
    var labelLabor = ts.replaceLabor ? 'Labour' : 'Labor';

    sheet.mergeRowValue(sheet.row, 1, 4, '', styles.logoBox);
    sheet.mergeRowValue(sheet.row, 5, 17, 'Weekly Timesheet - ' + h.docNumber, styles.sectionTitle);
    sheet.nextRow();
    sheet.blankRow(1);

    var leftInfo = [
      ['Client:', h.client],
      ['Customer Ref #:', h.customerRef],
      ['Week-Ending:', formatDateMMDDYYYY(h.weekEnding)],
      ['C2O Project Manager:', h.projectManager],
      ['Description of Work:', h.description],
      ['Document #:', h.docNumber]
    ];

    var rightInfo = [
      ['Project:', h.reportingProject],
      ['C2O Job:', h.projectName],
      ['Supervisor:', h.supervisor],
      ['Start Time: Monday – Friday', h.startTime, 'Finish Time:', h.endTime],
      ['Start Time: Weekend / Holiday', h.startTime, 'Finish Time:', h.endTime]
    ];

    var maxInfoRows = Math.max(leftInfo.length, rightInfo.length);
    for (var i = 0; i < maxInfoRows; i++) {
      var rowNum = sheet.row;

      if (leftInfo[i]) {
        sheet.mergeRowValue(rowNum, 1, 2, leftInfo[i][0], styles.infoHeader);
        sheet.mergeRowValue(rowNum, 3, 6, leftInfo[i][1], styles.infoValue);
      }

      sheet.mergeRowValue(rowNum, 7, 8, '', styles.noBorder);

      if (rightInfo[i]) {
        if (rightInfo[i].length === 2) {
          sheet.mergeRowValue(rowNum, 9, 10, rightInfo[i][0], styles.infoHeader);
          sheet.mergeRowValue(rowNum, 11, 16, rightInfo[i][1], styles.infoValue);
        } else {
          sheet.mergeRowValue(rowNum, 9, 10, rightInfo[i][0], styles.infoHeader);
          sheet.setCell(rowNum, 11, rightInfo[i][1], styles.infoValue);
          sheet.mergeRowValue(rowNum, 12, 13, rightInfo[i][2], styles.infoHeader);
          sheet.setCell(rowNum, 14, rightInfo[i][3], styles.infoValue);
        }
      }

      sheet.nextRow();
    }

    sheet.blankRow(2);

    if (x[laborKey]) {
      addLaborSection(sheet, x[laborKey], labelLabor);
      sheet.blankRow(2);
    }

    if (x['Equipment / Vehicle Rental']) {
      addEquipmentSection(sheet, x['Equipment / Vehicle Rental']);
      sheet.blankRow(2);
    }

    if (x.Materials) {
      addMaterialsSection(sheet, x.Materials);
      sheet.blankRow(2);
    }

    if (x.Expenses) {
      addExpensesSection(sheet, x.Expenses);
      sheet.blankRow(2);
    }

    if (ts.legendArray && ts.legendArray.length) {
      var legend = 'Time Type Legend: ';
      for (var l = 0; l < ts.legendArray.length; l++) {
        legend += ts.legendArray[l].abbr + ' - ' + ts.legendArray[l].label;
        if (l !== ts.legendArray.length - 1) legend += ' | ';
      }
      sheet.mergeRowValue(sheet.row, 1, 17, legend, styles.legend);
      sheet.nextRow();
    }
  }

  function addLaborSection(sheet, labor, labelLabor) {
    var dayCount = labor[0].days.length;
    var startColDays = 7;
    var endColDays = startColDays + dayCount - 1;
    var totalWeekCol = endColDays + 1;
    var rateCol = totalWeekCol + 1;
    var claimCol = rateCol + 1;

    sheet.mergeRowValue(sheet.row, 1, 6, labelLabor, styles.tableHeader);
    sheet.mergeRowValue(sheet.row, 7, 17, 'ALL HOURS SHOWN ARE HOURS WORKED', styles.centerBox);
    sheet.nextRow();

    var row1 = sheet.row;
    sheet.mergeRowValue(row1, 1, 2, 'Name', styles.tableHeader);
    sheet.mergeRowValue(row1, 3, 4, 'Role', styles.tableHeader);
    sheet.setCell(row1, 5, 'Time Type', styles.tableHeader);
    sheet.setCell(row1, 6, 'Shift Type', styles.tableHeader);

    var d;
    for (d = 0; d < dayCount; d++) {
      sheet.setCell(row1, startColDays + d, getDayName(labor[0].days[d].date), styles.tableHeader);
    }
    sheet.setCell(row1, totalWeekCol, 'Total Week', styles.tableHeader);
    sheet.setCell(row1, rateCol, 'Rate', styles.tableHeader);
    sheet.setCell(row1, claimCol, 'Claim Amount', styles.tableHeader);
    sheet.nextRow();

    var row2 = sheet.row;
    for (d = 0; d < dayCount; d++) {
      sheet.setCell(row2, startColDays + d, formatDateMMDDYYYY(labor[0].days[d].date), styles.tableHeaderDate);
    }
    sheet.nextRow();

    var detailStart = sheet.row;

    var i, j;
    for (i = 1; i < labor.length - 1; i++) {
      var r = sheet.row;

      sheet.mergeRowValue(r, 1, 2, labor[i].employee, styles.cell);
      sheet.mergeRowValue(r, 3, 4, labor[i].role, styles.cell);
      sheet.setCell(r, 5, labor[i].shiftType, styles.cell);
      sheet.setCell(r, 6, labor[i].shift, styles.cell);

      for (j = 0; j < dayCount; j++) {
        sheet.setCell(r, startColDays + j, parseFloat(labor[i].days[j].hours || 0), styles.decimal);
      }

      sheet.setFormula(r, totalWeekCol, '=SUM(' + colLetter(startColDays) + r + ':' + colLetter(endColDays) + r + ')', styles.decimal);
      sheet.setCell(r, rateCol, parseFloat(labor[i].rate || 0), styles.currency);
      sheet.setFormula(r, claimCol, '=ROUND(' + colLetter(totalWeekCol) + r + '*' + colLetter(rateCol) + r + ',2)', styles.currency);
      sheet.mergeRowValue(r, claimCol + 1, 17, labor[i].notes || '', styles.cell);
      sheet.nextRow();
    }

    var detailEnd = sheet.row - 1;
    var totalRow = sheet.row;

    sheet.mergeRowValue(totalRow, 1, 5, '', styles.noBorder);
    sheet.setCell(totalRow, 6, 'TOTAL', styles.totalLabel);

    for (j = 0; j < dayCount; j++) {
      var c = startColDays + j;
      sheet.setFormula(totalRow, c, '=SUM(' + colLetter(c) + detailStart + ':' + colLetter(c) + detailEnd + ')', styles.totalDecimal);
    }

    sheet.setFormula(totalRow, totalWeekCol, '=SUM(' + colLetter(totalWeekCol) + detailStart + ':' + colLetter(totalWeekCol) + detailEnd + ')', styles.totalDecimal);
    sheet.setCell(totalRow, rateCol, '', styles.totalLabel);
    sheet.setFormula(totalRow, claimCol, '=SUM(' + colLetter(claimCol) + detailStart + ':' + colLetter(claimCol) + detailEnd + ')', styles.totalCurrency);
    sheet.nextRow();
  }

  function addEquipmentSection(sheet, rows) {
    var dayCount = rows[0].days.length;
    var dayStart = 5;
    var dayEnd = dayStart + dayCount - 1;
    var totalWeekCol = dayEnd + 1;

    sheet.mergeRowValue(sheet.row, 1, 17, 'Equipment / Vehicle Rental', styles.tableHeader);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 1, 4, 'Role', styles.tableHeader);
    for (var d = 0; d < dayCount; d++) {
      sheet.setCell(sheet.row, dayStart + d, getDayName(rows[0].days[d].date), styles.tableHeader);
    }
    sheet.setCell(sheet.row, totalWeekCol, 'Total Week', styles.tableHeader);
    sheet.mergeRowValue(sheet.row, totalWeekCol + 1, 17, 'Notes', styles.tableHeader);
    sheet.nextRow();

    for (var d2 = 0; d2 < dayCount; d2++) {
      sheet.setCell(sheet.row, dayStart + d2, formatDateMMDDYYYY(rows[0].days[d2].date), styles.tableHeaderDate);
    }
    sheet.nextRow();

    for (var i = 1; i < rows.length; i++) {
      var r = sheet.row;
      var isTotal = (i === rows.length - 1 && String(rows[i].employee || '').toUpperCase() === 'TOTAL');

      sheet.mergeRowValue(r, 1, 4, rows[i].role, isTotal ? styles.totalLabel : styles.cell);
      for (var j = 0; j < dayCount; j++) {
        sheet.setCell(r, dayStart + j, parseFloat(rows[i].days[j].hours || 0), isTotal ? styles.totalDecimal : styles.decimal);
      }

      if (isTotal) {
        sheet.setCell(r, totalWeekCol, parseFloat(rows[i].totalWeek || 0), styles.totalDecimal);
      } else {
        sheet.setFormula(r, totalWeekCol, '=SUM(' + colLetter(dayStart) + r + ':' + colLetter(dayEnd) + r + ')', styles.decimal);
      }

      sheet.mergeRowValue(r, totalWeekCol + 1, 17, '', isTotal ? styles.totalLabel : styles.cell);
      sheet.nextRow();
    }
  }

  function addMaterialsSection(sheet, rows) {
    sheet.mergeRowValue(sheet.row, 1, 17, 'Materials', styles.tableHeader);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 1, 2, 'Supplier Invoice #', styles.tableHeader);
    sheet.mergeRowValue(sheet.row, 3, 5, 'Supplier', styles.tableHeader);
    sheet.mergeRowValue(sheet.row, 6, 7, 'PO #', styles.tableHeader);
    sheet.mergeRowValue(sheet.row, 8, 15, 'Description', styles.tableHeader);
    sheet.setCell(sheet.row, 16, 'Total Cost excl. Tax', styles.tableHeader);
    sheet.setCell(sheet.row, 17, 'Cost + Mark up', styles.tableHeader);
    sheet.nextRow();

    var start = sheet.row;

    for (var i = 0; i < rows.length; i++) {
      var r = sheet.row;
      var m = rows[i];

      if (m.documentNumber === 'TOTAL') {
        sheet.mergeRowValue(r, 1, 15, 'Total', styles.totalLabelRight);
        sheet.setFormula(r, 16, '=SUM(P' + start + ':P' + (r - 1) + ')', styles.totalCurrency);
        sheet.setFormula(r, 17, '=SUM(Q' + start + ':Q' + (r - 1) + ')', styles.totalCurrency);
      } else {
        sheet.mergeRowValue(r, 1, 2, m.documentNumber, styles.cell);
        sheet.mergeRowValue(r, 3, 5, m.mainName, styles.cell);
        sheet.mergeRowValue(r, 6, 7, m.cleanedPO, styles.cell);
        sheet.mergeRowValue(r, 8, 15, m.memo, styles.cell);
        sheet.setCell(r, 16, parseFloat(m.cost || 0), styles.currency);
        sheet.setCell(r, 17, parseFloat(m.amount || 0), styles.currency);
      }
      sheet.nextRow();
    }
  }

  function addExpensesSection(sheet, rows) {
    sheet.mergeRowValue(sheet.row, 1, 17, 'Expenses', styles.tableHeader);
    sheet.nextRow();

    sheet.mergeRowValue(sheet.row, 1, 5, 'Expense Category', styles.tableHeader);
    sheet.mergeRowValue(sheet.row, 6, 7, 'PO #', styles.tableHeader);
    sheet.mergeRowValue(sheet.row, 8, 15, 'Description', styles.tableHeader);
    sheet.setCell(sheet.row, 16, 'Total Cost excl. Tax', styles.tableHeader);
    sheet.setCell(sheet.row, 17, 'Cost + Mark up', styles.tableHeader);
    sheet.nextRow();

    var start = sheet.row;

    for (var i = 0; i < rows.length; i++) {
      var r = sheet.row;
      var e = rows[i];

      if (e.documentNumber === 'TOTAL') {
        sheet.mergeRowValue(r, 1, 15, 'Total', styles.totalLabelRight);
        sheet.setFormula(r, 16, '=SUM(P' + start + ':P' + (r - 1) + ')', styles.totalCurrency);
        sheet.setFormula(r, 17, '=SUM(Q' + start + ':Q' + (r - 1) + ')', styles.totalCurrency);
      } else {
        sheet.mergeRowValue(r, 1, 5, e.expCat, styles.cell);
        sheet.mergeRowValue(r, 6, 7, e.cleanedPO, styles.cell);
        sheet.mergeRowValue(r, 8, 15, e.memo, styles.cell);
        sheet.setCell(r, 16, parseFloat(e.cost || 0), styles.currency);
        sheet.setCell(r, 17, parseFloat(e.amount || 0), styles.currency);
      }
      sheet.nextRow();
    }
  }

  // --------------------------------------------------------------------------
  // SHEET BUILDER
  // --------------------------------------------------------------------------
  function createSheetBuilder(name) {
    return {
      name: name,
      row: 1,
      cells: {},
      merges: [],
      rowsMeta: {},

      setCell: function (r, c, v, s, t) {
        var ref = colLetter(c) + r;
        var cell = { v: v, s: s || styles.cell };
        if (t) {
          cell.t = t;
        } else {
          cell.t = inferType(v);
        }
        this.cells[ref] = cell;
      },

      setFormula: function (r, c, formula, s) {
        var ref = colLetter(c) + r;
        this.cells[ref] = { f: stripEqual(formula), s: s || styles.currency };
      },

      merge: function (r1, c1, r2, c2) {
        this.merges.push({
          s: { r: r1 - 1, c: c1 - 1 },
          e: { r: r2 - 1, c: c2 - 1 }
        });
      },

      mergeRowValue: function (r, c1, c2, v, s) {
        this.setCell(r, c1, v, s);
        if (c2 > c1) this.merge(r, c1, r, c2);
      },

      mergeRowFormula: function (r, c1, c2, formula, s) {
        this.setFormula(r, c1, formula, s);
        if (c2 > c1) this.merge(r, c1, r, c2);
      },

      nextRow: function () {
        this.row++;
      },

      blankRow: function (count) {
        count = count || 1;
        this.row += count;
      },

      toWorksheet: function () {
        var ws = {};
        var refs = Object.keys(this.cells);
        for (var i = 0; i < refs.length; i++) {
          ws[refs[i]] = this.cells[refs[i]];
        }
        ws['!merges'] = this.merges;
        ws['!ref'] = buildRefFromCells(this.cells);
        return ws;
      }
    };
  }

  // --------------------------------------------------------------------------
  // STYLES
  // --------------------------------------------------------------------------
  var styles = {
    invoiceTitle: {
      font: { bold: true, sz: 24 },
      alignment: { horizontal: 'left', vertical: 'center' },
      border: fullBorder()
    },
    sectionTitle: {
      font: { bold: true, sz: 20 },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: fullBorder()
    },
    logoBox: {
      border: fullBorder(),
      alignment: { horizontal: 'center', vertical: 'center' }
    },
    blackDivider: {
      fill: { fgColor: { rgb: '000000' } },
      border: fullBorder()
    },
    tableHeader: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '3A4B87' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: fullBorder()
    },
    tableHeaderLeft: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '3A4B87' } },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      border: fullBorder()
    },
    tableHeaderDate: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '3A4B87' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: fullBorder(),
      numFmt: 'dd-mmm-yyyy'
    },
    infoHeader: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '00A3E0' } },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      border: fullBorder()
    },
    infoValue: {
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      border: fullBorder()
    },
    wrapBox: {
      alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
      border: fullBorder()
    },
    centerBox: {
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: fullBorder()
    },
    cell: {
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      border: fullBorder()
    },
    cellLeft: {
      alignment: { horizontal: 'left', vertical: 'center' },
      border: fullBorder()
    },
    decimal: {
      alignment: { horizontal: 'right', vertical: 'center' },
      border: fullBorder(),
      numFmt: '0.0'
    },
    percent: {
      alignment: { horizontal: 'right', vertical: 'center' },
      border: fullBorder(),
      numFmt: '0.00%'
    },
    currency: {
      alignment: { horizontal: 'right', vertical: 'center' },
      border: fullBorder(),
      numFmt: '$#,##0.00'
    },
    labelRight: {
      font: { bold: true },
      alignment: { horizontal: 'right', vertical: 'center' },
      border: fullBorder()
    },
    totalLabel: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '3A4B87' } },
      alignment: { horizontal: 'right', vertical: 'center' },
      border: fullBorder()
    },
    totalLabelRight: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '3A4B87' } },
      alignment: { horizontal: 'right', vertical: 'center' },
      border: fullBorder()
    },
    totalCurrency: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '3A4B87' } },
      alignment: { horizontal: 'right', vertical: 'center' },
      border: fullBorder(),
      numFmt: '$#,##0.00'
    },
    totalDecimal: {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '3A4B87' } },
      alignment: { horizontal: 'right', vertical: 'center' },
      border: fullBorder(),
      numFmt: '0.0'
    },
    noBorder: {
      border: noBorder()
    },
    legend: {
      font: { italic: true },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: 'CCCCCC' } }
      }
    }
  };

  // --------------------------------------------------------------------------
  // DATA BUILDERS
  // --------------------------------------------------------------------------
  function normalizeTranIds(tranid) {
    if (!tranid) return [];
    var arr;
    if (Object.prototype.toString.call(tranid) === '[object Array]') {
      arr = tranid;
    } else {
      tranid = String(tranid);
      arr = tranid.indexOf(',') !== -1 ? tranid.split(',') : [tranid];
    }

    var unique = [];
    for (var i = 0; i < arr.length; i++) {
      var id = String(arr[i] || '').trim();
      if (id && unique.indexOf(id) === -1) unique.push(id);
    }
    return unique;
  }

  function buildConsolidatedInvoiceData(tranIds) {
    var firstSoId = tranIds[0];
    var firstSO = record.load({ type: record.Type.SALES_ORDER, id: firstSoId });
    var subId = firstSO.getValue({ fieldId: 'subsidiary' });
    var subrec = record.load({ type: 'subsidiary', id: subId });

    var logo = subrec.getValue('logo');
    var logoUrl = '';
    if (logo) {
      logoUrl = ('https://9873410-sb1.app.netsuite.com' + file.load({ id: logo }).url).replace(/&/g, '&amp;');
    }

    var soNumbers = [];
    var poNumbers = [];
    var projectRefs = [];
    var memos = [];
    var customerNames = [];

    var billAddrHtml = escBr(firstSO.getValue({ fieldId: 'billaddress' }) || '');
    var dueDate = esc(firstSO.getText({ fieldId: 'duedate' }) || firstSO.getValue({ fieldId: 'duedate' }) || '');
    var terms = esc(firstSO.getText({ fieldId: 'terms' }) || '');
    var currencyText = esc(firstSO.getText({ fieldId: 'currency' }) || '');
    var soDate = esc(firstSO.getText({ fieldId: 'trandate' }) || firstSO.getValue({ fieldId: 'trandate' }) || '');

    var region = subrec.getText({ fieldId: 'custrecord_c2o_region' }) || subrec.getValue({ fieldId: 'custrecord_c2o_region' }) || '';
    var isAmericas = (region === 'C2O Americas');

    var TAX_LABEL_RATE = isAmericas ? 'TAX RATE' : 'GST RATE';
    var TAX_LABEL_AMT = isAmericas ? 'TAX AMT' : 'GST AMT';
    var TAX_LABEL_TOTAL = isAmericas ? 'TAX TOTAL' : 'GST TOTAL';

    var subAddrHtml = escBr(subrec.getValue('mainaddress_text') || '');
    var subABN = esc(subrec.getValue('federalidnumber') || '');
    var remitEmail = esc(subrec.getValue('custrecord_bc_remittance_email') || '');
    var acctName = esc(subrec.getValue('custrecord_bc_account_name') || '');
    var bankName = esc(subrec.getValue('custrecord_bc_bank') || '');
    var bsb = esc(subrec.getValue('custrecord_bc_bsb') || '');
    var acctNum = esc(subrec.getValue('custrecord_bc_acc_num') || '');

    var catMap = {};
    var replaceLabor = false;

    for (var s = 0; s < tranIds.length; s++) {
      var soId = tranIds[s];
      var so = record.load({ type: record.Type.SALES_ORDER, id: soId });

      var soNum = so.getValue({ fieldId: 'tranid' }) || '';
      var otherRef = so.getValue({ fieldId: 'otherrefnum' }) || '';
      var memo = so.getValue({ fieldId: 'memo' }) || '';
      var entityText = so.getText({ fieldId: 'entity' }) || '';
      var projText = so.getText({ fieldId: 'cseg_bc_project' }) || '';
      var lineCount = so.getLineCount({ sublistId: 'item' }) || 0;

      if (soNum && soNumbers.indexOf(soNum) === -1) soNumbers.push(soNum);
      if (otherRef && poNumbers.indexOf(otherRef) === -1) poNumbers.push(otherRef);
      if (projText && projectRefs.indexOf(projText) === -1) projectRefs.push(projText);
      if (entityText && customerNames.indexOf(entityText) === -1) customerNames.push(entityText);
      if (memo && memos.indexOf(memo) === -1) memos.push(memo);

      try {
        var countryTxt = subrec.getText({ fieldId: 'country' }) || subrec.getValue({ fieldId: 'country' }) || '';
        if (countryTxt === 'Australia') replaceLabor = true;
      } catch (e1) {}

      for (var i = 0; i < lineCount; i++) {
        var categoryId = so.getSublistText({ sublistId: 'item', fieldId: 'custcol_invoicing_category', line: i });
        var relatedTimeId = so.getSublistValue({ sublistId: 'item', fieldId: 'custcol_bc_tm_time_bill', line: i });
        var relatedTranId = so.getSublistValue({ sublistId: 'item', fieldId: 'custcol_bc_tm_source_transaction', line: i });
        if (!categoryId || (!relatedTimeId && !relatedTranId)) continue;

        var lineAmt = parseFloat(so.getSublistValue({ sublistId: 'item', fieldId: 'amount', line: i }) || 0) || 0;
        var taxRateVal = parseFloat(so.getSublistValue({ sublistId: 'item', fieldId: 'taxrate1', line: i }) || 0) || 0;
        var taxAmtVal = parseFloat(so.getSublistValue({ sublistId: 'item', fieldId: 'tax1amt', line: i }) || 0) || 0;

        if (!catMap[categoryId]) {
          catMap[categoryId] = {
            amountSum: 0,
            taxAmtSum: 0,
            taxRateMax: 0
          };
        }

        catMap[categoryId].amountSum += lineAmt;
        catMap[categoryId].taxAmtSum += taxAmtVal;
        if (taxRateVal > catMap[categoryId].taxRateMax) catMap[categoryId].taxRateMax = taxRateVal;
      }
    }

    return {
      tranIds: tranIds,
      firstSoId: firstSoId,
      logoUrl: logoUrl,
      subId: subId,
      customerNames: customerNames,
      poNumbers: poNumbers,
      soNumbers: soNumbers,
      projectRefs: projectRefs,
      memos: memos,
      replaceLabor: replaceLabor,
      billAddrHtml: billAddrHtml,
      dueDate: dueDate,
      terms: terms,
      currencyText: currencyText,
      soDate: soDate,
      isAmericas: isAmericas,
      TAX_LABEL_RATE: TAX_LABEL_RATE,
      TAX_LABEL_AMT: TAX_LABEL_AMT,
      TAX_LABEL_TOTAL: TAX_LABEL_TOTAL,
      subAddrHtml: subAddrHtml,
      subABN: subABN,
      remitEmail: remitEmail,
      acctName: acctName,
      bankName: bankName,
      bsb: bsb,
      acctNum: acctNum,
      catMap: catMap
    };
  }

  function buildTimesheetData(soId, empNameArr) {
    var salesorderRec = record.load({ type: record.Type.SALES_ORDER, id: soId });
    var subId = salesorderRec.getValue({ fieldId: 'subsidiary' });
    var subrec = record.load({ type: 'subsidiary', id: subId });

    var logo = subrec.getValue('logo');
    var logoUrl = '';
    if (logo) {
      logoUrl = ('https://9873410-sb1.app.netsuite.com' + file.load({ id: logo }).url).replace(/&/g, '&amp;');
    }

    var replaceLabor = false;
    try {
      var countryTxt = subrec.getText({ fieldId: 'country' }) || subrec.getValue({ fieldId: 'country' }) || '';
      if (countryTxt === 'Australia') replaceLabor = true;
    } catch (e0) {}

    var projectId = salesorderRec.getValue({ fieldId: 'cseg_bc_project' }) || '';
    var projectName = '';
    var projectManager = '';
    var reportingProject = '';

    if (projectId) {
      try {
        var projectRec = record.load({ type: 'customrecord_cseg_bc_project', id: projectId });
        reportingProject = projectRec.getText({ fieldId: 'cseg_c2o_rep_proj' }) || '';
        projectManager = projectRec.getText({ fieldId: 'custrecord_bc_proj_manager' }) || '';
        projectName = projectRec.getText({ fieldId: 'name' }) || '';
      } catch (e1) {
        log.error('Project Load Failed', e1.message);
      }
    }

    var headerInfo = {
      soId: soId,
      client: salesorderRec.getText({ fieldId: 'entity' }) || '',
      customerRef: salesorderRec.getValue({ fieldId: 'otherrefnum' }) || '',
      weekEnding: salesorderRec.getText({ fieldId: 'trandate' }) || '',
      docNumber: salesorderRec.getValue({ fieldId: 'tranid' }) || '',
      description: salesorderRec.getValue({ fieldId: 'memo' }) || '',
      supervisor: salesorderRec.getText({ fieldId: 'custbody_client_supervisor' }) || '',
      startTime: salesorderRec.getText({ fieldId: 'custbody_start_time' }) || '',
      endTime: salesorderRec.getText({ fieldId: 'custbody_end_time' }) || '',
      projectName: projectName,
      projectManager: projectManager,
      reportingProject: reportingProject,
      logoUrl: logoUrl
    };

    var groupedFinalArray = buildTimesheetHourGroups(soId, empNameArr, replaceLabor);
    groupedFinalArray = mergeTimesheetSourceTransactionGroups(soId, groupedFinalArray, replaceLabor);
    var legendArray = buildLegendArray(groupedFinalArray, replaceLabor);

    return {
      soId: soId,
      replaceLabor: replaceLabor,
      headerInfo: headerInfo,
      groupedData: groupedFinalArray,
      legendArray: legendArray
    };
  }

  function buildTimesheetHourGroups(soId, empNameArr, replaceLabor) {
    var shiftSortOrder = ['ST', 'OT', 'OT1.5', 'DT', 'NT', 'RDO'];

    var salesorderSearchObj = search.create({
      type: 'salesorder',
      settings: [{ name: 'consolidationtype', value: 'NONE' }],
      filters: [
        ['type', 'anyof', 'SalesOrd'],
        'AND',
        ['custcol_bc_tm_time_bill', 'noneof', '@NONE@'],
        'AND',
        ['internalid', 'anyof', [soId]]
      ],
      columns: [
        search.createColumn({ name: 'custcol_invoicing_category', summary: 'GROUP' }),
        search.createColumn({ name: 'employee', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }),
        search.createColumn({ name: 'durationdecimal', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'SUM' }),
        search.createColumn({
          name: 'formulatext1',
          formula: "NVL(NVL(NVL({custcol_c2o_billing_class_override},{custcol_bc_tm_time_bill.custcol_bc_tm_labor_billing_class}), {custcol_bc_tm_source_transaction.memo}),'')",
          summary: 'GROUP'
        }),
        search.createColumn({ name: 'memo', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }),
        search.createColumn({ name: 'custcol_bc_time_type', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }),
        search.createColumn({ name: 'custcol_bc_tm_billing_shift', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }),
        search.createColumn({ name: 'date', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP', sort: search.Sort.ASC }),
        search.createColumn({ name: 'memo', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }),
        search.createColumn({ name: 'trandate', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }),
        search.createColumn({ name: 'quantity', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'SUM' }),
        search.createColumn({ name: 'rate', summary: 'MAX' }),
        search.createColumn({ name: 'amount', summary: 'SUM' })
      ]
    });

    var employeeMap = {};
    var uniqueDates = {};

    salesorderSearchObj.run().each(function (result) {
      var billRentalRole = result.getValue({ name: 'memo', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' });
      billRentalRole = (billRentalRole === '- None -' || !billRentalRole) ? '' : billRentalRole;

      var empName = result.getValue({ name: 'employee', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }) || 1039;
      var role = result.getValue({ name: 'formulatext1', summary: 'GROUP' });
      role = (role === '- None -' || !role) ? '' : role;

      var shiftType = result.getText({ name: 'custcol_bc_time_type', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }) || '';
      var dateStr = result.getValue({ name: 'date', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }) ||
                    result.getValue({ name: 'trandate', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' });

      var hours = parseFloat(result.getValue({ name: 'durationdecimal', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'SUM' })) ||
                  parseFloat(result.getValue({ name: 'quantity', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'SUM' })) || 0;

      var note = result.getValue({ name: 'memo', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }) || '';
      var groupType = result.getText({ name: 'custcol_invoicing_category', summary: 'GROUP' }) || '';
      var shift = result.getText({ name: 'custcol_bc_tm_billing_shift', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }) || '';
      var rateTime = parseFloat(result.getValue({ name: 'rate', summary: 'MAX' }) || 0) || 0;
      var amtTime = parseFloat(result.getValue({ name: 'amount', summary: 'SUM' }) || 0) || 0;

      uniqueDates[dateStr] = true;

      var empKey = empName + '_' + shiftType + '_' + role + '_' + groupType + '_' + shift;
      if (!employeeMap[empKey]) {
        employeeMap[empKey] = {
          employee: empNameArr[empName] || '',
          role: role || billRentalRole,
          shiftType: (shiftType === '- None -') ? '' : shiftType,
          shift: (shift === '- None -') ? '' : String(shift).replace(' Time', ''),
          dateMap: {},
          totalWeek: 0,
          rate: rateTime,
          amt: 0,
          notes: '',
          groupType: groupType
        };
      }

      employeeMap[empKey].dateMap[dateStr] = hours.toFixed(1);
      employeeMap[empKey].totalWeek += hours;
      employeeMap[empKey].amt += amtTime;
      if (note) employeeMap[empKey].notes += (employeeMap[empKey].notes ? ' | ' : '') + note;

      return true;
    });

    var sortedDates = Object.keys(uniqueDates).sort(function (a, b) {
      return new Date(a) - new Date(b);
    });

    var headerRow = {
      employee: 'Name',
      role: 'Role',
      shiftType: 'Shift Type',
      shift: 'Shift',
      days: [],
      totalWeek: 'TOTAL WEEK',
      rate: 'Rate',
      amt: 'Claim Amount',
      notes: 'Notes',
      groupType: 'Group Type'
    };

    for (var hd = 0; hd < sortedDates.length; hd++) {
      headerRow.days.push({ date: sortedDates[hd] });
    }

    var groupedFinalArray = {};
    var finalAmtByGroup = {};

    var employeeKeys = Object.keys(employeeMap);
    for (var i = 0; i < employeeKeys.length; i++) {
      var emp = employeeMap[employeeKeys[i]];
      var row = {
        employee: escPlain(emp.employee),
        role: escPlain(emp.role),
        shiftType: escPlain(emp.shiftType),
        shift: escPlain(emp.shift),
        days: [],
        totalWeek: emp.totalWeek.toFixed(1),
        rate: emp.rate.toFixed(2),
        amt: emp.amt.toFixed(2),
        notes: escPlain(emp.notes),
        groupType: escPlain(emp.groupType)
      };

      for (var d = 0; d < sortedDates.length; d++) {
        var dt = sortedDates[d];
        row.days.push({
          date: dt,
          hours: emp.dateMap[dt] || 0
        });
      }

      if (!groupedFinalArray[emp.groupType]) {
        groupedFinalArray[emp.groupType] = [cloneHeaderRow(headerRow)];
        finalAmtByGroup[emp.groupType] = 0;
      }

      finalAmtByGroup[emp.groupType] += emp.amt;
      groupedFinalArray[emp.groupType].push(row);
    }

    var groupNames = Object.keys(groupedFinalArray);
    for (var g = 0; g < groupNames.length; g++) {
      var group = groupNames[g];
      var header = groupedFinalArray[group].shift();

      var sortedGroup = groupedFinalArray[group].sort(function (a, b) {
        var empA = String(a.employee || '').toLowerCase();
        var empB = String(b.employee || '').toLowerCase();
        if (empA !== empB) return empA < empB ? -1 : 1;

        var indexA = shiftSortOrder.indexOf(a.shiftType);
        var indexB = shiftSortOrder.indexOf(b.shiftType);
        if (indexA === -1) indexA = 999;
        if (indexB === -1) indexB = 999;
        return indexA - indexB;
      });

      var totalRow = {
        employee: 'TOTAL',
        role: '',
        shiftType: '',
        shift: '',
        days: [],
        totalWeek: 0,
        rate: '',
        amt: formatCurrency(finalAmtByGroup[group] || 0),
        notes: '',
        groupType: group
      };

      for (var td = 0; td < sortedDates.length; td++) {
        var date = sortedDates[td];
        var dateSum = 0;
        for (var sr = 0; sr < sortedGroup.length; sr++) {
          var rowData = sortedGroup[sr];
          for (var dy = 0; dy < rowData.days.length; dy++) {
            if (rowData.days[dy].date === date) {
              dateSum += parseFloat(rowData.days[dy].hours || 0) || 0;
              break;
            }
          }
        }

        totalRow.days.push({ date: date, hours: dateSum.toFixed(1) });
        totalRow.totalWeek = parseFloat(totalRow.totalWeek) + parseFloat(dateSum);
      }

      totalRow.totalWeek = parseFloat(totalRow.totalWeek).toFixed(1);
      groupedFinalArray[group] = [header].concat(sortedGroup).concat([totalRow]);
    }

    if (replaceLabor) groupedFinalArray = replaceLaborText(groupedFinalArray);
    return groupedFinalArray;
  }

  function mergeTimesheetSourceTransactionGroups(soId, groupedFinalArray, replaceLabor) {
    var transactionSearch = search.create({
      type: 'salesorder',
      settings: [{ name: 'consolidationtype', value: 'NONE' }],
      filters: [
        ['type', 'anyof', 'SalesOrd'],
        'AND',
        ['custcol_bc_tm_source_transaction', 'noneof', '@NONE@'],
        'AND',
        ['internalid', 'anyof', [soId]],
        'AND',
        ['formulatext: case when {custcol_bc_tm_line_id} = {custcol_bc_tm_source_transaction.line} then 1 else 0 end', 'is', '1']
      ],
      columns: [
        search.createColumn({ name: 'custcol_invoicing_category', summary: 'GROUP' }),
        search.createColumn({ name: 'tranid', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }),
        search.createColumn({ name: 'mainname', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }),
        search.createColumn({ name: 'amount', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'MAX' }),
        search.createColumn({ name: 'taxamount', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'MAX' }),
        search.createColumn({ name: 'memo', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }),
        search.createColumn({
          name: 'formulatext',
          summary: 'MAX',
          formula: "CASE WHEN {custcol_bc_tm_source_transaction.appliedtotransaction} LIKE 'Purchase Order%' THEN TRIM(REPLACE({custcol_bc_tm_source_transaction.appliedtotransaction}, 'Purchase Order', '')) ELSE {custcol_bc_tm_source_transaction.tranid} END"
        }),
        search.createColumn({ name: 'amount', summary: 'MAX' }),
        search.createColumn({ name: 'expensecategory', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' })
      ]
    });

    var tranFinalArray = {};

    transactionSearch.run().each(function (result) {
      var invoicingCategory = result.getText({ name: 'custcol_invoicing_category', summary: 'GROUP' }) || 'Uncategorized';
      var docNumber = result.getValue({ name: 'tranid', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }) || '';
      var mainName = result.getText({ name: 'mainname', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }) || '';
      var expCat = result.getText({ name: 'expensecategory', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }) || '';
      var cost = parseFloat(result.getValue({ name: 'amount', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'MAX' }) || 0) || 0;
      var tax = parseFloat(result.getValue({ name: 'taxamount', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'MAX' }) || 0) || 0;
      var amount = parseFloat(result.getValue({ name: 'amount', summary: 'MAX' }) || 0) || 0;
      var memo = result.getValue({ name: 'memo', join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION', summary: 'GROUP' }) || '';
      var cleanedPO = result.getValue({ name: 'formulatext', summary: 'MAX' }) || '';

      var row = {
        documentNumber: escPlain(docNumber),
        mainName: escPlain(mainName),
        expCat: escPlain(expCat === '- None -' ? mainName : expCat),
        amount: amount.toFixed(2),
        cost: cost.toFixed(2),
        tax: tax.toFixed(2),
        memo: escPlain(memo).replace(/:/g, '-'),
        cleanedPO: (cleanedPO === '- None -' ? '' : escPlain(cleanedPO))
      };

      if (!tranFinalArray[invoicingCategory]) {
        tranFinalArray[invoicingCategory] = {
          rows: [],
          totalAmount: 0,
          totalCost: 0,
          totalTax: 0
        };
      }

      tranFinalArray[invoicingCategory].rows.push(row);
      tranFinalArray[invoicingCategory].totalAmount += amount;
      tranFinalArray[invoicingCategory].totalCost += cost;
      tranFinalArray[invoicingCategory].totalTax += tax;
      return true;
    });

    var categories = Object.keys(tranFinalArray);
    for (var i = 0; i < categories.length; i++) {
      var category = categories[i];
      var group = tranFinalArray[category];

      var totalRow = {
        documentNumber: 'TOTAL',
        mainName: '',
        expCat: '',
        amount: formatCurrency(group.totalAmount || 0),
        cost: formatCurrency(group.totalCost || 0),
        tax: formatCurrency(Math.abs(group.totalTax || 0)),
        total: formatCurrency((Math.abs(group.totalTax || 0)) + (group.totalCost || 0)),
        memo: '',
        cleanedPO: ''
      };

      groupedFinalArray[category] = group.rows.concat([totalRow]);
    }

    if (replaceLabor) groupedFinalArray = replaceLaborText(groupedFinalArray);
    return groupedFinalArray;
  }

  function buildLegendArray(groupedFinalArray, replaceLabor) {
    var TIME_LEGEND_MAP = {
      'ST': 'Standard Time',
      'OT': 'Overtime',
      'DT': 'Double Time',
      'PT': 'Part Time',
      'PTO': 'Paid Time Off',
      'Per Diem': 'Per Diem Allowance',
      'DR1': 'Day Rate 1',
      'DR2': 'Day Rate 2',
      'DR3': 'Day Rate 3'
    };

    var seenTypes = {};
    var legendArray = [];
    var categories = Object.keys(groupedFinalArray);

    for (var i = 0; i < categories.length; i++) {
      var rows = groupedFinalArray[categories[i]];
      for (var r = 0; r < rows.length; r++) {
        if (rows[r].shiftType) {
          var cleanType = String(rows[r].shiftType).trim();
          if (TIME_LEGEND_MAP[cleanType] && !seenTypes[cleanType]) {
            seenTypes[cleanType] = true;
            legendArray.push({ abbr: cleanType, label: TIME_LEGEND_MAP[cleanType] });
          }
        }
      }
    }

    var order = ['ST', 'OT', 'DT', 'PT', 'PTO', 'Per Diem', 'DR1', 'DR2', 'DR3'];
    legendArray.sort(function (a, b) {
      return order.indexOf(a.abbr) - order.indexOf(b.abbr);
    });

    if (replaceLabor) {
      for (var j = 0; j < legendArray.length; j++) {
        legendArray[j].label = String(legendArray[j].label || '').replace(/\bLabor\b/g, 'Labour');
      }
    }

    return legendArray;
  }

  function replaceLaborText(groupedFinalArray) {
    var newObj = {};
    var keys = Object.keys(groupedFinalArray);

    for (var i = 0; i < keys.length; i++) {
      var oldKey = keys[i];
      var newKey = oldKey === 'Labor' ? 'Labour' : oldKey;
      var rows = groupedFinalArray[oldKey];

      for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        var rowKeys = Object.keys(row);
        for (var k = 0; k < rowKeys.length; k++) {
          var rk = rowKeys[k];
          if (typeof row[rk] === 'string') {
            row[rk] = row[rk].replace(/\bLabor\b/g, 'Labour');
          }
        }
      }

      newObj[newKey] = rows;
    }

    return newObj;
  }

  function getEmployeeList() {
    var returnObj = {};
    var employeeSearchObj = search.create({
      type: 'employee',
      filters: [],
      columns: [
        search.createColumn({ name: 'internalid' }),
        search.createColumn({ name: 'formulatext', formula: "{firstname} || ' ' || {lastname}" })
      ]
    });

    employeeSearchObj.run().each(function (result) {
      returnObj[String(result.getValue({ name: 'internalid' }))] =
        result.getValue({ name: 'formulatext' }) || '';
      return true;
    });

    return returnObj;
  }

  // --------------------------------------------------------------------------
  // HELPERS
  // --------------------------------------------------------------------------
  function cloneHeaderRow(headerRow) {
    return {
      employee: headerRow.employee,
      role: headerRow.role,
      shiftType: headerRow.shiftType,
      shift: headerRow.shift,
      days: headerRow.days.slice(0),
      totalWeek: headerRow.totalWeek,
      rate: headerRow.rate,
      amt: headerRow.amt,
      notes: headerRow.notes,
      groupType: headerRow.groupType
    };
  }

  function escBr(v) {
    if (v == null) return '';
    return String(v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/\r?\n/g, '<br>');
  }

  function esc(v) {
    if (v == null) return '';
    return String(v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/\r?\n/g, '<br />');
  }

  function escPlain(v) {
    if (v == null) return '';
    return String(v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatCurrency(amount) {
    if (amount === '' || amount == null) return '';
    if (String(amount).indexOf('$') !== -1) return amount;
    return '$ ' + parseFloat(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  function formatDateMMDDYYYY(dateStr) {
    if (!dateStr) return '';
    dateStr = String(dateStr);

    if (dateStr.indexOf('/') !== -1) {
      var parts = dateStr.split('/');
      if (parts.length === 3) {
        var m = parseInt(parts[0], 10) - 1;
        var d = parseInt(parts[1], 10);
        var yStr = parts[2];
        if (yStr.length === 2) yStr = '20' + yStr;
        var y = parseInt(yStr, 10);
        var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        return d + '-' + months[m] + '-' + y;
      }
    }

    var dt = new Date(dateStr);
    if (!isNaN(dt)) {
      var months2 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return dt.getDate() + '-' + months2[dt.getMonth()] + '-' + dt.getFullYear();
    }
    return dateStr;
  }

  function getDayName(dateStr) {
    var date = new Date(dateStr);
    if (isNaN(date)) return '';
    var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    return days[date.getDay()];
  }

  function colLetter(colNum) {
    var temp = '';
    var letter = '';
    while (colNum > 0) {
      temp = (colNum - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      colNum = (colNum - temp - 1) / 26;
    }
    return letter;
  }

  function buildRefFromCells(cells) {
    var refs = Object.keys(cells);
    if (!refs.length) return 'A1';

    var minCol = 9999, minRow = 999999, maxCol = 1, maxRow = 1;

    for (var i = 0; i < refs.length; i++) {
      var rc = refToRC(refs[i]);
      if (rc.c < minCol) minCol = rc.c;
      if (rc.r < minRow) minRow = rc.r;
      if (rc.c > maxCol) maxCol = rc.c;
      if (rc.r > maxRow) maxRow = rc.r;
    }

    return colLetter(minCol) + minRow + ':' + colLetter(maxCol) + maxRow;
  }

  function refToRC(ref) {
    var m = ref.match(/^([A-Z]+)(\d+)$/);
    var col = m[1], row = parseInt(m[2], 10), c = 0;
    for (var i = 0; i < col.length; i++) {
      c = c * 26 + (col.charCodeAt(i) - 64);
    }
    return { r: row, c: c };
  }

  function stripEqual(f) {
    return String(f || '').replace(/^=/, '');
  }

  function inferType(v) {
    if (v == null || v === '') return 's';
    if (typeof v === 'number') return 'n';
    if (typeof v === 'boolean') return 'b';
    return 's';
  }

  function stripHtmlBreaks(v) {
    return String(v || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'");
  }

  function fullBorder() {
    return {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } }
    };
  }

  function noBorder() {
    return {
      top: { style: 'none' },
      bottom: { style: 'none' },
      left: { style: 'none' },
      right: { style: 'none' }
    };
  }

  return {
    onRequest: onRequest
  };
});