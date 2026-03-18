/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define([
  'N/ui/serverWidget',
  'N/search',
  'N/log',
  'N/file',
  'N/encode',
  'N/runtime',
  'N/record'
], function (ui, search, log, file, encode, runtime, record) {

  function onRequest(context) {
    try {
      if (context.request.method !== 'GET') return;

      var params = context.request.parameters || {};
      var tranid = params.tranid;
      var exportType = params.export || '';

      var tranIds = normalizeTranIds(tranid);
      log.debug('Normalized tranIds', tranIds);

      if (!tranIds || !tranIds.length) {
        context.response.write('No tranid provided.');
        return;
      }

      var empNameArr = getEmployeeList();

      // 1) Consolidated draft invoice data for ALL tranids together
      var invoiceData = buildConsolidatedInvoiceData(tranIds);

      // 2) Timesheet data PER tranid
      var timesheetDataByTran = [];
      for (var i = 0; i < tranIds.length; i++) {
        timesheetDataByTran.push(buildTimesheetData(tranIds[i], empNameArr));
      }

      if (exportType === 'excel') {
        var html = buildExcelHtml(invoiceData, timesheetDataByTran);

        var excelFile = file.create({
          name: 'Weekly_Timesheet_' + new Date().toISOString().slice(0, 10) + '.xls',
          fileType: file.Type.PLAINTEXT,
          contents: html,
          encoding: file.Encoding.UTF_8
        });

        context.response.writeFile(excelFile, false);
        return;
      }

      var responseObj = {
        invoice: invoiceData,
        timesheets: timesheetDataByTran
      };

      var strReturn = "<#assign ObjDetail=" + JSON.stringify(responseObj) + " />";
      context.response.writeLine(strReturn);

    } catch (e) {
      log.error('Error running Suitelet', e);
      context.response.write('Error: ' + e.message);
    }
  }

  // --------------------------------------------------------------------------
  // NORMALIZE INPUT
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
      if (id && unique.indexOf(id) === -1) {
        unique.push(id);
      }
    }
    return unique;
  }

  // --------------------------------------------------------------------------
  // CONSOLIDATED DRAFT INVOICE
  // --------------------------------------------------------------------------
  function buildConsolidatedInvoiceData(tranIds) {
    var firstSoId = tranIds[0];
    var firstSO = record.load({ type: record.Type.SALES_ORDER, id: firstSoId });

    var subId = firstSO.getValue({ fieldId: 'subsidiary' });
    var subrec = record.load({ type: 'subsidiary', id: subId });

    var logo = subrec.getValue('logo');
    var logoUrl = '';
    if (logo) {
      logoUrl = ('https://9873410-sb1.app.netsuite.com' + file.load({ id: logo }).url).replace(/&/g, '&amp;');
    } else {
      logoUrl = 'https://9873410-sb1.app.netsuite.com/core/media/media.nl?id=11486&amp;c=9873410_SB1&amp;h=1hbkOLk3U5GSjdY4GjdiGdKUZDkL4wsovPepc9ocNenvsfSW';
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

    var subLegal = esc(subrec.getValue('legalname') || subrec.getValue('name') || '');
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

      if (soNumbers.indexOf(soNum) === -1 && soNum) soNumbers.push(soNum);
      if (poNumbers.indexOf(otherRef) === -1 && otherRef) poNumbers.push(otherRef);
      if (projectRefs.indexOf(projText) === -1 && projText) projectRefs.push(projText);
      if (customerNames.indexOf(entityText) === -1 && entityText) customerNames.push(entityText);
      if (memos.indexOf(memo) === -1 && memo) memos.push(memo);

      // detect AU from this SO subsidiary if needed
      try {
        var currentSubId = so.getValue({ fieldId: 'subsidiary' });
        if (String(currentSubId) === String(subId)) {
          var countryTxt = subrec.getText({ fieldId: 'country' }) || subrec.getValue({ fieldId: 'country' }) || '';
          if (countryTxt === 'Australia') replaceLabor = true;
        }
      } catch (e1) {}

      for (var i = 0; i < lineCount; i++) {
        var categoryId = so.getSublistText({ sublistId: 'item', fieldId: 'custcol_invoicing_category', line: i });
        var relatedTimeId = so.getSublistValue({ sublistId: 'item', fieldId: 'custcol_bc_tm_time_bill', line: i });
        var relatedTranId = so.getSublistValue({ sublistId: 'item', fieldId: 'custcol_bc_tm_source_transaction', line: i });
        if (!categoryId || (!relatedTimeId && !relatedTranId)) continue;

        var lineAmt = parseFloat(so.getSublistValue({ sublistId: 'item', fieldId: 'amount', line: i }) || 0) || 0;
        var qtyVal = so.getSublistValue({ sublistId: 'item', fieldId: 'quantity', line: i });
        if (!lineAmt && (!qtyVal || qtyVal === 0 || qtyVal === '0')) continue;

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

        if (taxRateVal > catMap[categoryId].taxRateMax) {
          catMap[categoryId].taxRateMax = taxRateVal;
        }
      }
    }

    var rowsHtml = '';
    var subTotalExTax = 0;
    var totalTax = 0;
    var categories = Object.keys(catMap).sort();

    for (var c = 0; c < categories.length; c++) {
      var cat = categories[c];
      var amt = catMap[cat].amountSum;
      var txa = catMap[cat].taxAmtSum;
      var txr = catMap[cat].taxRateMax;

      subTotalExTax += amt;
      totalTax += txa;

      rowsHtml += ''
        + '<tr>'
        + '<td colspan="8" style="border:none;" >' + esc(cat) + '</td>'
        + '<td colspan="2" style="border:none;" align="right">' + money(amt) + '</td>'
        + '<td colspan="2" style="border:none;" align="center">' + pct(txr) + '</td>'
        + '<td colspan="2" style="border:none;" align="right">' + money(Math.abs(txa)) + '</td>'
        + '<td colspan="2" style="border:none;" align="right">' + money(amt + txa) + '</td>'
        + '</tr>';
    }

    var grandTotal = subTotalExTax + Math.abs(totalTax);

    // var invoiceBlockHtml = ''
    //   + '<table style="width:100%; border-collapse:collapse; font-family:Arial; border:1px solid #000;">'
    //   + '<tr>'
    //   + '<td colspan="12" rowspan="7" style="font-size:30pt; vertical-align:middle; font-weight:bold;">DRAFT INVOICE</td>'
    //   + '<td colspan="4" rowspan="7" align="right" style="vertical-align:middle; font-weight:bold;"><img src="' + logoUrl + '" height="100" /></td>'
    //   + '</tr>'
    //   + '</table>'

    //   + '<table style="width:100%; border-collapse:collapse; font-family:Arial; font-size:10pt; border:1px solid #000;">'
    //   + '<tr>'
    //   + '<td colspan="5" rowspan="4" valign="top">'
    //   + '<b>ATTN:</b><br/>'
    //   + billAddrHtml
    //   + '</td>'
    //   + '<td colspan="6" valign="top"><b>Invoice Date:</b><br/>' + soDate + '</td>'
    //   + '<td colspan="5" rowspan="4" align="right" valign="top">'
    //   + subAddrHtml + '<br/>'
    //   + '<b>ABN:</b> ' + subABN
    //   + '</td>'
    //   + '</tr>'

    //   + '<tr><td colspan="6" valign="top"><b>Invoice Number:</b><br/>DRAFT</td></tr>'
    //   + '<tr><td colspan="6" valign="top"><b>PO Number:</b><br/>' + esc(poNumbers.join(', ')) + '</td></tr>'
    //   + '<tr><td colspan="6" valign="top"><b>Customer Reference:</b><br/>' + esc(projectRefs.join(', ')) + '</td></tr>'

    //   + '<tr><td colspan="16">&nbsp;</td></tr>'
    //   + '<tr><td colspan="16"><b>Memo:</b><br/>' + esc(memos.join(' | ')) + '</td></tr>'
    //   + '<tr><td colspan="16">&nbsp;</td></tr>'

    //   + '<tr style="border-bottom:1px solid #000;">'
    //   + '<th class="table-header" colspan="8"><b>Description</b></th>'
    //   + '<th class="table-header" colspan="2" align="right"><b>Price</b></th>'
    //   + '<th class="table-header" colspan="2" align="center"><b>' + (isAmericas ? 'TAX' : 'GST') + '</b></th>'
    //   + '<th class="table-header" colspan="2" align="right"><b>' + TAX_LABEL_AMT + '</b></th>'
    //   + '<th class="table-header" colspan="2" align="right"><b>Amount ' + currencyText + '</b></th>'
    //   + '</tr>'

    //   + rowsHtml

    //   + '<tr><td colspan="10">&nbsp;</td></tr>'
    //   + '<tr><td rowspan = "3" colspan="12"></td><td colspan="2" align="right">Subtotal</td><td colspan="2" align="right">' + money(subTotalExTax) + '</td></tr>'
    //   + '<tr><td colspan="2" align="right">' + TAX_LABEL_TOTAL + '</td><td colspan="2" align="right">' + money(Math.abs(totalTax)) + '</td></tr>'
    //   + '<tr><td colspan="2" align="right" style="border-top:1px solid #000;"><b>TOTAL ' + currencyText + '</b></td><td colspan="2" align="right" style="border-top:1px solid #000;"><b>' + money(grandTotal) + '</b></td></tr>'

    //   + '<tr><td colspan="16">&nbsp;</td></tr>'
    //   + '<tr><td colspan="16">'
    //   + '<b>Sales Orders:</b> ' + esc(soNumbers.join(', ')) + '<br/><br/>'
    //   + '<b>Customer:</b> ' + esc(customerNames.join(', ')) + '<br/><br/>'
    //   + '<b>Due Date:</b> ' + dueDate + '<br/><br/>'
    //   + '<b>Payment Terms:</b> ' + terms + '<br/><br/>'
    //   + 'Please email remittance advice to ' + remitEmail + '<br/><br/>'
    //   + '<b>BANK ACCOUNT DETAILS</b><br/>'
    //   + 'Account Name: ' + acctName + '<br/>'
    //   + 'Bank: ' + bankName + '<br/>'
    //   + 'BSB: ' + bsb + '<br/>'
    //   + 'Account: ' + acctNum
    //   + '</td></tr>'
    //   + '</table>';


    var invoiceBlockHtml = ''
  + '<table style="width:100%; border-collapse:collapse; font-family:Arial; border:1px solid #000;">'
  + '<tr>'
  + '<td colspan="12" rowspan="7" style="font-size:30pt; vertical-align:middle; font-weight:bold; border:none;">DRAFT INVOICE</td>'
  + '<td colspan="4" rowspan="7" align="right" style="vertical-align:middle; font-weight:bold; border:none;"><img src="' + logoUrl + '" height="100" /></td>'
  + '</tr>'
  + '</table>'

  + '<table style="width:100%; border-collapse:collapse; font-family:Arial; font-size:10pt; border:1px solid #000;">'
  + '<tr>'
  + '<td colspan="5" rowspan="4" valign="top" style="border:none;">'
  + '<b>ATTN:</b><br/>'
  + billAddrHtml
  + '</td>'
  + '<td colspan="6" valign="top" style="border:none;"><b>Invoice Date:</b><br/>' + soDate + '</td>'
  + '<td colspan="5" rowspan="4" align="right" valign="top" style="border:none;">'
  + subAddrHtml + '<br/>'
  + '<b>ABN:</b> ' + subABN
  + '</td>'
  + '</tr>'

  + '<tr><td colspan="6" valign="top" style="border:none;"><b>Invoice Number:</b><br/>DRAFT</td></tr>'
  + '<tr><td colspan="6" valign="top" style="border:none;"><b>PO Number:</b><br/>' + esc(poNumbers.join(', ')) + '</td></tr>'
  + '<tr><td colspan="6" valign="top" style="border:none;"><b>Customer Reference:</b><br/>' + esc(projectRefs.join(', ')) + '</td></tr>'

  + '<tr><td colspan="16" style="border:none;">&nbsp;</td></tr>'
  + '<tr><td colspan="16" style="border:none;"><b>Memo:</b><br/>' + esc(memos.join(' | ')) + '</td></tr>'
  + '<tr><td colspan="16" style="border:none;">&nbsp;</td></tr>'

  + '<tr>'
  + '<th  colspan="8" style="border:none; border-bottom: 1px;"><b>Description</b></th>'
  + '<th  colspan="2" align="right" style="border:none; border-bottom: 1px;"><b>Price</b></th>'
  + '<th  colspan="2" align="center" style="border:none; border-bottom: 1px;"><b>' + (isAmericas ? 'TAX' : 'GST') + '</b></th>'
  + '<th  colspan="2" align="right" style="border:none; border-bottom: 1px;"><b>' + TAX_LABEL_AMT + '</b></th>'
  + '<th  colspan="2" align="right" style="border:none; border-bottom: 1px;"><b>Amount ' + currencyText + '</b></th>'
  + '</tr>'

  + rowsHtml

  + '<tr><td colspan="10" style="border:none;">&nbsp;</td></tr>'
  + '<tr><td rowspan="3" colspan="12" style="border:none;"></td><td colspan="2" align="right" style="border:none;">Subtotal</td><td colspan="2" align="right" style="border:none;">' + money(subTotalExTax) + '</td></tr>'
  + '<tr><td colspan="2" align="right" style="border:none;">' + TAX_LABEL_TOTAL + '</td><td colspan="2" align="right" style="border:none;">' + money(Math.abs(totalTax)) + '</td></tr>'
  + '<tr><td colspan="2" align="right" style="border:none;"><b>TOTAL ' + currencyText + '</b></td><td colspan="2" align="right" style="border:none;"><b>' + money(grandTotal) + '</b></td></tr>'

  + '<tr><td colspan="16" style="border:none;">&nbsp;</td></tr>'
  + '<tr><td colspan="16" style="border:none;">'
  + '<b>Sales Orders:</b> ' + esc(soNumbers.join(', ')) + '<br/><br/>'
  + '<b>Customer:</b> ' + esc(customerNames.join(', ')) + '<br/><br/>'
  + '<b>Due Date:</b> ' + dueDate + '<br/><br/>'
  + '<b>Payment Terms:</b> ' + terms + '<br/><br/>'
  + 'Please email remittance advice to ' + remitEmail + '<br/><br/>'
  + '<b>BANK ACCOUNT DETAILS</b><br/>'
  + 'Account Name: ' + acctName + '<br/>'
  + 'Bank: ' + bankName + '<br/>'
  + 'BSB: ' + bsb + '<br/>'
  + 'Account: ' + acctNum
  + '</td></tr>'
  + '</table>';

    return {
      tranIds: tranIds,
      firstSoId: firstSoId,
      logoUrl: logoUrl,
      subId: subId,
      rowsHtml: rowsHtml,
      invoiceBlockHtml: invoiceBlockHtml,
      customerNames: customerNames,
      poNumbers: poNumbers,
      soNumbers: soNumbers,
      projectRefs: projectRefs,
      memos: memos,
      subTotalExTax: subTotalExTax,
      totalTax: totalTax,
      grandTotal: grandTotal,
      replaceLabor: replaceLabor
    };
  }

  // --------------------------------------------------------------------------
  // TIMESHEET DATA BY SINGLE SO
  // --------------------------------------------------------------------------
  function buildTimesheetData(soId, empNameArr) {
    var salesorderRec = record.load({ type: record.Type.SALES_ORDER, id: soId });

    var subId = salesorderRec.getValue({ fieldId: 'subsidiary' });
    var subrec = record.load({ type: 'subsidiary', id: subId });

    var logo = subrec.getValue('logo');
    var logoUrl = '';
    if (logo) {
      logoUrl = ('https://9873410-sb1.app.netsuite.com' + file.load({ id: logo }).url).replace(/&/g, '&amp;');
    } else {
      logoUrl = 'https://9873410-sb1.app.netsuite.com/core/media/media.nl?id=11486&amp;c=9873410_SB1&amp;h=1hbkOLk3U5GSjdY4GjdiGdKUZDkL4wsovPepc9ocNenvsfSW';
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
        var projectRec = record.load({
          type: 'customrecord_cseg_bc_project',
          id: projectId
        });
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
    log.debug('groupedFinalArray', groupedFinalArray)
    groupedFinalArray = mergeTimesheetSourceTransactionGroups(soId, groupedFinalArray, replaceLabor);
log.debug('groupedFinalArray', groupedFinalArray)
    var legendArray = buildLegendArray(groupedFinalArray, replaceLabor);
log.debug('legendArray', legendArray)
    return {
      soId: soId,
      replaceLabor: replaceLabor,
      headerInfo: headerInfo,
      groupedData: groupedFinalArray,
      legendArray: legendArray
    };
  }

  function buildTimesheetHourGroups(soId, empNameArr, replaceLabor) {
    log.debug('soId', soId)
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
        search.createColumn({ name: 'item', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }),
        search.createColumn({
          name: 'formulatext1',
          formula: "NVL(NVL(NVL({custcol_c2o_billing_class_override},{custcol_bc_tm_time_bill.custcol_bc_tm_labor_billing_class}), {custcol_bc_tm_source_transaction.memo}),'')",
          summary: 'GROUP'
        }),
        search.createColumn({ name: 'memo', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }),
        search.createColumn({ name: 'custcol_bc_time_type', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }),
        search.createColumn({ name: 'custcol_bc_tm_billing_shift', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP' }),
        search.createColumn({ name: 'date', join: 'CUSTCOL_BC_TM_TIME_BILL', summary: 'GROUP', sort: search.Sort.ASC }),
        search.createColumn({
          name: 'custcol_bc_tm_line_id',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'memo',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'trandate',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'quantity',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'SUM'
        }),
        search.createColumn({
          name: 'rate',
          summary: 'MAX'
        }),
        search.createColumn({
          name: 'amount',
          summary: 'SUM'
        })
      ]
    });

    var employeeMap = {};
    var uniqueDates = {};

    salesorderSearchObj.run().each(function (result) {
      log.debug('result', result)
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

      if (note) {
        employeeMap[empKey].notes += (employeeMap[empKey].notes ? ' | ' : '') + note;
      }

      return true;
    });

    var sortedDates = Object.keys(uniqueDates).sort(function (a, b) {
      return new Date(a) - new Date(b);
    });

    var headerRow = {
      employee: 'Name',
      role: 'Role',
      shiftType: 'Shift<br/>Type',
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

    if (replaceLabor) {
      groupedFinalArray = replaceLaborText(groupedFinalArray);
    }

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
        ["formulatext: case when {custcol_bc_tm_line_id} = {custcol_bc_tm_source_transaction.line} then 1 else 0 end", 'is', '1']
      ],
      columns: [
        search.createColumn({
          name: 'custcol_invoicing_category',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'tranid',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'mainname',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'amount',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'MAX'
        }),
        search.createColumn({
          name: 'taxamount',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'MAX'
        }),
        search.createColumn({
          name: 'line',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'memo',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'formulatext',
          summary: 'MAX',
          formula: "CASE WHEN {custcol_bc_tm_source_transaction.appliedtotransaction} LIKE 'Purchase Order%' THEN TRIM(REPLACE({custcol_bc_tm_source_transaction.appliedtotransaction}, 'Purchase Order', '')) ELSE {custcol_bc_tm_source_transaction.tranid} END"
        }),
        search.createColumn({
          name: 'custcol_bc_tm_line_id',
          summary: 'GROUP'
        }),
        search.createColumn({
          name: 'amount',
          summary: 'MAX'
        }),
        search.createColumn({
          name: 'expensecategory',
          join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
          summary: 'GROUP'
        })
      ]
    });

    var tranFinalArray = {};

    transactionSearch.run().each(function (result) {
      
      var invoicingCategory = result.getText({
        name: 'custcol_invoicing_category',
        summary: 'GROUP'
      }) || 'Uncategorized';

      var docNumber = result.getValue({
        name: 'tranid',
        join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
        summary: 'GROUP'
      }) || '';

      var mainName = result.getText({
        name: 'mainname',
        join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
        summary: 'GROUP'
      }) || '';

      var expCat = result.getText({
        name: 'expensecategory',
        join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
        summary: 'GROUP'
      }) || '';

      var cost = parseFloat(result.getValue({
        name: 'amount',
        join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
        summary: 'MAX'
      }) || 0) || 0;

      var tax = parseFloat(result.getValue({
        name: 'taxamount',
        join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
        summary: 'MAX'
      }) || 0) || 0;

      var amount = parseFloat(result.getValue({
        name: 'amount',
        summary: 'MAX'
      }) || 0) || 0;

      var memo = result.getValue({
        name: 'memo',
        join: 'CUSTCOL_BC_TM_SOURCE_TRANSACTION',
        summary: 'GROUP'
      }) || '';

      var cleanedPO = result.getValue({
        name: 'formulatext',
        summary: 'MAX'
      }) || '';

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

      var grpAmt = parseFloat(group.totalAmount || 0).toFixed(2);
      var grpCost = parseFloat(group.totalCost || 0).toFixed(2);
      var grpTax = Math.abs(parseFloat(group.totalTax || 0)).toFixed(2);
      var grpTotal = (Math.abs(parseFloat(group.totalTax || 0)) + parseFloat(group.totalCost || 0)).toFixed(2);

      var totalRow = {
        documentNumber: 'TOTAL',
        mainName: '',
        expCat: '',
        amount: formatCurrency(grpAmt),
        cost: formatCurrency(grpCost),
        tax: formatCurrency(grpTax),
        total: formatCurrency(grpTotal),
        memo: '',
        cleanedPO: ''
      };

      groupedFinalArray[category] = group.rows.concat([totalRow]);
    }

    if (replaceLabor) {
      groupedFinalArray = replaceLaborText(groupedFinalArray);
    }

    return groupedFinalArray;
  }

  // --------------------------------------------------------------------------
  // EXCEL HTML
  // --------------------------------------------------------------------------
  function buildExcelHtml(invoiceData, timesheetDataByTran) {
    var html = ''
      + '<html xmlns:x="urn:schemas-microsoft-com:office:excel">'
      + '<head>'
      + '<meta charset="UTF-8">'
      + '<style>'
      + 'table { border-collapse: collapse; width: 100%; font-size: 10pt; table-layout: fixed; }'
      + 'th, td { border: 1px solid black; padding: 5px; word-wrap: break-word; }'
      + 'th { background-color: #3a4b87; color: white; font-weight: bold; }'
      + '.section-label { background-color: #e3e3e3; font-weight: bold; padding: 6px; border: 1px solid #000; }'
      + '.row-label { background-color: #3a4b87; color: white; font-weight: bold; }'
      + '.info-header { background-color: #00a3e0; color: white; font-weight: bold; }'
      + '.table-header { background-color: #3a4b87; color: white; font-weight: bold; }'
      + '.sub-header { background-color: #00a3e0; color: white; font-weight: bold; }'
      + '</style>'
      + '</head>'
      + '<body>';

    // Consolidated draft invoice
    html += '<div id="sheet1">';
    html += invoiceData.invoiceBlockHtml;
    html += '<br/><br/>';

    html += '<table style="width:100%; border-collapse:collapse;">'
      + '<tr><td colspan="20" style="background-color:#000; height:8px;"></td></tr>'
      + '</table>'
      + '<br/><br/>';

    // Timesheet blocks by tranid
    for (var i = 0; i < timesheetDataByTran.length; i++) {
      html += buildTimesheetHtmlBlock(timesheetDataByTran[i]);

      if (i !== timesheetDataByTran.length - 1) {
        html += '<br/><br/><br/><table style="width:100%; border-collapse:collapse;"><tr><td colspan="20" style="background-color:#000; height:8px;"></td></tr></table><br/><br/><br/>';
      }
    }

    html += '</div></body></html>';
    return html;
  }

  function buildTimesheetHtmlBlock(ts) {
    var x = ts.groupedData || {};
    var h = ts.headerInfo || {};
    var labelLabor = ts.replaceLabor ? 'Labour' : 'Labor';

    var html = '';

    html += ''
      + '<table style="width:100%; border-collapse: collapse; font-size:10pt;">'
      + '<tr>'
      + '<td colspan="4" rowspan="7" style="padding:10px;">'
      + '<img src="' + h.logoUrl + '" height="100" />'
      + '</td>'
      + '<td colspan="17" rowspan="7" style="font-size:26pt; font-weight:bold; text-align:center; vertical-align:middle;">Weekly Timesheet - ' + esc(h.docNumber) + '</td>'
      + '</tr>'
      + '</table>'
      + '<br/><br/><br/>';

    html += ''
      + '<table style="width:100%; border-collapse:collapse; font-size:9pt; table-layout:fixed; margin-top:10px;">'
      + '<tr>'
      + '<td colspan="2" style="width:49%; vertical-align:top;">'
      + '<table style="width:100%; border-collapse:collapse;">'
      + '<tr><td class="info-header" colspan="2">Client:</td><td style="border:1px solid #000;" colspan="4">' + esc(h.client) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">Customer Ref #:</td><td style="border:1px solid #000; " colspan="4">' + esc(h.customerRef) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">Week-Ending:</td><td style="border:1px solid #000; mso-number-format:\\@;" colspan="4">' + formatDateMMDDYYYY(h.weekEnding) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">C2O Project Manager:</td><td style="border:1px solid #000;" colspan="4">' + esc(h.projectManager) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">Description of Work:</td><td style="border:1px solid #000;" colspan="4">' + esc(h.description) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">Document #:</td><td style="border:1px solid #000;" colspan="4">' + esc(h.docNumber) + '</td></tr>'
      + '</table>'
      + '</td>'

      + '<td style="width:2%; border:0px;" colspan="4"></td>'

      + '<td colspan="2" style="width:49%; vertical-align:top;">'
      + '<table style="width:100%; border-collapse:collapse;">'
      + '<tr><td class="info-header" colspan="2">Project:</td><td colspan="4" style="border:1px solid #000;">' + esc(h.reportingProject) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">C2O Job:</td><td colspan="4" style="border:1px solid #000;">' + esc(h.projectName) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">Supervisor:</td><td colspan="4" style="border:1px solid #000;">' + esc(h.supervisor) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">Start Time: Monday – Friday</td><td style="border:1px solid #000;">' + esc(h.startTime) + '</td><td class="info-header" colspan="2">Finish Time:</td><td style="border:1px solid #000;">' + esc(h.endTime) + '</td></tr>'
      + '<tr><td class="info-header" colspan="2">Start Time: Weekend / Holiday</td><td style="border:1px solid #000;">' + esc(h.startTime) + '</td><td class="info-header" colspan="2">Finish Time:</td><td style="border:1px solid #000;">' + esc(h.endTime) + '</td></tr>'
      + '</table>'
      + '</td>'
      + '</tr>'
      + '</table>'
      + '<br/><br/><br/>';
    log.debug('LaborMap', x.Labor || x.Labour)
    if (x.Labor || x.Labour) {
      html += '<table>'
        + '<tr>'
        + '<th class="table-header" colspan="6">' + labelLabor + '</th>'
        + '<td colspan="15" align="center" style="border-top:1px solid #000; border-bottom:1px solid #000; border-left:1px solid #000; border-right:1px solid #000;">ALL HOURS SHOWN ARE HOURS WORKED</td>'
        + '</tr>'
        + '<tr>'
        + '<th class="table-header" colspan="2" rowspan="2">Name</th>'
        + '<th class="table-header" colspan="2" rowspan="2">Role</th>'
        + '<th class="table-header" rowspan="2">Time Type</th>'
        + '<th class="table-header" rowspan="2">Shift Type</th>';

      var labor = x.Labor || x.Labour;
      for (var i1 = 0; i1 < labor[0].days.length; i1++) {
        html += '<th class="table-header">' + getDayName(labor[0].days[i1].date) + '</th>';
      }

      html += '<th class="table-header" rowspan="2">Total Week</th>'
        + '<th class="table-header" rowspan="2">Rate</th>'
        + '<th class="table-header" rowspan="2">Claim Amount</th>'
        + '<th class="table-header" rowspan="2" colspan="' + (12 - labor[0].days.length) + '">Notes</th>'
        + '</tr>'
        + '<tr>';

      for (var i2 = 0; i2 < labor[0].days.length; i2++) {
        html += '<th class="table-header" style="mso-number-format:\\@;">' + formatDateMMDDYYYY(labor[0].days[i2].date) + '</th>';
      }

      html += '</tr>';

      for (var q = 1; q < labor.length - 1; q++) {
  html += '<tr>'
    + '<td colspan="2">' + labor[q].employee + '</td>'
    + '<td colspan="2">' + labor[q].role + '</td>'
    + '<td>' + labor[q].shiftType + '</td>'
    + '<td>' + labor[q].shift + '</td>';

  for (var w = 0; w < labor[q].days.length; w++) {
    html += '<td>' + labor[q].days[w].hours + '</td>';
  }

  html += '<td>' + labor[q].totalWeek + '</td>'
    + '<td>' + formatCurrency(labor[q].rate) + '</td>'
    + '<td>' + formatCurrency(labor[q].amt) + '</td>'
    + '<td colspan="' + (12 - labor[q].days.length) + '"></td>'
    + '</tr>';
}

// last element separate
if (labor.length > 1) {
  var last = labor[labor.length - 1];

  html += '<tr>'
    + '<td colspan="5" style="border-left: 0; border-bottom: 0;"></td>'
    + '<td class="table-header"><b>' + last.employee + '</b></td>';

  for (var w = 0; w < last.days.length; w++) {
    html += '<td class="table-header"><b>' + last.days[w].hours + '</b></td>';
  }

  html += '<td class="table-header"><b>' + last.totalWeek + '</b></td>'
    + '<td class="table-header"><b>' + formatCurrency(last.rate) + '</b></td>'
    + '<td class="table-header"><b>' + formatCurrency(last.amt) + '</b></td>'
    + '</tr>';
}
      html += '</table>';
    }

    if (x['Equipment / Vehicle Rental']) {
      var equp = x['Equipment / Vehicle Rental'];

      html += '<br/><br/><br/><table>'
        + '<tr><th colspan="8">Equipment / Vehicle Rental</th></tr>'
        + '<tr>'
        + '<th colspan="4" rowspan="2">Role</th>';

      for (var e1 = 0; e1 < equp[0].days.length; e1++) {
        html += '<th>' + getDayName(equp[0].days[e1].date) + '</th>';
      }

      html += '<th rowspan="2">Total Week</th>'
        + '<th rowspan="2" colspan="' + (14 - equp[0].days.length) + '">Notes</th>'
        + '</tr>'
        + '<tr>';

      for (var e2 = 0; e2 < equp[0].days.length; e2++) {
        html += '<th style="mso-number-format:\\@;">' + formatDateMMDDYYYY(equp[0].days[e2].date) + '</th>';
      }

      html += '</tr>';

      for (var r = 1; r < equp.length; r++) {
        html += '<tr><td colspan="4">' + equp[r].role + '</td>';
        for (var t = 0; t < equp[r].days.length; t++) {
          html += '<td>' + equp[r].days[t].hours + '</td>';
        }
        html += '<td>' + equp[r].totalWeek + '</td>'
          + '<td colspan="' + (14 - equp[r].days.length) + '"></td>'
          + '</tr>';
      }

      html += '</table>';
    }

    if (x.Materials) {
      html += '<br/><br/><br/><table>'
        + '<tr><th class="table-header" colspan="5">Materials</th></tr>'
        + '<tr>'
        + '<th class="table-header" colspan="2">Supplier Invoice #</th>'
        + '<th class="table-header" colspan="3">Supplier</th>'
        + '<th class="table-header" colspan="2">PO #</th>'
        + '<th class="table-header" colspan="8">Description</th>'
        + '<th class="table-header">Total Cost excl. Tax</th>'
        + '<th class="table-header">Cost + Mark up</th>'
        + '</tr>';

      for (var p = 0; p < x.Materials.length; p++) {
        var m = x.Materials[p];
        if (m.documentNumber === 'TOTAL') {
          html += '<tr>'
            + '<td colspan="13" style="border:0px solid #000;"></td>'
            + '<td colspan="2" align="right" style="background-color:#3a4b87; color:white; font-weight:bold;">Total</td>'
            + '<td align="right"></td>'
            + '<td align="right" style="font-weight:bold;" >' + m.amount + '</td>'
            + '</tr>';
        } else {
          html += '<tr>'
            + '<td colspan="2">' + m.documentNumber + '</td>'
            + '<td colspan="3">' + m.mainName + '</td>'
            + '<td colspan="2">' + m.cleanedPO + '</td>'
            + '<td colspan="8">' + m.memo + '</td>'
            + '<td align="right">' + formatCurrency(m.cost) + '</td>'
            + '<td align="right">' + formatCurrency(m.amount) + '</td>'
            + '</tr>';
        }
      }

      html += '</table>';
    }

    if (x.Expenses) {
      html += '<br/><br/><br/><table>'
        + '<tr><th class="table-header" colspan="5">Expenses</th></tr>'
        + '<tr>'
        + '<th class="table-header" colspan="5">Expense Category</th>'
        + '<th class="table-header" colspan="2">PO #</th>'
        + '<th class="table-header" colspan="8">Description</th>'
        + '<th class="table-header">Total Cost excl. Tax</th>'
        + '<th class="table-header">Cost + Mark up</th>'
        + '</tr>';

      for (var a = 0; a < x.Expenses.length; a++) {
        var e = x.Expenses[a];
        if (e.documentNumber === 'TOTAL') {
          html += '<tr>'
            + '<td colspan="5" style="border:0px solid #000; background-color:#3a4b87; color:white; font-weight:bold;">Total</td>'
            + '<td colspan="10" align="right"></td>'
            + '<td align="right" style="background-color:#3a4b87; color:white; font-weight:bold;">' + e.cost + '</td>'
            + '<td align="right" style="background-color:#3a4b87; color:white; font-weight:bold;">' + e.amount + '</td>'
            + '</tr>';
        } else {
          html += '<tr>'
            + '<td colspan="5">' + e.expCat + '</td>'
            + '<td colspan="2">' + e.cleanedPO + '</td>'
            + '<td colspan="8">' + e.memo + '</td>'
            + '<td align="right">' + formatCurrency(e.cost) + '</td>'
            + '<td align="right">' + formatCurrency(e.amount) + '</td>'
            + '</tr>';
        }
      }

      html += '</table><br/><br/><br/><br/>';
    }

    if (ts.legendArray && ts.legendArray.length > 0) {
      html += '<br/><br/><table style="width:100%; border-top:1px solid #ccc; border-collapse:collapse; font-size:9pt;"><tr><td colspan="15" style="padding-top:8px; padding-bottom:8px;"><strong>Time Type Legend:</strong>&nbsp;&nbsp;';

      for (var lg = 0; lg < ts.legendArray.length; lg++) {
        html += '<strong>' + ts.legendArray[lg].abbr + '</strong> – ' + ts.legendArray[lg].label;
        if (lg < ts.legendArray.length - 1) {
          html += '&nbsp;&nbsp;|&nbsp;&nbsp;';
        }
      }

      html += '</td></tr></table>';
    }

    return html;
  }

  // --------------------------------------------------------------------------
  // LEGEND + TEXT REPLACEMENT
  // --------------------------------------------------------------------------
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
        var row = rows[r];
        if (row.shiftType && row.shiftType !== '') {
          var cleanType = String(row.shiftType).replace(/<br\/>/g, '').trim();
          if (TIME_LEGEND_MAP[cleanType] && !seenTypes[cleanType]) {
            seenTypes[cleanType] = true;
            legendArray.push({
              abbr: cleanType,
              label: TIME_LEGEND_MAP[cleanType]
            });
          }
        }
      }
    }

    var LEGEND_ORDER = ['ST', 'OT', 'DT', 'PT', 'PTO', 'Per Diem', 'DR1', 'DR2', 'DR3'];
    legendArray.sort(function (a, b) {
      return LEGEND_ORDER.indexOf(a.abbr) - LEGEND_ORDER.indexOf(b.abbr);
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
    if (v === null || v === undefined) return '';
    return String(v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/\r?\n/g, '<br>');
  }

  function esc(v) {
    if (v === null || v === undefined) return '';
    return String(v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/\r?\n/g, '<br />');
  }

  function escPlain(v) {
    if (v === null || v === undefined) return '';
    return String(v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function money(n) {
    var x = parseFloat(n || 0);
    if (!isFinite(x)) x = 0;
    return '$' + x.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  function pct(n) {
    var x = parseFloat(n || 0);
    if (!isFinite(x)) x = 0;
    return x.toFixed(2) + '%';
  }

  function formatCurrency(amount) {
    if (amount == '' || amount == null) return '';
    if (String(amount).indexOf('$') != -1) return amount;
    return '$ ' + parseFloat(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  function getEmployeeList() {
    var returnObj = {};

    var employeeSearchObj = search.create({
      type: 'employee',
      filters: [],
      columns: [
        search.createColumn({ name: 'internalid' }),
        search.createColumn({
          name: 'formulatext',
          formula: "{firstname} || ' ' || {lastname}"
        })
      ]
    });

    employeeSearchObj.run().each(function (result) {
      returnObj[String(result.getValue({ name: 'internalid' }))] = result.getValue({ name: 'formulatext' }) || '';
      return true;
    });

    return returnObj;
  }

  function formatDateMMDDYYYY(dateStr) {
    if (!dateStr) return '';

    dateStr = String(dateStr);

    // Handle MM/DD/YYYY
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

    // Fallback
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

  return {
    onRequest: onRequest
  };
});