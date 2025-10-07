function runBigQueryValidationv2() {
  var projectId = 'spinnaker-dev-315722';
  var dataset = 'dw_develop_extracts';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // ---------- Helpers ----------
  function formatCloseMonth(val) {
    if (!val) return "";
    if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val)) {
      var y = val.getFullYear();
      var m = val.getMonth() + 1;
      return y.toString() + ('0' + m).slice(-2);
    }
    var s = String(val).trim();
    var digits = s.replace(/\D/g, '');
    return digits.length >= 6 ? digits.substr(0, 6) : digits;
  }

  function getPreviousMonth(currMonth) {
    var y = parseInt(currMonth.substring(0, 4), 10);
    var m = parseInt(currMonth.substring(4, 6), 10) - 1;
    if (m === 0) { m = 12; y -= 1; }
    return y.toString() + ('0' + m).slice(-2);
  }

  function numVal(v) {
    if (v === null || v === "") return 0;
    if (typeof v === 'number') return v;
    var s = String(v).trim().replace(/[, ]+/g, '');
    var f = parseFloat(s);
    return isNaN(f) ? 0 : f;
  }

  function growth(curr, prev) {
    if (!prev || prev === 0) return "";
    return (curr - prev) / prev;
  }

  // ---------- Inputs ----------
  var rawMonth = sheet.getRange("C4").getValue();
  var currMonth = formatCloseMonth(rawMonth);
  if (!currMonth || currMonth.length < 6) {
    SpreadsheetApp.getUi().alert("Invalid Close Month in C4. Use YYYYMM or a date cell.");
    return;
  }
  currMonth = currMonth.substr(0, 6);
  var prevMonth = getPreviousMonth(currMonth);

  var programRawInput = sheet.getRange("C5").getValue();
  if (!programRawInput) {
    SpreadsheetApp.getUi().alert("Please set Program in C5.");
    return;
  }
  var programRaw = String(programRawInput).trim();

  // ---------- Program mapping ----------
  var programMap = {
    "Ahoy": "ahoy", "Battleface": "battleface", "Amrisc": "amrisc",
    "Annex-flood": "annex-flood", "Arrowhead": "arrowhead", "Cabrillo": "cbr",
    "Coterie": "coterie", "Euclid": "euclid", "HDVI": "hdvi", "MileAuto": "mileauto",
    "Outdoorsy": "outdoorsy", "RVP": "rvp", "Simply Business": "sb",
    "sola": "sola", "boost-cyber": "boost-cyber", "rg": "rg", "hippo": "hippo"
  };

  var bqProgram = programMap[programRaw] ||
    Object.keys(programMap).find(k => k.toLowerCase() === programRaw.toLowerCase());
  if (!bqProgram) {
    SpreadsheetApp.getUi().alert("Program not found in mapping: '" + programRaw + "'. Check C5 value.");
    return;
  }

  // ---------- Display setup ----------
  //sheet.getRange("E5").setValue(currMonth);
  sheet.getRange("E4").setValue(programRaw);

  // ---------- BigQuery helper ----------
  function runQuery(sql) {
    var request = { query: sql, useLegacySql: false };
    var queryResults = BigQuery.Jobs.query(request, projectId);
    if (!queryResults || !queryResults.rows) return {};
    var out = {};
    var fields = queryResults.schema?.fields || [];
    queryResults.rows.forEach(function(r, i) {
      if (i === 0) {
        r.f.forEach(function(cell, idx) {
          var fname = fields[idx].name;
          out[fname] = cell.v;
        });
      }
    });
    return out;
  }

  // ---------- BigQuery Queries ----------
  var policiesCurr = runQuery(`
    SELECT COUNT(*) AS policy_count,
           SUM(gross_premium_written) AS GPW,
           SUM(gross_premium_earned) AS GPE
    FROM \`${projectId}.${dataset}.ext_policies\`
    WHERE CAST(close_month AS STRING)='${currMonth}'
      AND LOWER(CAST(program AS STRING))=LOWER('${bqProgram}')
  `);

  var policiesPrev = runQuery(`
    SELECT COUNT(*) AS policy_count,
           SUM(gross_premium_written) AS GPW,
           SUM(gross_premium_earned) AS GPE
    FROM \`${projectId}.${dataset}.ext_policies\`
    WHERE CAST(close_month AS STRING)='${prevMonth}'
      AND LOWER(CAST(program AS STRING))=LOWER('${bqProgram}')
  `);

  var cashCurr = runQuery(`
    SELECT SUM(raw_collected_cash) AS CollectedPremium
    FROM \`${projectId}.${dataset}.ext_cash\`
    WHERE CAST(close_month AS STRING)='${currMonth}'
      AND LOWER(CAST(program AS STRING))=LOWER('${bqProgram}')
  `);

  var cashPrev = runQuery(`
    SELECT SUM(raw_collected_cash) AS CollectedPremium
    FROM \`${projectId}.${dataset}.ext_cash\`
    WHERE CAST(close_month AS STRING)='${prevMonth}'
      AND LOWER(CAST(program AS STRING))=LOWER('${bqProgram}')
  `);

  var claimsCurr = runQuery(`
    SELECT COUNT(*) AS claim_count,
           SUM(indemnity_paid_itd) AS loss_paid,
           SUM(alae_paid_itd) AS alae_paid
    FROM \`${projectId}.${dataset}.ext_claims\`
    WHERE CAST(close_month AS STRING)='${currMonth}'
      AND LOWER(CAST(program AS STRING))=LOWER('${bqProgram}')
  `);

  var claimsPrev = runQuery(`
    SELECT COUNT(*) AS claim_count,
           SUM(indemnity_paid_itd) AS loss_paid,
           SUM(alae_paid_itd) AS alae_paid
    FROM \`${projectId}.${dataset}.ext_claims\`
    WHERE CAST(close_month AS STRING)='${prevMonth}'
      AND LOWER(CAST(program AS STRING))=LOWER('${bqProgram}')
  `);

  // ---------- Raw Data Helpers ----------
  var rawSheet = ss.getSheetByName("raw data");
  var rawSheet2 = ss.getSheetByName("raw data2");

  function getRawValues(sheetObj, month, program) {
    var values = sheetObj.getDataRange().getValues();
    var result = { GPW: 0, GPE: 0, Collected: 0, PolicyCount: 0, PaidLoss: 0, PaidExpense: 0, ClaimCount: 0 };
    for (var i = 1; i < values.length; i++) {
      var rowMonth = formatCloseMonth(values[i][0]);
      var rowProgram = String(values[i][1]).trim();
      if (rowMonth === month && rowProgram === program) {
        result.GPW += Number(values[i][2]) || 0;
        result.GPE += Number(values[i][3]) || 0;
        result.Collected += Number(values[i][4]) || 0;
        result.PaidLoss += Number(values[i][5]) || 0;
        result.PaidExpense += Number(values[i][6]) || 0;
        result.ClaimCount += Number(values[i][7]) || 0;
        result.PolicyCount += Number(values[i][8]) || 0;
      }
    }
    return result;
  }

  var rawCurr = getRawValues(rawSheet, currMonth, programRaw);
  var rawPrev = getRawValues(rawSheet, prevMonth, programRaw);

  var raw2Curr = getRawValues(rawSheet2, currMonth, programRaw);
  var raw2Prev = getRawValues(rawSheet2, prevMonth, programRaw);

  if (raw2Curr && !isNaN(Number(raw2Curr.GPW))) rawCurr.GPW = raw2Curr.GPW;
  if (raw2Prev && !isNaN(Number(raw2Prev.GPW))) rawPrev.GPW = raw2Prev.GPW;

  // ---------- Populate Raw & BigQuery values ----------
  // Raw Previous (F), Current (G)
  sheet.getRange("F7").setValue(numVal(rawPrev.GPW));
  sheet.getRange("G7").setValue(numVal(rawCurr.GPW));
  sheet.getRange("F8").setValue(numVal(rawPrev.GPE));
  sheet.getRange("G8").setValue(numVal(rawCurr.GPE));
  sheet.getRange("F9").setValue(numVal(rawPrev.Collected));
  sheet.getRange("G9").setValue(numVal(rawCurr.Collected));
  sheet.getRange("F10").setValue(numVal(rawPrev.PolicyCount));
  sheet.getRange("G10").setValue(numVal(rawCurr.PolicyCount));
  sheet.getRange("F13").setValue(numVal(rawPrev.PaidLoss));
  sheet.getRange("G13").setValue(numVal(rawCurr.PaidLoss));
  sheet.getRange("F14").setValue(numVal(rawPrev.PaidExpense));
  sheet.getRange("G14").setValue(numVal(rawCurr.PaidExpense));
  sheet.getRange("F15").setValue(numVal(rawPrev.ClaimCount));
  sheet.getRange("G15").setValue(numVal(rawCurr.ClaimCount));

  // BigQuery Previous (I), Current (J)
  sheet.getRange("I7").setValue(numVal(policiesPrev.GPW));
  sheet.getRange("J7").setValue(numVal(policiesCurr.GPW));
  sheet.getRange("I8").setValue(numVal(policiesPrev.GPE));
  sheet.getRange("J8").setValue(numVal(policiesCurr.GPE));
  sheet.getRange("I9").setValue(numVal(cashPrev.CollectedPremium));
  sheet.getRange("J9").setValue(numVal(cashCurr.CollectedPremium));
  sheet.getRange("I10").setValue(numVal(policiesPrev.policy_count));
  sheet.getRange("J10").setValue(numVal(policiesCurr.policy_count));
  sheet.getRange("I13").setValue(numVal(claimsPrev.loss_paid));
  sheet.getRange("J13").setValue(numVal(claimsCurr.loss_paid));
  sheet.getRange("I14").setValue(numVal(claimsPrev.alae_paid));
  sheet.getRange("J14").setValue(numVal(claimsCurr.alae_paid));
  sheet.getRange("I15").setValue(numVal(claimsPrev.claim_count));
  sheet.getRange("J15").setValue(numVal(claimsCurr.claim_count));

  // ---------- Growth Table ----------
  sheet.getRange("H7").setValue(growth(rawCurr.GPW, rawPrev.GPW));
  sheet.getRange("H8").setValue(growth(rawCurr.GPE, rawPrev.GPE));
  sheet.getRange("H9").setValue(growth(rawCurr.Collected, rawPrev.Collected));
  sheet.getRange("H10").setValue(growth(rawCurr.PolicyCount, rawPrev.PolicyCount));
  sheet.getRange("H13").setValue(growth(rawCurr.PaidLoss, rawPrev.PaidLoss));
  sheet.getRange("H14").setValue(growth(rawCurr.PaidExpense, rawPrev.PaidExpense));
  sheet.getRange("H15").setValue(growth(rawCurr.ClaimCount, rawPrev.ClaimCount));

  sheet.getRange("K7").setValue(growth(numVal(policiesCurr.GPW), numVal(policiesPrev.GPW)));
  sheet.getRange("K8").setValue(growth(numVal(policiesCurr.GPE), numVal(policiesPrev.GPE)));
  sheet.getRange("K9").setValue(growth(numVal(cashCurr.CollectedPremium), numVal(cashPrev.CollectedPremium)));
  sheet.getRange("K10").setValue(growth(numVal(policiesCurr.policy_count), numVal(policiesPrev.policy_count)));
  sheet.getRange("K13").setValue(growth(numVal(claimsCurr.loss_paid), numVal(claimsPrev.loss_paid)));
  sheet.getRange("K14").setValue(growth(numVal(claimsCurr.alae_paid), numVal(claimsPrev.alae_paid)));
  sheet.getRange("K15").setValue(growth(numVal(claimsCurr.claim_count), numVal(claimsPrev.claim_count)));

  // ---------- Delta Table ----------
  var deltaMetrics = [
    { name: "Gross Written Premium", row: 7, bq: policiesCurr.GPW, raw: rawCurr.GPW },
    { name: "Gross Earned Premium",  row: 8, bq: policiesCurr.GPE, raw: rawCurr.GPE },
    { name: "Collected Premium",     row: 9, bq: cashCurr.CollectedPremium, raw: rawCurr.Collected },
    { name: "Policy Count",          row: 10, bq: policiesCurr.policy_count, raw: rawCurr.PolicyCount },
    { name: "Paid Loss",             row: 13, bq: claimsCurr.loss_paid, raw: rawCurr.PaidLoss },
    { name: "Paid Expense",          row: 14, bq: claimsCurr.alae_paid, raw: rawCurr.PaidExpense },
    { name: "Claim Count",           row: 15, bq: claimsCurr.claim_count, raw: rawCurr.ClaimCount }
  ];

  deltaMetrics.forEach(function(metric) {
    var delta = numVal(metric.bq) - numVal(metric.raw);
    var tolerance = 0.01;
    if (Math.abs(delta) < tolerance) delta = 0;
    sheet.getRange("L" + metric.row).setValue(delta);
    var status = (delta === 0) ? "Match" : "Mismatch";
    sheet.getRange("M" + metric.row).setValue(status);
  });

// Apply growth conditional formatting after populating all values
  applyGrowthTextColorFormatting();
}
