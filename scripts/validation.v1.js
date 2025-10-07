function runBigQueryValidation() {
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

  var bqProgram = programMap[programRaw] || Object.keys(programMap).find(k => k.toLowerCase() === programRaw.toLowerCase());
  if (!bqProgram) {
    SpreadsheetApp.getUi().alert("Program not found in mapping: '" + programRaw + "'. Check C5 value.");
    return;
  }

  // ---------- Display setup ----------
  sheet.getRange("E6").setValue(currMonth);   // Latest close month
  sheet.getRange("F6").setValue(programRaw);  // Program
  sheet.getRange("E19").setValue(programRaw); // Program display (optional)

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
  var rawSheet2 = ss.getSheetByName("raw data2"); // GPW only

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

  // Override GPW from raw data2
  var raw2Curr = getRawValues(rawSheet2, currMonth, programRaw);
  var raw2Prev = getRawValues(rawSheet2, prevMonth, programRaw);

  rawCurr.GPW = raw2Curr.GPW;
  rawPrev.GPW = raw2Prev.GPW;

  // ---------- Growth Table ----------
  sheet.getRange("F21").setValue(currMonth);
  sheet.getRange("G21").setValue(prevMonth);
  sheet.getRange("F27").setValue(currMonth);
  sheet.getRange("G27").setValue(prevMonth);
  sheet.getRange("I21").setValue(currMonth);
  sheet.getRange("J21").setValue(prevMonth);
  sheet.getRange("I27").setValue(currMonth);
  sheet.getRange("J27").setValue(prevMonth);


  // BigQuery Current
  sheet.getRange("F22").setValue(numVal(policiesCurr.GPW));
  sheet.getRange("F23").setValue(numVal(policiesCurr.GPE));
  sheet.getRange("F24").setValue(numVal(cashCurr.CollectedPremium));
  sheet.getRange("F25").setValue(numVal(policiesCurr.policy_count));
  sheet.getRange("F28").setValue(numVal(claimsCurr.loss_paid));
  sheet.getRange("F29").setValue(numVal(claimsCurr.alae_paid));
  sheet.getRange("F30").setValue(numVal(claimsCurr.claim_count));

  // BigQuery Previous
  sheet.getRange("G22").setValue(numVal(policiesPrev.GPW));
  sheet.getRange("G23").setValue(numVal(policiesPrev.GPE));
  sheet.getRange("G24").setValue(numVal(cashPrev.CollectedPremium));
  sheet.getRange("G25").setValue(numVal(policiesPrev.policy_count));
  sheet.getRange("G28").setValue(numVal(claimsPrev.loss_paid));
  sheet.getRange("G29").setValue(numVal(claimsPrev.alae_paid));
  sheet.getRange("G30").setValue(numVal(claimsPrev.claim_count));

  // Raw Current
  sheet.getRange("I22").setValue(numVal(rawCurr.GPW));
  sheet.getRange("I23").setValue(numVal(rawCurr.GPE));
  sheet.getRange("I24").setValue(numVal(rawCurr.Collected));
  sheet.getRange("I25").setValue(numVal(rawCurr.PolicyCount));
  sheet.getRange("I28").setValue(numVal(rawCurr.PaidLoss));
  sheet.getRange("I29").setValue(numVal(rawCurr.PaidExpense));
  sheet.getRange("I30").setValue(numVal(rawCurr.ClaimCount));

  // Raw Previous
  sheet.getRange("J22").setValue(numVal(rawPrev.GPW));
  sheet.getRange("J23").setValue(numVal(rawPrev.GPE));
  sheet.getRange("J24").setValue(numVal(rawPrev.Collected));
  sheet.getRange("J25").setValue(numVal(rawPrev.PolicyCount));
  sheet.getRange("J28").setValue(numVal(rawPrev.PaidLoss));
  sheet.getRange("J29").setValue(numVal(rawPrev.PaidExpense));
  sheet.getRange("J30").setValue(numVal(rawPrev.ClaimCount));

  // Growth %
  sheet.getRange("H22").setValue(growth(numVal(policiesCurr.GPW), numVal(policiesPrev.GPW)));
  sheet.getRange("H23").setValue(growth(numVal(policiesCurr.GPE), numVal(policiesPrev.GPE)));
  sheet.getRange("H24").setValue(growth(numVal(cashCurr.CollectedPremium), numVal(cashPrev.CollectedPremium)));
  sheet.getRange("H25").setValue(growth(numVal(policiesCurr.policy_count), numVal(policiesPrev.policy_count)));
  sheet.getRange("H28").setValue(growth(numVal(claimsCurr.loss_paid), numVal(claimsPrev.loss_paid)));
  sheet.getRange("H29").setValue(growth(numVal(claimsCurr.alae_paid), numVal(claimsPrev.alae_paid)));
  sheet.getRange("H30").setValue(growth(numVal(claimsCurr.claim_count), numVal(claimsPrev.claim_count)));

  sheet.getRange("K22").setValue(growth(rawCurr.GPW, rawPrev.GPW));
  sheet.getRange("K23").setValue(growth(rawCurr.GPE, rawPrev.GPE));
  sheet.getRange("K24").setValue(growth(rawCurr.Collected, rawPrev.Collected));
  sheet.getRange("K25").setValue(growth(rawCurr.PolicyCount, rawPrev.PolicyCount));
  sheet.getRange("K28").setValue(growth(rawCurr.PaidLoss, rawPrev.PaidLoss));
  sheet.getRange("K29").setValue(growth(rawCurr.PaidExpense, rawPrev.PaidExpense));
  sheet.getRange("K30").setValue(growth(rawCurr.ClaimCount, rawPrev.ClaimCount));

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

    // Convert very small deltas to 0
    var tolerance = 0.01; // Adjust as needed for acceptable precision
    if (Math.abs(delta) < tolerance) {
      delta = 0;
    }

    sheet.getRange("H" + metric.row).setValue(numVal(metric.bq));
    sheet.getRange("I" + metric.row).setValue(numVal(metric.raw));
    sheet.getRange("J" + metric.row).setValue(delta);

    var status = (delta === 0) ? "Match" : "Mismatch";
    sheet.getRange("K" + metric.row).setValue(status);
  });

}
