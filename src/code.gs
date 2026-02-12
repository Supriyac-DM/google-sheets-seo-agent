/***********************
 * AI-RankTracker-GSheet (PRINCETON ONLY - 10 KW BATCH)
 * - Targets: Princeton, New Jersey, United States
 * - Logic: Deep-Scans Rank 1-50 (5 pages)
 * - API: SerpApi (Integrated Key)
 ***********************/

function getConfig_() {
  const ss = SpreadsheetApp.getActive();
  const cfg = ss.getSheetByName('Config');
  if (!cfg) throw new Error('Missing "Config" sheet.');
  
  const getVal = (k) => {
    const cell = cfg.createTextFinder(k).matchEntireCell(true).findNext();
    return cell ? String(cfg.getRange(cell.getRow(), cell.getColumn() + 1).getValue()).trim() : null;
  };
  
  return {
    apiKey: 'XXXXXXXXXX',
    domain: normalizeDomain_(getVal('DOMAIN')),
    location: 'Princeton, New Jersey, United States', // Locked to Princeton
    hl: getVal('HL') || 'en',
    gl: getVal('GL') || 'us'
  };
}

function normalizeDomain_(d) {
  return d.replace(/^https?:\/\//, '').replace(/^www\./, '').split('/')[0].toLowerCase();
}

function runPrincetonRankTracker() {
  const ss = SpreadsheetApp.getActive();
  const keysSheet = ss.getSheetByName('Keywords');
  const cfg = getConfig_();
  const lastRow = keysSheet.getLastRow();
  if (lastRow < 2) return;

  const dataRange = keysSheet.getRange(2, 1, lastRow - 1, 2);
  const data = dataRange.getValues();
  const out = ss.getSheetByName('Ranks') || ss.insertSheet('Ranks');
  if (out.getLastRow() < 1) out.appendRow(['Date','Engine','Location','Keyword','Rank','Ranking URL','Title']);

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  // Process exactly 10 keywords per run as requested
  let processed = 0;
  for (let i = 0; i < data.length; i++) {
    if (processed >= 10) break; 
    
    let [kw, status] = data[i];
    if (status === 'COMPLETE' || !kw) continue;

    Logger.log(`Checking: "${kw}" in ${cfg.location}`);
    const hit = fetchDeepRank_(kw, cfg.location, cfg);
    
    out.appendRow([
      stamp, 'Google', cfg.location, kw, 
      hit ? hit.rank : '>50', 
      hit ? hit.link : '', 
      hit ? hit.title : ''
    ]);
    
    keysSheet.getRange(i + 2, 2).setValue('COMPLETE');
    processed++;
  }
}

/** * Loops through first 5 pages of Google (Positions 1-50)
 */
function fetchDeepRank_(keyword, location, cfg) {
  for (let page = 0; page < 5; page++) {
    const startOffset = page * 10;
    const url = `https://serpapi.com/search.json?q=${encodeURIComponent(keyword)}&location=${encodeURIComponent(location)}&hl=${cfg.hl}&gl=${cfg.gl}&start=${startOffset}&api_key=${cfg.apiKey}`;
    
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      
      if (response.getResponseCode() !== 200) break;

      const organic = json.organic_results || [];
      if (organic.length === 0) break; 

      for (const res of organic) {
        const foundDomain = normalizeDomain_(res.link || '');
        if (foundDomain === cfg.domain || foundDomain.endsWith('.' + cfg.domain)) {
          return {
            rank: res.position,
            link: res.link,
            title: res.title
          };
        }
      }
    } catch (e) {
      Logger.log(`Page ${page} Error: ${e.message}`);
    }
    Utilities.sleep(250); // Speed optimized
  }
  return null;
}
/**
 * CREATES A VISUAL DASHBOARD FOR PORTFOLIO SHOWCASE
 * Generates summaries of ranking performance.
 */
function createAgentDashboard() {
  const ss = SpreadsheetApp.getActive();
  let dash = ss.getSheetByName('Dashboard');
  if (!dash) {
    dash = ss.insertSheet('Dashboard');
  } else {
    dash.clear();
  }

  const rankSheet = ss.getSheetByName('Ranks');
  if (!rankSheet) return;
  const data = rankSheet.getDataRange().getValues();
  const rows = data.slice(1); // Remove header

  // 1. Calculate Metrics
  const totalChecks = rows.length;
  const top10 = rows.filter(r => parseInt(r[4]) <= 10).length;
  const top3 = rows.filter(r => parseInt(r[4]) <= 3).length;
  const successRate = totalChecks > 0 ? ((top10 / totalChecks) * 100).toFixed(1) : 0;

  // 2. Format Dashboard
  dash.getRange('A1:C1').setValues([['AI AGENT PERFORMANCE DASHBOARD', '', '']])
      .merge().setFontSize(16).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');

  dash.getRange('A3:B6').setValues([
    ['Total Keywords Tracked', totalChecks],
    ['Ranking in Top 10', top10],
    ['Ranking in Top 3', top3],
    ['Visibility Score (%)', successRate + '%']
  ]).setBorder(true, true, true, true, true, true);

  // 3. List "Quick Wins" (Ranked 11-20)
  const quickWins = rows.filter(r => parseInt(r[4]) > 10 && parseInt(r[4]) <= 25)
                        .map(r => [r[3], r[4], r[5]]); // Keyword, Rank, URL

  dash.getRange('A8').setValue('SEO QUICK WINS (POS 11-25)').setFontWeight('bold');
  if (quickWins.length > 0) {
    dash.getRange(9, 1, 1, 3).setValues([['Keyword', 'Current Rank', 'Target URL']]).setFontWeight('bold');
    dash.getRange(10, 1, quickWins.length, 3).setValues(quickWins);
  } else {
    dash.getRange(9, 1).setValue('No keywords currently in striking distance.');
  }
  
  dash.setColumnWidth(1, 250);
  dash.setColumnWidth(3, 400);
  SpreadsheetApp.getUi().alert('Dashboard Updated! Check the Dashboard tab.');
}

/**
 * ADDS A CUSTOM MENU TO THE SHEET
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ AI AGENT')
    .addItem('Run Princeton Batch (10)', 'runPrincetonRankTracker')
    .addSeparator()
    .addItem('Update Dashboard', 'createAgentDashboard')
    .addItem('Reset Status', 'resetAndRun')
    .addToUi();
}
