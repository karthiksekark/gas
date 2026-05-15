// ============================================================
//  GAS Trigger — Google Apps Script Backend  v6
//  Deploy as Web App: Execute as Me | Access: Anyone
//  Secret key: GET ?key=<val>  |  POST body.key
// ============================================================

const SECRET_KEY     = 'your-secret-key-here'
const COLUMNS        = ['Ticket Number', 'Title', 'Status', 'Due Date', 'Comments']
const TICKET_IDX     = 0
const TITLE_IDX      = 1
const STATUS_IDX     = 2
const DUEDATE_IDX    = 3
const COMMENTS_IDX   = 4
const JIRA_COL_COUNT = 5   // A–E

// Date row styling
const DATE_ROW_BG    = '#fff9c4'   // yellow — cols A–E on date rows
const DATE_LABEL_COL = 3           // col C — "Total Tickets"
const DATE_COUNT_COL = 4           // col D — ticket count
const DATE_FONT_CLR  = '#b8860b'   // dark amber text on date rows

// Header row styling
const HDR_BG         = '#4f6ef7'
const HDR_FG         = '#ffffff'

const BLANK_VAL  = '—'
const DUP_SUFFIX = ' modified'
const JIRA_BASE_URL = 'https://yourorg.atlassian.net'

// ── Responses ──
function respond(d) {
  return ContentService.createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON)
}
function unauth()    { return respond({ success:false, error:'Unauthorized', code:401 }) }
function badReq(m)   { return respond({ success:false, error:m||'Bad Request', code:400 }) }
function srvErr(m)   { return respond({ success:false, error:m||'Server Error', code:500 }) }
function ok(d)       { return respond(Object.assign({ success:true }, d)) }

function validateKey(k) { return k === SECRET_KEY }

// Returns only non-hidden sheets, preserving their left-to-right order.
function getVisibleSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(function(s) {
    return !s.isSheetHidden()
  })
}

// Targets the first visible (non-hidden) sheet by position.
function getSheet() {
  var sheets = getVisibleSheets()
  if (!sheets.length) throw new Error('No visible sheets found in spreadsheet')
  return sheets[0]
}

// ── Date parsing → M/D/YYYY or null ──
function parseDate(val) {
  if (!val && val !== 0) return null
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null
    return (val.getMonth()+1)+'/'+val.getDate()+'/'+val.getFullYear()
  }
  var s = String(val).trim()
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (m && +m[1]>=1&&+m[1]<=12&&+m[2]>=1&&+m[2]<=31&&+m[3]>=2000)
    return +m[1]+'/'+m[2]+'/'+m[3]
  var y = s.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (y && +y[2]>=1&&+y[2]<=12&&+y[3]>=1&&+y[3]<=31&&+y[1]>=2000)
    return +y[2]+'/'+y[3]+'/'+y[1]
  return null
}

// ── Style helpers ──

// Date row: full A–E yellow bg, "Total Tickets" in C, count in D
function styleDateRow(sheet, row, count) {
  // Clear full row formatting first
  sheet.getRange(row, 1, 1, JIRA_COL_COUNT)
       .setBackground(DATE_ROW_BG)
       .setFontColor(DATE_FONT_CLR)
       .setFontWeight('bold')
       .setHorizontalAlignment('left')

  // Col A — date value already there, just style
  // Col B — empty, already styled above
  // Col C — "Total Tickets" label
  sheet.getRange(row, DATE_LABEL_COL).setValue('Total Tickets')
  // Col D — count
  sheet.getRange(row, DATE_COUNT_COL).setValue(count)
  // Col E — empty, already styled above
}

// Header row: blue bg, white bold text
function styleHeader(sheet, row) {
  sheet.getRange(row, 1, 1, JIRA_COL_COUNT)
       .setValues([COLUMNS])
       .setBackground(HDR_BG)
       .setFontColor(HDR_FG)
       .setFontWeight('bold')
       .setHorizontalAlignment('left')
}

// Clear data row formatting (no background, default text)
function clearRowStyle(sheet, row, numCols) {
  sheet.getRange(row, 1, 1, numCols || JIRA_COL_COUNT)
       .setBackground(null)
       .setFontColor(null)
       .setFontWeight('normal')
       .setHorizontalAlignment('left')
}

// ── Sheet scanning ──

function getAllDates(sheet) {
  var lr = sheet.getLastRow()
  if (!lr) return []
  var vals = sheet.getRange(1,1,lr,1).getValues()
  var seen={}, out=[]
  for (var i=0;i<vals.length;i++) {
    var p=parseDate(vals[i][0])
    if (p&&!seen[p]) { seen[p]=true; out.push(p) }
  }
  return out
}

function rowsForDate(colAVals, dateStr) {
  var out=[]
  for (var i=0;i<colAVals.length;i++)
    if (parseDate(colAVals[i][0])===dateStr) out.push(i+1)
  return out
}

// Block: rows from triggerRow+2 (after header) to next date row-1 or lastRow
function blockBounds(triggerRow, colAVals, lastRow) {
  var s=triggerRow+2, e=lastRow
  for (var r=s;r<=lastRow;r++)
    if (r-1<colAVals.length&&parseDate(colAVals[r-1][0])!==null){e=r-1;break}
  return {start:s, end:e}
}

// Read block → { map: ticket→{rowNum,jiraVals,fullRow}, lastReal }
function readBlock(sheet, s, e, triggerRow) {
  var map={}, lastReal=triggerRow+1
  if (s>e) return {map:map, lastReal:lastReal}
  var lc=Math.max(sheet.getLastColumn(), JIRA_COL_COUNT)
  var rows=sheet.getRange(s,1,e-s+1,lc).getValues()
  for (var i=0;i<rows.length;i++) {
    var tk=String(rows[i][TICKET_IDX]||'').trim()
    var ti=String(rows[i][TITLE_IDX] ||'').trim()
    if (!tk||ti===BLANK_VAL||ti===''||ti.toLowerCase()==='title') continue
    map[tk]={rowNum:s+i, jiraVals:rows[i].slice(0,JIRA_COL_COUNT), fullRow:rows[i]}
    lastReal=s+i
  }
  return {map:map, lastReal:lastReal}
}

function headerExists(sheet, triggerRow) {
  if (triggerRow>=sheet.getLastRow()) return false
  var r=sheet.getRange(triggerRow+1,1,1,JIRA_COL_COUNT).getValues()[0]
  for (var i=0;i<COLUMNS.length;i++)
    if (String(r[i]).toLowerCase().trim()!==COLUMNS[i].toLowerCase()) return false
  return true
}

function countRealInBlock(sheet, triggerRow) {
  var lr = sheet.getLastRow()
  // Block data starts at triggerRow+2 if header exists, else triggerRow+1
  var s = headerExists(sheet, triggerRow) ? triggerRow + 2 : triggerRow + 1
  if (s > lr) return 0
  var e = lr
  // Batch-read colA from s to find next date row boundary
  var cA = sheet.getRange(s, 1, lr - s + 1, 1).getValues()
  for (var i = 0; i < cA.length; i++) {
    if (parseDate(cA[i][0]) !== null) { e = s + i - 1; break }
  }
  if (e < s) return 0
  var v = sheet.getRange(s, TITLE_IDX + 1, e - s + 1, 1).getValues()
  var n = 0
  for (var i = 0; i < v.length; i++) {
    var t = String(v[i][0] || '').trim()
    if (t && t !== BLANK_VAL && t.toLowerCase() !== 'title') n++
  }
  return n
}

// Write HYPERLINK formula in col A, setValue for B–E (skip Comments — user managed)
function writeTicketFormula(sheet, row, ticketKey) {
  if (!ticketKey) return
  sheet.getRange(row, TICKET_IDX+1)
       .setFormula('=HYPERLINK("'+JIRA_BASE_URL+'/browse/'+ticketKey+'","'+ticketKey+'")')
}

function applyStatusColor(sheet, row, status) {
  var cell = sheet.getRange(row, STATUS_IDX + 1)
  var s = String(status || '').trim().toLowerCase()
  if (s === 'production') {
    cell.setBackground('#cfe2ff').setFontColor('#084298')
  } else if (s === 'done') {
    cell.setBackground('#d1e7dd').setFontColor('#0a3622')
  } else if (s === 'in progress') {
    cell.setBackground('#fff3cd').setFontColor('#664d03')
  } else {
    cell.setBackground(null).setFontColor(null)
  }
}

function writeJiraRow(sheet, row, vals) {
  writeTicketFormula(sheet, row, String(vals[TICKET_IDX]||'').trim())
  sheet.getRange(row, TITLE_IDX+1)  .setValue(String(vals[TITLE_IDX]  ||''))
  sheet.getRange(row, STATUS_IDX+1) .setValue(String(vals[STATUS_IDX] ||''))
  sheet.getRange(row, DUEDATE_IDX+1).setValue(String(vals[DUEDATE_IDX]||''))
  clearRowStyle(sheet, row)
  applyStatusColor(sheet, row, vals[STATUS_IDX])
}

function updateJiraFields(sheet, row, vals) {
  sheet.getRange(row, TITLE_IDX+1)  .setValue(String(vals[TITLE_IDX]  ||''))
  sheet.getRange(row, STATUS_IDX+1) .setValue(String(vals[STATUS_IDX] ||''))
  sheet.getRange(row, DUEDATE_IDX+1).setValue(String(vals[DUEDATE_IDX]||''))
  clearRowStyle(sheet, row)
  applyStatusColor(sheet, row, vals[STATUS_IDX])
}

function restoreUserCols(sheet, destRow, fullRow) {
  // Only restore Comments (col E) and beyond — never overwrite with empty
  var commentsVal = fullRow[COMMENTS_IDX]
  if (commentsVal !== undefined && commentsVal !== null && commentsVal !== '') {
    sheet.getRange(destRow, COMMENTS_IDX + 1).setValue(commentsVal)
  }
  // Restore any extra user cols (F+) if they exist and are non-empty
  var lc = Math.max(sheet.getLastColumn(), JIRA_COL_COUNT)
  if (lc > JIRA_COL_COUNT) {
    var extraVals = fullRow.slice(JIRA_COL_COUNT)
    for (var i = 0; i < extraVals.length; i++) {
      if (extraVals[i] !== undefined && extraVals[i] !== null && extraVals[i] !== '') {
        sheet.getRange(destRow, JIRA_COL_COUNT + 1 + i).setValue(extraVals[i])
      }
    }
  }
}

// ── Finalise a date row after all writes ──
// Sets styling, count, header or empty-row as needed
function finaliseDate(sheet, triggerRow) {
  var count = countRealInBlock(sheet, triggerRow)
  styleDateRow(sheet, triggerRow, count)

  if (count > 0) {
    // Remove stale separator row if present (empty row between date row and header)
    var nextVal = String(sheet.getRange(triggerRow + 1, 1).getValue()).trim()
    if (nextVal === '' && !headerExists(sheet, triggerRow)) {
      // Check if row after the empty row is the header
      var rowAfter = sheet.getLastRow() >= triggerRow + 2
        ? String(sheet.getRange(triggerRow + 2, 1).getValue()).trim()
        : ''
      if (rowAfter.toLowerCase() === 'ticket number') {
        sheet.deleteRow(triggerRow + 1)  // remove the empty separator
      }
    }
    // Ensure header
    if (!headerExists(sheet, triggerRow)) {
      sheet.insertRowsAfter(triggerRow, 1)
    }
    styleHeader(sheet, triggerRow + 1)
    // Clear data row backgrounds
    var lr = sheet.getLastRow()
    var s  = triggerRow + 2
    var e  = lr
    if (s <= lr) {
      var cA = sheet.getRange(s, 1, lr - s + 1, 1).getValues()
      for (var i = 0; i < cA.length; i++)
        if (parseDate(cA[i][0]) !== null) { e = s + i - 1; break }
      if (e >= s) {
        var lc = Math.max(sheet.getLastColumn(), JIRA_COL_COUNT)
        sheet.getRange(s, 1, e - s + 1, lc)
             .setBackground(null).setFontColor(null).setFontWeight('normal')
      }
      // Insert ONE empty separator after block — skip if already empty
      if (e >= s) {
        var sepRow = e + 1
        var lrNow  = sheet.getLastRow()
        if (sepRow <= lrNow) {
          var sepVal = String(sheet.getRange(sepRow, 1).getValue()).trim()
          var sepIsDate = parseDate(sheet.getRange(sepRow, 1).getValue()) !== null
          // Only insert if next row is NOT already empty
          if (sepVal !== '' || sepIsDate) {
            sheet.insertRowsAfter(e, 1)
            sheet.getRange(sepRow, 1, 1, Math.max(sheet.getLastColumn(), JIRA_COL_COUNT))
                 .setBackground(null).setFontColor(null).setFontWeight('normal')
          }
        }
      }
    }
  } else {
    // Remove header if present
    if (headerExists(sheet, triggerRow)) sheet.deleteRow(triggerRow + 1)
    // Ensure exactly one empty separator row
    var lr2    = sheet.getLastRow()
    var nextRow = triggerRow + 1
    if (nextRow <= lr2) {
      var nv = String(sheet.getRange(nextRow, 1).getValue()).trim()
      if (nv !== '') {
        // No empty row yet — insert one
        sheet.insertRowsAfter(triggerRow, 1)
        var lc2 = Math.max(sheet.getLastColumn(), JIRA_COL_COUNT)
        sheet.getRange(nextRow, 1, 1, lc2)
             .setBackground(null).setFontColor(null).setFontWeight('normal')
      }
    }
  }
}

// ── CRUD ──

function readAll() {
  try {
    var sheet=getSheet(), data=sheet.getDataRange().getValues()
    if (data.length<=1) return ok({data:[]})
    var rows=[]
    for (var i=1;i<data.length;i++) {
      var obj={}
      COLUMNS.forEach(function(c,j){obj[c]=String(data[i][j]||'')})
      rows.push(obj)
    }
    return ok({data:rows})
  } catch(e){return srvErr(e.message)}
}

function getDates() {
  try { return ok({dates:getAllDates(getSheet())}) }
  catch(e){return srvErr(e.message)}
}

function createRow(data) {
  if (!data||!data['Title']) return badReq('Title required')
  try {
    var sheet=getSheet()
    // Duplicate check
    var all=sheet.getDataRange().getValues()
    var norm=data['Title'].trim().toLowerCase(), dups=[]
    for (var i=0;i<all.length;i++) {
      var t=String(all[i][TITLE_IDX]||'').trim().toLowerCase()
      if (t===norm&&t!=='title'&&t!==BLANK_VAL&&t!=='') dups.push(i+1)
    }
    if (dups.length) {
      dups.forEach(function(r){
        var cur=String(sheet.getRange(r,STATUS_IDX+1).getValue()).trim()
        if (!cur.endsWith(DUP_SUFFIX)) sheet.getRange(r,STATUS_IDX+1).setValue(cur+DUP_SUFFIX)
      })
      return ok({action:'updated', message:'Duplicate status updated at '+dups.length+' row(s)'})
    }
    var lr=sheet.getLastRow()
    var colA=lr?sheet.getRange(1,1,lr,1).getValues():[]
    var allDates=getAllDates(sheet)
    if (!allDates.length) return respond({success:false,error:'No date cells in column A',code:403})
    var vals=[data['Ticket Number']||'',data['Title']||'',data['Status']||'active',parseDate(data['Due Date'])||data['Due Date']||'','']
    var inserted=[]
    // Bottom-up
    var triggers=[]
    allDates.forEach(function(d){rowsForDate(colA,d).forEach(function(r){triggers.push(r)})})
    triggers.sort(function(a,b){return b-a})
    triggers.forEach(function(tr){
      // Ensure header
      if (!headerExists(sheet,tr)){sheet.insertRowsAfter(tr,1);styleHeader(sheet,tr+1)}
      var lr2=sheet.getLastRow()
      var cA=sheet.getRange(1,1,lr2,1).getValues()
      var bb=blockBounds(tr,cA,lr2)
      var bd=readBlock(sheet,bb.start,bb.end,tr)
      sheet.insertRowsAfter(bd.lastReal,1)
      writeJiraRow(sheet,bd.lastReal+1,vals)
      inserted.push(bd.lastReal+1)
    })
    // Re-read triggers and finalise
    var lr3=sheet.getLastRow()
    var cA3=sheet.getRange(1,1,lr3,1).getValues()
    triggers.forEach(function(tr){
      // find fresh row
      for (var i=0;i<cA3.length;i++)
        if (parseDate(cA3[i][0])===allDates[0]||(function(){
          for(var d=0;d<allDates.length;d++) if(parseDate(cA3[i][0])===allDates[d]&&i+1===tr) return true; return false
        })()) break
      finaliseDate(sheet,tr)
    })
    return ok({action:'created',message:'Inserted at '+inserted.length+' location(s)',insertedAtRows:inserted})
  } catch(e){return srvErr(e.message)}
}

function updateRow(ticket, data) {
  if (!ticket) return badReq('Ticket Number required')
  try {
    var sheet=getSheet(), all=sheet.getDataRange().getValues()
    var row=-1
    for (var i=0;i<all.length;i++) if(String(all[i][TICKET_IDX]).trim()===String(ticket).trim()){row=i+1;break}
    if (row===-1) return respond({success:false,error:'Not found: '+ticket,code:404})
    var ex=sheet.getRange(row,1,1,JIRA_COL_COUNT).getValues()[0]
    var nv=[ex[TICKET_IDX],
      data['Title']   !==undefined?data['Title']   :ex[TITLE_IDX],
      data['Status']  !==undefined?data['Status']  :ex[STATUS_IDX],
      data['Due Date']!==undefined?(parseDate(data['Due Date'])||data['Due Date']):ex[DUEDATE_IDX],
      ex[COMMENTS_IDX]]
    updateJiraFields(sheet,row,nv)
    return ok({message:'Updated',ticket:ticket})
  } catch(e){return srvErr(e.message)}
}

function deleteRow(ticket) {
  if (!ticket) return badReq('Ticket Number required')
  try {
    var sheet=getSheet(), all=sheet.getDataRange().getValues()
    var row=-1
    for (var i=0;i<all.length;i++) if(String(all[i][TICKET_IDX]).trim()===String(ticket).trim()){row=i+1;break}
    if (row===-1) return respond({success:false,error:'Not found: '+ticket,code:404})
    sheet.deleteRow(row)
    return ok({message:'Deleted '+ticket})
  } catch(e){return srvErr(e.message)}
}

// ── syncJira — single global pass ──
function syncJira(issuesByDate) {
  if (!issuesByDate||typeof issuesByDate!=='object') return badReq('issuesByDate required')
  try {
    var sheet=getSheet()
    var allDates=getAllDates(sheet)

    // Step 1: build globalMap — ticket → {blockDate, triggerRow, rowNum, jiraVals, fullRow}
    var globalMap={}
    var snapLr=sheet.getLastRow()
    var snapCA=snapLr?sheet.getRange(1,1,snapLr,1).getValues():[]

    for (var di=0;di<allDates.length;di++) {
      var d=allDates[di]
      var trs=rowsForDate(snapCA,d)
      for (var ri=0;ri<trs.length;ri++) {
        var tr=trs[ri]
        var bb=blockBounds(tr,snapCA,snapLr)
        var bd=readBlock(sheet,bb.start,bb.end,tr)
        for (var tk in bd.map)
          globalMap[tk]={blockDate:d,triggerRow:tr,rowNum:bd.map[tk].rowNum,
                         jiraVals:bd.map[tk].jiraVals,fullRow:bd.map[tk].fullRow}
      }
    }

    // Step 2: classify
    var updateBatch=[], moveBatch=[], insertBatch={}

    for (var dateKey in issuesByDate) {
      var issues=issuesByDate[dateKey]
      if (!Array.isArray(issues)) continue
      for (var ii=0;ii<issues.length;ii++) {
        var iss=issues[ii]
        var ticket =String(iss['Ticket Number']||'').trim()
        var title  =String(iss['Title']  ||'').trim()
        var status =String(iss['Status'] ||'unknown').trim()
        var rawDue =String(iss['Due Date']||'').trim()
        var normDue=parseDate(rawDue)||rawDue
        if (!ticket) continue
        var ex=globalMap[ticket]
        if (!ex) {
          if (!insertBatch[dateKey]) insertBatch[dateKey]=[]
          insertBatch[dateKey].push([ticket,title,status,normDue,''])
          continue
        }
        if (ex.blockDate!==dateKey) {
          moveBatch.push({ticket:ticket,newVals:[ticket,title,status,normDue,''],
            sourceRowNum:ex.rowNum,sourceTrigger:ex.triggerRow,
            sourceDate:ex.blockDate,destDate:dateKey,fullRow:ex.fullRow})
          continue
        }
        var xv=ex.jiraVals
        var normEx=parseDate(xv[DUEDATE_IDX])||String(xv[DUEDATE_IDX]||'').trim()
        if (String(xv[TITLE_IDX]||'').trim()!==title||
            String(xv[STATUS_IDX]||'').trim()!==status||normEx!==normDue)
          updateBatch.push({rowNum:ex.rowNum,newVals:[ticket,title,status,normDue,xv[COMMENTS_IDX]||''],
                            triggerRow:ex.triggerRow})
      }
    }

    var affDates={}, totalU=0, totalM=0, totalI=0

    // 3a: updates
    for (var u=0;u<updateBatch.length;u++) {
      updateJiraFields(sheet,updateBatch[u].rowNum,updateBatch[u].newVals)
      affDates[parseDate(sheet.getRange(updateBatch[u].triggerRow,1).getValue())||'']=true
      totalU++
    }

    // 3b: move deletes bottom-up
    moveBatch.sort(function(a,b){return b.sourceRowNum-a.sourceRowNum})
    for (var mv=0;mv<moveBatch.length;mv++) {
      sheet.deleteRow(moveBatch[mv].sourceRowNum)
      affDates[moveBatch[mv].sourceDate]=true
    }

    // 3c: move inserts into destination
    if (moveBatch.length) {
      var byDest={}
      moveBatch.forEach(function(m){if(!byDest[m.destDate])byDest[m.destDate]=[];byDest[m.destDate].push(m)})
      for (var dd in byDest) {
        var lr=sheet.getLastRow(), cA=sheet.getRange(1,1,lr,1).getValues()
        var destTrs=rowsForDate(cA,dd)
        if (!destTrs.length){Logger.log('WARN: dest date '+dd+' not found');continue}
        destTrs.sort(function(a,b){return b-a})
        destTrs.forEach(function(dtr){
          if (!headerExists(sheet,dtr)){sheet.insertRowsAfter(dtr,1);styleHeader(sheet,dtr+1)}
          var lr2=sheet.getLastRow(),cA2=sheet.getRange(1,1,lr2,1).getValues()
          var dbb=blockBounds(dtr,cA2,lr2), dbd=readBlock(sheet,dbb.start,dbb.end,dtr)
          var items=byDest[dd]
          sheet.insertRowsAfter(dbd.lastReal,items.length)
          items.forEach(function(m,mi){
            writeJiraRow(sheet,dbd.lastReal+1+mi,m.newVals)
            restoreUserCols(sheet,dbd.lastReal+1+mi,m.fullRow)
          })
          affDates[dd]=true
          totalM+=items.length
        })
      }
    }

    // 3d: new inserts
    for (var insDt in insertBatch) {
      var insRows=insertBatch[insDt]
      var lr=sheet.getLastRow(), cA=sheet.getRange(1,1,lr,1).getValues()
      var itrs=rowsForDate(cA,insDt)
      if (!itrs.length){Logger.log('WARN: date '+insDt+' not in sheet');continue}
      itrs.sort(function(a,b){return b-a})
      itrs.forEach(function(itr){
        if (!headerExists(sheet,itr)){sheet.insertRowsAfter(itr,1);styleHeader(sheet,itr+1)}
        var lr2=sheet.getLastRow(),cA2=sheet.getRange(1,1,lr2,1).getValues()
        var ibb=blockBounds(itr,cA2,lr2), ibd=readBlock(sheet,ibb.start,ibb.end,itr)
        sheet.insertRowsAfter(ibd.lastReal,insRows.length)
        insRows.forEach(function(rv,ri){writeJiraRow(sheet,ibd.lastReal+1+ri,rv)})
        affDates[insDt]=true
        totalI+=insRows.length
      })
    }

    // Mark all dates from issuesByDate + all dates already in sheet as affected
    for (var ibd2 in issuesByDate) affDates[ibd2]=true
    getAllDates(sheet).forEach(function(d){ affDates[d]=true })

    // 3f: finalise every affected date — collect ALL positions first, then
    // process bottom-up so insertRowsAfter in finaliseDate never shifts
    // rows we haven't visited yet. done{} prevents same date twice.
    var finalLr=sheet.getLastRow()
    if (finalLr>0) {
      var finalCA=sheet.getRange(1,1,finalLr,1).getValues()
      // Build list of {date, row} for affected dates
      var toFinalise=[]
      var seen={}
      for (var fd=0;fd<finalCA.length;fd++) {
        var fdDate=parseDate(finalCA[fd][0])
        if (!fdDate||!affDates[fdDate]||seen[fdDate]) continue
        seen[fdDate]=true
        toFinalise.push({date:fdDate, row:fd+1})
      }
      // Sort bottom-up so separator inserts don't shift unvisited rows above
      toFinalise.sort(function(a,b){return b.row-a.row})
      toFinalise.forEach(function(item){ finaliseDate(sheet, item.row) })
    }

    var totalSkip=0
    for (var dk in issuesByDate)
      if (Array.isArray(issuesByDate[dk])) totalSkip+=issuesByDate[dk].length
    totalSkip=Math.max(0,totalSkip-totalU-totalM-totalI)

    return ok({message:'Sync complete',
      stats:{updated:totalU,moved:totalM,inserted:totalI,skipped:totalSkip}})
  } catch(e){return srvErr(e.message)}
}

// ── Snapshot helpers ──────────────────────────────────────────────────────
//
//  takeSnapshot  — copies Sheet1 (values + formatting) into a hidden
//                  _snapshot tab before the syncJira write begins.
//  revertSnapshot — restores Sheet1 from _snapshot, then deletes the tab.
//  deleteSnapshot — deletes the _snapshot tab after a successful sync.
//
//  Called by the extension's background worker:
//    takeSnapshot   → GET  ?action=takeSnapshot
//    revertSnapshot → POST body.action='revertSnapshot'
//    deleteSnapshot → POST body.action='deleteSnapshot'

var SNAPSHOT_SHEET = '_snapshot'

function takeSnapshot() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet()
    var sheet = getSheet()

    // Remove any stale snapshot from a previous aborted run
    var existing = ss.getSheetByName(SNAPSHOT_SHEET)
    if (existing) ss.deleteSheet(existing)

    var lr = sheet.getLastRow()
    var lc = sheet.getLastColumn()

    // Create a hidden snapshot sheet
    var snap = ss.insertSheet(SNAPSHOT_SHEET)
    snap.hideSheet()

    if (lr > 0 && lc > 0) {
      var src = sheet.getRange(1, 1, lr, lc)
      var dst = snap.getRange(1, 1, lr, lc)
      // PASTE_NORMAL copies values, formulas, and all formatting
      src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
    }

    return ok({ message: 'Snapshot taken', rows: lr, cols: lc })
  } catch(e) { return srvErr(e.message) }
}

function revertSnapshot() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet()
    var sheet = getSheet()
    var snap  = ss.getSheetByName(SNAPSHOT_SHEET)

    if (!snap) {
      // Cancel happened before snapshot was taken — nothing to revert
      return ok({ message: 'No snapshot found — no changes were made', noSnapshot: true })
    }

    var snapLr = snap.getLastRow()
    var snapLc = snap.getLastColumn()
    var currLr = sheet.getLastRow()

    // 1. Clear Sheet1 (values + formatting)
    sheet.clear()

    // 2. Delete rows that were added during the partial sync so the
    //    sheet dimensions match the snapshot exactly
    var sheetLrNow = sheet.getMaxRows()
    var targetRows = Math.max(snapLr, 1)
    if (sheetLrNow > targetRows) {
      sheet.deleteRows(targetRows + 1, sheetLrNow - targetRows)
    }

    // 3. Restore from snapshot (values + formatting + formulas)
    if (snapLr > 0 && snapLc > 0) {
      var src = snap.getRange(1, 1, snapLr, snapLc)
      var dst = sheet.getRange(1, 1, snapLr, snapLc)
      src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
    }

    // 4. Delete the snapshot tab
    ss.deleteSheet(snap)

    return ok({ message: 'Sheet reverted to pre-sync state', rows: snapLr })
  } catch(e) { return srvErr(e.message) }
}

function deleteSnapshot() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet()
    var snap = ss.getSheetByName(SNAPSHOT_SHEET)
    if (snap) ss.deleteSheet(snap)
    return ok({ message: 'Snapshot deleted' })
  } catch(e) { return srvErr(e.message) }
}

// ── Entry points ──
function doGet(e) {
  var key=e.parameter&&e.parameter.key?e.parameter.key:null
  if (!validateKey(key)) return unauth()
  switch((e.parameter&&e.parameter.action)||'read') {
    case 'read':         return readAll()
    case 'getDates':     return getDates()
    case 'takeSnapshot': return takeSnapshot()
    default:             return badReq('Unknown action')
  }
}

function doPost(e) {
  var body={}
  try{body=JSON.parse(e.postData.contents)}catch(_){return badReq('Invalid JSON')}
  if (!validateKey(body.key||null)) return unauth()
  switch(body.action) {
    case 'create':         return createRow(body.data)
    case 'update':         return updateRow(body.id, body.data||{})
    case 'delete':         return deleteRow(body.id)
    case 'syncJira':       return syncJira(body.issuesByDate)
    case 'revertSnapshot': return revertSnapshot()
    case 'deleteSnapshot': return deleteSnapshot()
    default:               return badReq('Unknown action')
  }
}
