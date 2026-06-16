
  const DT = ["Real","Float","Dint","Sint","Uint","Int","Smallint","Bigint","String","Word","Nvarchar","Datetime","Bool","Bit","Byte"];
  const plcMap = (()=>{const m=new Map();["real","float","double"].forEach(v=>m.set(v,"float"));["int","integer","dint","sint","uint","smallint","bigint","byte"].forEach(v=>m.set(v,"int"));["nvarchar","nvar","string","word","char","nchar","varchar"].forEach(v=>m.set(v,"nvarchar"));["bool","bit","boolean"].forEach(v=>m.set(v,"bit"));return m;})();
  function normDT(dt){const r=String(dt||"").trim().toLowerCase();return plcMap.get(r)||dt;}
  const UOM=[
    {u:"%",d:"percent"},{u:"A",d:"amps"},{u:"APM",d:"arcs per minute"},{u:"Bar",d:"metric pressure"},
    {u:"btu",d:"British Thermal Unit"},{u:"deg",d:"degrees"},{u:"deg C",d:"degrees Celsius"},
    {u:"deg F",d:"degrees Fahrenheit"},{u:"ft",d:"feet"},{u:"ft/min",d:"feet per minute"},
    {u:"ft2",d:"square feet"},{u:"ft3",d:"cubic feet"},{u:"ft-lb",d:"foot-pound"},
    {u:"g",d:"gram"},{u:"gal",d:"gallon"},{u:"GPH",d:"gallons per hour"},{u:"Hrs",d:"hours"},
    {u:"Hz",d:"Hertz"},{u:"in",d:"inch"},{u:"inWC",d:"inch water column"},
    {u:"inch/sec",d:"inches per second"},{u:"kg",d:"kilogram"},{u:"kg/hr",d:"kg per hour"},
    {u:"kg/m3",d:"kg per cubic meter"},{u:"kPa",d:"kilopascal"},{u:"kW",d:"kilowatt"},
    {u:"kWh",d:"kilowatt hour"},{u:"l",d:"liter"},{u:"lb",d:"pound"},
    {u:"lb/ft3",d:"pounds per cubic foot"},{u:"lbs/hr",d:"pounds per hour"},
    {u:"L/min",d:"liters per minute"},{u:"m",d:"meter"},{u:"m/sec",d:"meters per second"},
    {u:"m16",d:"M 1/16"},{u:"m16/hr",d:"M 1/16 per hour"},{u:"m3",d:"cubic meter"},
    {u:"mA",d:"milliamps"},{u:"mg",d:"milligram"},{u:"mils",d:"mils paint thickness"},
    {u:"min",d:"minute"},{u:"ml",d:"milliliter"},{u:"mm",d:"millimeter"},
    {u:"mV",d:"millivolts"},{u:"on/off",d:"on or off"},{u:"oz",d:"ounce"},
    {u:"psi",d:"pounds per square inch"},{u:"PSIG",d:"PSI gauge"},
    {u:"rpm",d:"revolutions per minute"},{u:"s",d:"seconds"},{u:"SPM",d:"sparks per minute"},
    {u:"t",d:"ton"},{u:"t/hr",d:"ton per hour"},{u:"V",d:"volts"},{u:"W",d:"watt"},
  ];
  // ── Header mapping for Excel paste ──────────────────────────────────────────
  function canonicalizeHeader(h){return String(h||"").toLowerCase().replace(/[^a-z0-9]+/g,"");}
  const headerAliases={
    tableName:["table","tablename","tbl","table_name","tblname"],
    tagName:["tag","tagname","name","signal"],
    plcPath:["plcpath","plc path","path","plctagname","plc tag name"],
    plcAddress:["plcipaddress","plc ip address","ip","ipaddress","plcip","plc ip"],
    dataType:["datatype","data_type","type","plctype","plc data type"],
    units:["unit","units","uom","measure","measurement"],
    transactionFrequency:["transactionfrequency","transaction frequency","frequency","freq","transfreq","interval","period","scan rate"],
    tagDescription:["description","desc","tagdescription","details","note","notes"],
    min:["min","minimum","lower","low","lsl","lowerlimit"],
    max:["max","maximum","upper","high","usl","upperlimit"],
    newTable:["newtable","new table","newtbl","createtable","create table"]
  };
  const headerRegexMap={
    tableName:/table\s*name|^tablename$/i,
    tagName:/tag\s*name|^tagname$/i,
    plcAddress:/(plc.*address)|^plc&address$|address|plc\s*ip|ip\s*address|^ip$/i,
    plcPath:/plc\s*path|^plcpath$/i,
    dataType:/data\s*type|datatype/i,
    units:/^(units?|uom)$/i,
    transactionFrequency:/transaction\s*frequency|trans.*freq|^frequency$/i,
    tagDescription:/tag\s*description|^(description|desc)$/i,
    min:/^min(imum)?$/i,
    max:/^max(imum)?$/i,
    newTable:/new\s*table\??|^newtable\??$|create\s*table/i
  };
  const defaultColumnMap={tableName:0,tagName:1,plcAddress:2,plcPath:3,dataType:4,units:5,transactionFrequency:6,tagDescription:7,min:8,max:9,newTable:10};

  function headerScore(cell,key){
    const raw=String(cell||"");const norm=canonicalizeHeader(raw);
    if(headerRegexMap[key]&&headerRegexMap[key].test(raw))return 100;
    if(headerAliases[key]?.some(a=>a===norm))return 90;
    if(headerAliases[key]?.some(a=>norm.includes(a)))return 75;
    const bestAlias=(headerAliases[key]||[]).reduce((best,a)=>{const d=lev(norm,a);return(d<best.d)?{d,a}:best;},{d:Infinity,a:""});
    if(bestAlias.d<=2)return 60-bestAlias.d;
    if(norm.includes(canonicalizeHeader(key)))return 40;
    return 0;
  }
  function buildColumnMap(headerRow){
    const takenIdx=new Set();const map={};const keys=Object.keys(headerRegexMap);
    headerRow.forEach((cell,idx)=>{const cellNorm=(cell||"").trim().toLowerCase();for(const key of keys){if(headerRegexMap[key].test(cellNorm)&&!takenIdx.has(idx)&&map[key]===undefined){map[key]=idx;takenIdx.add(idx);break;}}});
    keys.forEach(key=>{if(map[key]!==undefined)return;let best={idx:-1,score:-1,text:""};headerRow.forEach((cell,idx)=>{if(takenIdx.has(idx))return;const s=headerScore(cell,key);if(s>best.score)best={idx,score:s,text:cell};});if(best.score>=55){map[key]=best.idx;takenIdx.add(best.idx);}});
    return map;
  }
  function isYesOrTrue(v){if(!v)return false;const n=String(v).trim().toLowerCase();return n==="yes"||n==="true"||n==="1"||n==="y";}

function openExcelTemplate(){
  const a=document.createElement("a");
  a.href="TagRequestTemplate.xlsx";
  a.download="TagRequestTemplate.xlsx";
  a.click();
}

  // Fuzzy match pasted data to known DT and Units options
  function fuzzyMatchDT(val){
    if(!val)return val;
    const v=val.trim().toLowerCase();
    // Exact match first
    const exact=DT.find(d=>d.toLowerCase()===v);
    if(exact)return exact.toUpperCase();
    // Partial match
    const partial=DT.find(d=>d.toLowerCase().includes(v)||v.includes(d.toLowerCase()));
    if(partial)return partial.toUpperCase();
    return val; // keep original if no match
  }
  function fuzzyMatchUnit(val){
    if(!val)return val;
    const v=val.trim().toLowerCase();
    // Exact unit match
    const exactU=UOM.find(u=>u.u.toLowerCase()===v);
    if(exactU)return exactU.u+" ("+exactU.d+")";
    // Exact description match
    const exactD=UOM.find(u=>u.d.toLowerCase()===v);
    if(exactD)return exactD.u+" ("+exactD.d+")";
    // Partial match on unit code
    const partialU=UOM.find(u=>u.u.toLowerCase().includes(v)||v.includes(u.u.toLowerCase()));
    if(partialU)return partialU.u+" ("+partialU.d+")";
    // Partial match on description
    const partialD=UOM.find(u=>u.d.toLowerCase().includes(v)||v.includes(u.d.toLowerCase()));
    if(partialD)return partialD.u+" ("+partialD.d+")";
    return val; // keep original if no match
  }

  // Vertical format headers (in order)
  const VERTICAL_HEADERS = [
    "tablename","tagname","plcipaddress","plcpath","datatype",
    "units","transactionfrequency","tagdescription","min","max"
  ];

  function isVerticalFormat(raw) {
    const lines = raw.trim().split(/\r?\n/).filter(l => l.trim() !== "");
    if (lines.length < 2) return false;
    // If any line contains a tab, it's Excel/TSV format — not vertical
    if (lines.some(l => l.includes("\t"))) return false;
    // Vertical format: first line matches a known header exactly
    const first = lines[0].trim().toLowerCase().replace(/[^a-z0-9]/g, "");
    if (VERTICAL_HEADERS.some(h => h === first)) return true;
    // Headerless vertical: single column, count is a multiple of the field count
    // (one complete row = 10 fields). Require at least one full row.
    if (lines.length >= VERTICAL_HEADERS.length && lines.length % VERTICAL_HEADERS.length === 0) return true;
    return false;
  }

  function parseVerticalFormat(raw) {
    const lines = raw.trim().split(/\r?\n/).map(l => l.trim()).filter(l => l !== "");
    const numFields = VERTICAL_HEADERS.length;
    // Detect header block: check if first numFields lines all look like headers
    // by checking each normalized line starts with or contains a known header keyword
    const headerKeywords = ["table","tag","plc","data","unit","freq","desc","min","max","desired","transaction","address","path","type","value"];
    let headerCount = 0;
    for (let i = 0; i < lines.length && i < numFields; i++) {
      const norm = lines[i].toLowerCase().replace(/[^a-z0-9]/g, "");
      if (headerKeywords.some(k => norm.includes(k))) {
        headerCount++;
      } else break;
    }
    // Skip header block if all numFields lines matched
    const dataLines = headerCount >= numFields ? lines.slice(numFields) : lines;
    const rows = [];
    for (let i = 0; i + numFields <= dataLines.length; i += numFields) {
      const chunk = dataLines.slice(i, i + numFields);
      rows.push({
        tableName:            chunk[0] || "",
        tagName:              chunk[1] || "",
        plcAddress:           chunk[2] || "",
        plcPath:              chunk[3] || "",
        dataType:             chunk[4] || "",
        units:                chunk[5] || "",
        transactionFrequency: chunk[6] || "",
        tagDescription:       chunk[7] || "",
        min:                  chunk[8] || "",
        max:                  chunk[9] || "",
        newTable:             false
      });
    }
    return rows;
  }

  function processExcelData(){
    const pasteArea=document.getElementById("pasteArea");
    const raw=pasteArea.value||"";
    if(!raw.trim()){alert("Paste some data first.");return;}
    if(isVerticalFormat(raw)){
      const rows=parseVerticalFormat(raw);
      let added=0;
      rows.forEach(r=>{
        if(!r.tableName&&!r.tagName&&!r.plcAddress&&!r.tagDescription)return;
        addRow({...r, dataType:fuzzyMatchDT(r.dataType), units:fuzzyMatchUnit(r.units)});added++;
      });
      pasteArea.value="";
      chkDups();
      alert(added+" row"+(added===1?"":"s")+" imported.");
    } else {
      Papa.parse(raw,{delimiter:"\t",skipEmptyLines:true,complete:({data})=>{
        if(!data||!data.length)return;
        const columnMap=buildColumnMap(data[0]);
        // Determine if first row is a header or data
        // A header row has recognizable column name keywords
        // A data row typically has an IP address, numeric values, or known data types in expected positions
        const firstRow=data[0]||[];
        const firstRowNorm=firstRow.map(c=>String(c||"").trim().toLowerCase().replace(/[^a-z0-9]/g,""));
        const headerKeywords=["table","tag","plc","data","unit","freq","desc","min","max","type","address","path","name","signal","measure"];
        const headerMatches=firstRowNorm.filter(c=>c.length>1&&headerKeywords.some(k=>c.includes(k))).length;
        // Also check if first row looks like data: IP pattern, known DT value, or numeric min/max
        const ipPattern=/^\d{1,3}\.\d{1,3}/;
        const hasIPInRow=firstRow.some(c=>ipPattern.test(String(c||"").trim()));
        const knownDTs=["real","float","dint","sint","uint","int","bool","bit","string","nvarchar","datetime","word","byte","smallint","bigint"];
        const hasDTInRow=firstRow.some(c=>knownDTs.includes(String(c||"").trim().toLowerCase()));
        const hasHeader=(headerMatches>=2)&&!hasIPInRow&&!hasDTInRow;
        const rows=hasHeader?data.slice(1):data;
        const map=hasHeader?columnMap:defaultColumnMap;
        let added=0;
        rows.forEach(r=>{
          const meaningful=[map.tableName,map.tagName,map.plcAddress,map.plcPath,map.dataType,map.units,map.transactionFrequency,map.tagDescription];
          if(meaningful.every(idx=>idx===undefined||!r[idx]||String(r[idx]).trim()===""))return;
          addRow({
            tableName:r[map.tableName]||"",tagName:r[map.tagName]||"",
            plcAddress:r[map.plcAddress]||"",plcPath:r[map.plcPath]||"",
            dataType:fuzzyMatchDT(r[map.dataType]||""),units:fuzzyMatchUnit(r[map.units]||""),
            transactionFrequency:r[map.transactionFrequency]||"",
            tagDescription:r[map.tagDescription]||"",
            min:r[map.min]||"",max:r[map.max]||"",
            newTable:isYesOrTrue(r[map.newTable])
          });added++;
        });
        pasteArea.value="";
        chkDups();
        alert(added+" row"+(added===1?"":"s")+" imported.");
      }});
    }
  }
  // ── End Excel paste ──────────────────────────────────────────────────────────

  function can(h){return String(h||"").toLowerCase().replace(/[^a-z0-9]+/g,"");}
  function lev(a,b){a=can(a);b=can(b);const m=a.length,n=b.length;if(!m)return n;if(!n)return m;const dp=Array.from({length:m+1},(_,i)=>new Array(n+1).fill(0));for(let i=0;i<=m;i++)dp[i][0]=i;for(let j=0;j<=n;j++)dp[0][j]=j;for(let i=1;i<=m;i++)for(let j=1;j<=n;j++)dp[i][j]=Math.min(dp[i-1][j]+1,dp[i][j-1]+1,dp[i-1][j-1]+(a[i-1]===b[j-1]?0:1));return dp[m][n];}

  function mkUnitsSearch(selectedVal){
    const items=UOM.map(u=>u.u+" ("+u.d+")");
    // Try to find display value for selected
    let displayVal="";
    if(selectedVal){
      const norm=selectedVal.trim().toLowerCase();
      const m=UOM.find(u=>u.u.toLowerCase()===norm)||UOM.find(u=>{const part=norm.match(/^([^(]+)/);return part&&u.u.toLowerCase()===part[1].trim();});
      if(m)displayVal=m.u+" ("+m.d+")";
      else displayVal=selectedVal;
    }
    const wrap=mkSearchDrop(items, displayVal, "— units —", function(inp, val, w){
      upJust();chkDups();
    });
    wrap._input.className="ci units-value";
    return wrap;
  }

  function mkSearchDrop(items, selectedVal, placeholder, onSelect, extraCheck){
    const wrap=document.createElement("div");wrap.style.position="relative";
    const inp=document.createElement("input");
    inp.type="text";inp.className="ci";
    inp.value=selectedVal||"";inp.placeholder=placeholder;
    inp.setAttribute("autocomplete","off");
    const drop=document.createElement("div");
    drop.style.cssText="position:absolute;top:100%;left:0;right:0;background:white;border:1.5px solid var(--wf);border-radius:6px;z-index:9999;max-height:180px;overflow-y:auto;display:none;box-shadow:0 4px 12px rgba(0,0,0,0.15);margin-top:2px;";
    function renderDrop(filter){
      drop.innerHTML="";
      const f=(filter||"").toLowerCase();
      const matches=items.filter(i=>i.toLowerCase().includes(f));
      if(!matches.length){drop.style.display="none";return;}
      matches.forEach(t=>{
        const item=document.createElement("div");
        item.textContent=t;
        item.style.cssText="padding:7px 10px;font-size:12px;cursor:pointer;color:var(--g900);border-bottom:1px solid var(--g100);";
        item.addEventListener("mousedown",function(e){
          e.preventDefault();inp.value=t;drop.style.display="none";
          if(onSelect)onSelect(inp,t,wrap);
        });
        item.addEventListener("mouseenter",function(){this.style.background="var(--wf-light)";});
        item.addEventListener("mouseleave",function(){this.style.background="";});
        drop.appendChild(item);
      });
      drop.style.display="block";
      const rect=inp.getBoundingClientRect();
      const spaceBelow=window.innerHeight-rect.bottom;
      const spaceAbove=rect.top;
      const dropH=Math.min(drop.scrollHeight,180);
      if(spaceBelow<dropH&&spaceAbove>spaceBelow){
        drop.style.top="auto";drop.style.bottom="100%";
        drop.style.marginTop="0";drop.style.marginBottom="2px";
        drop.style.boxShadow="0 -4px 12px rgba(0,0,0,0.15)";
      } else {
        drop.style.top="100%";drop.style.bottom="auto";
        drop.style.marginTop="2px";drop.style.marginBottom="0";
        drop.style.boxShadow="0 4px 12px rgba(0,0,0,0.15)";
      }
    }
    inp.addEventListener("focus",function(){renderDrop(this.value);});
    inp.addEventListener("input",function(){renderDrop(this.value);if(onSelect)onSelect(inp,this.value,wrap);});
    inp.addEventListener("blur",function(){setTimeout(()=>{drop.style.display="none";},200);});
    inp.addEventListener("keydown",function(e){
      const its=[...drop.querySelectorAll("div")];
      const ai=its.findIndex(i=>i.classList.contains("sd-active"));
      if(e.key==="ArrowDown"){e.preventDefault();if(ai>=0)its[ai].classList.remove("sd-active");const nx=its[ai+1]||its[0];if(nx){nx.classList.add("sd-active");nx.style.background="var(--wf-light)";}}
      else if(e.key==="ArrowUp"){e.preventDefault();if(ai>=0)its[ai].classList.remove("sd-active");const pv=its[ai-1]||its[its.length-1];if(pv){pv.classList.add("sd-active");pv.style.background="var(--wf-light)";}}
      else if(e.key==="Enter"){const ac=drop.querySelector(".sd-active");if(ac){e.preventDefault();inp.value=ac.textContent;drop.style.display="none";if(onSelect)onSelect(inp,ac.textContent,wrap);}}
      else if(e.key==="Escape"){drop.style.display="none";}
    });
    wrap.appendChild(inp);wrap.appendChild(drop);
    wrap._input=inp;
    return wrap;
  }

  function mkDTSelect(sel){
    const items=DT.map(d=>d);
    const wrap=mkSearchDrop(items, sel||"", "— type —", function(inp, val, w){
      // BIT warning
      w.querySelectorAll(".bit-warn").forEach(m=>m.remove());
      if(val.toUpperCase()==="BIT"){
        const wn=document.createElement("div");wn.className="msg warn bit-warn";
        wn.textContent="BIT not permitted — use INT instead";w.appendChild(wn);
      }
      chkDups();
    });
    wrap._input.className="ci dataType-input";
    return wrap;
  }

  function upEmpty(){const n=document.querySelectorAll("#tbody tr").length;document.getElementById("emptyState").style.display=n===0?"block":"none";document.getElementById("rowCounter").textContent=n+" row"+(n===1?"":"s");}
  function updateUI(){
    const ok=document.getElementById("millSelect").value!=="";
    ["addBtn","clrBtn","chkAllBtn","unchkBtn","copyBtn","subBtn"].forEach(id=>{const el=document.getElementById(id);if(el)el.disabled=!ok;});
    document.getElementById("tabManual").disabled=!ok;
    document.getElementById("tabPaste").disabled=!ok;
    document.getElementById("pasteArea").disabled=!ok;
    document.getElementById("processBtn").disabled=!ok;
    // Refresh table name fields in existing rows when mill changes
    document.querySelectorAll("#tbody tr").forEach(row=>{
      const tdCell=row.querySelector("td.col-tbl");
      if(!tdCell)return;
      const currentVal=getTableName(row);
      // Replace the cell content with a fresh field
      tdCell.innerHTML="";
      const newField=mkTableNameField(currentVal);
      tdCell.appendChild(newField);
    });
  }
  function switchTab(mode){
    const isManual=mode==="manual";
    document.getElementById("tabManual").classList.toggle("active",isManual);
    document.getElementById("tabPaste").classList.toggle("active",!isManual);
    document.getElementById("pastePanel").classList.toggle("visible",!isManual);
    // Show/hide Add row button based on mode
    document.getElementById("addBtn").style.display=isManual?"":"none";
  }

  // Table data injected by updateTables.js — do not edit this line manually
  let MILL_TABLES = {"ALD":["_RecipeUploadTest","_TagManager_Test_Audit","_tblLog","_to_delete_OpacityCalibrationTable_Temporary_Troubleshoot","BaghouseDataTable","Barcode_Error_Log","Barcode_NBInOutEntry_Insert_Audit","BlendingTable","BlendingTotalizerContinuousTable","BlendingTotalizerTable","BookSaw_CycleTime","CrossCutDataTable","CurrentProductTable","CycleTime_MS1","CycleTime_MS2","DashboardTable","DowntimeMachineIdTable","DowntimeMachineTable","Dryer1DataTable","Dryer2DataTable","EndOfShift","Energy1DataTable","Energy2DataTable","EnergyDataTable","EnvironmentalCalibTable","EnvironmentalTable","FinishingDataTable","FormingDataTable","FormingParametersTable","FormScanTable","Furnace1GrateCam","Furnace2GrateCam","GradeStampTable","MasterStackerDataTable","MasterWeightTable","OpacityCalibrationTable","OpacityCalibrationTable_20240718","OpacityCalibrationTypeTable","OpacityMeter","PanelPointThicknessTable","PanelThicknessTable","PressEfficiencyTable","PressLoadTable","PressParametersCyclicTable","PressParametersTable","PressScanTable","PressTestDataTable","RipCutDataTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RobotCyclicDataTable","RobotDataTable","RoutingDataTable","SK_Trend","SSMAbortTable","Strander1DataByEvent","Strander1KnifeChangeTable","Strander2DataByEvent","Strander2KnifeChangeTable","TestTable","WoodRoomDataTable"],"BEM":["AlertTriggerTable","Barcode_Error_Log","Barcode_PLCWeights_Insert_Audit","Blender_Data_Table","BlenderDowntimeTable","BlenderID","BlendingDataTable","BlendingTotalizerTable","Burner_Data_Table","CaulScreenTable","CoreDryerDowntimeTable","CoreDryerEfficiencyTable","CoreFaceAlarms_TO_DELETE","CoreLambTable","CoreRotaryDryer","Crew Table","CurrentProductTable","CycleTimeLoaderArm","CycleTimeLoaderArmLookup","CycleTimeLoaderCage","CycleTimeLoaderCageLookup","CycleTimeUnloaderCage","CycleTimeUnloaderCageLookup","DebarkerEfficiencyTable","DebarkerEfficiencyTable_20220315","DebarkerTable","DowntimeTable","Dryer3DowntimeTable","Dryer3EfficiencyTable","Dryer3Table","DryStorageTable","dtproperties","Efficiency","EnvironmentalTable","EnvironmentalTable_20230411","EnvironmentalTable_Config","Equipment","ERS_EnvironmentalScheduledEventsTable","ESPField1DownTimeTable","ESPField2DownTimeTable","Face_Lamb_Table","FaceDryerDowntimeTable","FaceDryerEfficiencyTable","FaceRotaryDryer","FinishingDataTable","FinishingDownGradeTable","FinishingDownTimeTable","FinishingTable","FlatLineDryerTable","FormingDataTable","FormingDataTableNew","FormingLineTable","FormingParametersTable","Konus1Table","Konus2Table","LogPondTable","MasterWtTable","MatHeightTCLTable","MatHeightTSLTable","PowderResinTable","Press Downtime Codes Table","PressAlarms_TO_BE_DELETED","PressDataTable","PressManTable","PressScanTable","PressTemperature","ProductTable","ProductTargetTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RTO15MinAverage","Rto1DownTimeTable","RTO1Table","Rto2DownTimeTable","RTO2Table","Rto3DowntimeTable","RTO3HrAverage","RTO3Table","Sander1SecondTable","Scott_Trend","ScreenCycleTable","ShiftEndTable","SideLiftDowntimeTable","StranderDowntimeTable","tblChartProperties_TO_DELETE","tblFormingLineMCRange_TO_DELETE","tblProcessVariableLimits_TO_DELETE","Trigger_Temp","UnitTable","Waferizer1Table","Waferizer2Table","Waferizer3Table","WellonsTable","WoodRoom","WoodRoom1Second"],"BWK":["_FinishParametersSP_Audit","AllEvent","Barcode_Error_Log","Barcode_NBInOutEntry_Insert_Audit","Barcode_PLCWIPMasterPanel_Insert_Audit","BarcodeTable","BlenderDowntimeTable","BlenderIDTable","BlenderTotalizerTable","BlendingDataTable","BlendingTotal","BlendingTotalCont","CurrentProductTable","Diagnostic","DowntimeTable","DryerTable","EFBTable","EnergyTable","EnvironmentTable","FinesTable","FinishingDowntime","FinishingTable","FlakerDowntime","FormingDataTable","FormingScanTable","FTAEInstance","HeatEnergyDataTable","LineStopTrapsTable","MasterWtTable","Messages","PressClosingDataTable","PressEfficiencyTable","PressLoadTable","PressProgramTable","PressScanTable","ReportCardTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankAndMeterReading_20250227","RMW_TankLevel","RMW_TotalizedMeterReading","RobotCyclic","RobotDataTable","SmartsheetRobotRowID","SSMAbortStack","SSMAbortTable","SteamRoomTable","Strander1Table","Strander2Table","Strander3Table","StranderDowntimeTable","StranderTable","SummingScale","sysdiagrams","VibSensorDataTable","WaxApplication","WaxStorageTable","WoodRoomTable"],"CHA":["_TagManager_Test_Audit","_TroubleShootTable","Barcode_Error_Log","Barcode_PLCWeights_Insert_Audit","BlenderTotalizerContinuousTable","BlenderTotalizerTable","BlendingDataTable","BlendingTable","BundleWeightDataTable","CurrentProductTable","DowntimeMachineIdTable","DowntimeMachineTable","Dryer1Table","DryerDataTable","EnergyDataTableF3","EnergyDataTableF4","Environnement","EnviroShiftEndTable","F1_LastBatchTable","F2_LastBatchTable","F3_LastBatchTable","FinishingCyclicTable","FlakingTable","FormingDataTable","Furnace3GrateCam","Furnace4GrateCam","GreenEndTable","Imal_RejectListTable","Imal_RejectTable","LoaderCageEventTable","MasterPanelWeightTable","MESShiftEndTable","MpcDataTable","MVMotorAmpsTempTable","PanelPointThicknessTable","PanelThicknessTable","PressEfficiencyTable","PressLoadTable","PressScanTable","ProductMapping","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RobotCyclicDataTable","RobotDataTable","RSSQLTriggersTable","ScreeningDataTable","ShiftEndDataTable","Strander1ByEvent","Strander2ByEvent","Strander3ByEvent","Strander3DataByEvent","Strander3KnifeChangeTable","Strander4DataByEvent","Strander4KnifeChangeTable","TankDesc","TankFarmLevels","WESPTable","Woodroom"],"COR":["_TagManager_Test_Audit","AirCompressorManagement","Barcode_Error_Log","Barcode_InsertUnit_Audit_____TO DELETE","Barcode_NBInOutEntry_Insert_Audit","Barcode_PLCWeights_Insert_Audit","Barcode_PLCWIPMasterPanel_Insert_Audit","BarCodeProductsTable","Blender2SnapShotTable","BlenderDowntimeTable","BlenderID","BlenderSnapShotTable","BlenderTotalsTable1","BlenderTotalsTable1Cont","CrossCarriageTimers","CurrentProductTable","CurrentYearTable","DieffensorDataTable","DowntimeCodesTable","DowntimeTable","Dryer1SnapShotTable","Dryer2SnapShotTable","DryerAreaTable","DryerDowntimeCodesTable","DryerDowntimeTable","DryerIdTable","DryerOutletAirlocksDataTable","DryerSnapShotTable","DryersSparkDetectTable","DryersSparkRateCounterTable","dtproperties","DustCollSnapShotTable","EndOfShift","EnergySystem1Status","EnergySystem2Status","EnviroDataTable","EnvironmentalLogTable","EnvironmentYearlyTable","EWSTable","FinesCameraTable","FinishingDowntimeCodesTable","FinishingDowntimeTable","FinishingSawSpeedsTable","FinishingSnapShotTable","Flaker_1_Gate","Flaker_2_Gate","Flaker1Status","Flaker2Status","FlakerCameraHandshakeTable","FlakerDowntimeTable","FlakerDowTimeCodesTable","FlakerSnapShotTable","FlakerSnapShotTableNew","FlakesFinesTable","FlameXMiniFogTable","FLineSnapShot2Table","FlineSnapShotTable","FormatorDataTable","FormatorDefenserTable","FuelSystemsTable","GradingBlister","GradingScale","GradingThickness","GreenBinPickerRollCurrentTable","HeatEnergySnapShotTable","LabSampleTable","Line2PressData2Table","Line2PressDataTable","Line2PressManTable","LoaderBeltTorqueTable","LoaderPositioningTable","MasterWtTable","MatingChainMatPosTable","MillDowntimeTable","MTR_Preloader1Upper","MTR_Preloader2Lower","MTR_TransferBelt1","MTR_Transferbelt2","NotNearProductChangeTable","OpacityCalibrationTable","OpacityTable","PaintInventoryTable","PaintUsageTable","PikaDataTable","PLCForceTable","PLCWeights_Insert_Audit","PressChainPositioningTimes","PressChainTempTable","PressEfficeincyCodesTable","PressEfficiencyTable","PressHeatingTable","PressHydTable","PressLoadTable","PressParametersTable","PressProgramTable","PressScanTable","PressSnapShotTable","PressSnapShotTempTable","ProductTargetTable","RBSDowntimeTable","RBSSnapshotTable","RBSTotalUnitsTable","RcoTable","RMW_M16","RMW_TankAndMeterReading","RobotDataTable","RTO1DowntimeLogTable","RTO2DowntimeLogTable","RtoChamberTempTable","RTOFanVibrationTable","RtoTable","SawAmpsFirstPassTable","Scott_Dint","ScottDint_1_Table","ScottDint_2_Table","ShiftEnd_Stencil1Data","ShiftEndDataTable","ShiftTotalsTable","SSMAbortTable","Table42","TankLevelReadingShiftChange","tblChartProperties","tblFormingLineMCRange","tblProcessVariableLimits","TCODowntimeLogTable","TeafordSystem1StatusTable","TeafordSystem2StatusTable","ThermalOilStatus","TitleVTable","TnGDowntimeCodeTable","TnGDowntimeTable","TnGDustFanVibrationTable","TNGShiftProductChangeTable","UnitTable","UnloaderArmTempTable","UnloaderCyclesTable","UnloaderPositioningTable","VariabilityBlenderStop","VariabilityBoardWeight","VariabilityCrossWeight","VariabilityLineSpeed","VariabilityMoistureContent","VariabilityPMDITemp","VariabilityPressTemp","VariabilitySparkEvent","WeightControlTable","WespCausticSystem","WoodyardSnapShotTable"],"GP":["BlenderDowntimeTable","BlenderIDTable","BlenderTable","BlenderTotals","BlenderTotalsCont","CurrentProductTable","DB1Log_DataTable","DB2ILog_DataTable","DelugeValves","DownTime_Area_IDs","DownTime_Area_tbl","DowntimeTable","Dryer","DryerB","Energy","EnergyB","EnergyManagementRuntimeTable","EnviromentReportingTable","F1LogPush_DataTable","F2LogPush_DataTable","FinesManagement","FinishingDataTable","FinishingDataTable1hr","FormingDataTable","Furnace3GrateCam","Furnace4GrateCam","FurnaceCycleTimesTable","HMIAlarmAreaTable","HMIAlarmTable","MasterWtTable","MPCDryer3Table","MPCDryer4Table","NASmartGearboxDataTable","PikaWaxTable","PLC01","PressEfficiencyTable","PressLoadTable","PressTable","RimboardDataTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankAndMeterReading_20250227","RMW_TankAndMeterReading_20250327","RMW_TankLevel","RMW_TankLevel_20250227","RMW_TankLevel_20250327","RMW_TotalizedMeterReading","RMW_TotalizedMeterReading_20250227","RMW_TotalizedMeterReading_20250327","RSSQLTriggers","ShiftEndTable","SparkDetectTable","Strander1DataByEvent","Strander2DataByEvent","TanksTable","ThomasTable","VariabilityTable","VSDataTable","WoodroomDataTable"],"GUN":["_TagManager_Test_Audit","AirCompressors","Barcode_Error_Log","Barcode_NBInOutEntry_Insert_Audit","BlenderDowntimeTable","BlenderID","BlenderLiquidResinTable","BlenderModeTable","BlenderSnapShotTable","BlenderTotalizerTable","CraneTable","CrewTable","CurrentProductTable","DistributionConveyor1Table","DistributionConveyor2Table","DistributionConveyor3Table","DowntimeTable","Dryer1RTOTable","Dryer1Table","Dryer2RTOTable","Dryer2Table","Dryer3RTOTable","Dryer3Table","DryerDowntimeTable","DryerEfficiencyTable","DryFlakeReclaim","dtproperties","DustSystemPluggedTable","DustSystemStateTable","EnvironmentalLogTable","EnvironmentalTable","ESPTable","EWSDataTable","EWSProductTable","EWSProductTableAudit","EWSSummaryTable","ExhaustFanTable","FBeltRegPos","FinishingDownTimeTable","FinishingTable","FinishingUnitTable","FlakerEfficiencyTable","FlakerKnifechangeTable","FlameXTable","FLineSnapShotTable","FLineToScreenRetModeTable","FormerModeTable","FormingLineScanTable","GantryTable","GreenBinRackbackTable","GreenBins","GreenEndEfficiencyTable","Heater1Table","Heater2Table","Heater3Table","HeaterCommonTable","HMIAlarmsDryer","HMIAlarmsPress","LineSpeedTable","MasterPanelWeightTable","MasterWtTable","MatFormationKPI","MillEnergyUsageTable","MPCDryer1Table","MPCDryer2Table","MPCDryer3Table","OpacityMinuteGroupTable","PaintTankInventory","PanelDataTable","PanelPointThicknessTable","PanelThicknessTable","PreDryer2Table","PreDryer3Table","PressChains","PressEfficiency","PressHeatTable","PressLoadingCycle","PressloadTable","PressLoopTempEventTable","PressManTable","PressParametersTable","PressProductTable","PressProgramTable","PressScanTable","ProccessChangeLog","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RobotDataTable","RobotDataTable1","RobotDataTable2","ShiftEndTable","ShortBoardAtTransferTable","ShortMat_LOG","SmartsheetRowID","SSMBioFilterBypass","SSMDryerAbortsAllTable","SSMDryerAbortTable","SSMDryerShutdownTable","SSMDryerStartUpTable","SSMOpacityHighTable","SSMPPressShutDownTable","SSMPressStartUPTable","StandardPressloadTable","SuperScreenTable","sysdiagrams","TankFarmTable","tblChartProperties","tblCyclesPerHourTarget","tblDryerName","tblFieldDesc","tblFormingLineMCRange","tblLogPerHourTarget","tblProcessVariableLimits","tblTableDesc","TNGRejectRecordTable","TNGThicknessReadings","TrolleyTable","UnitTable","VibrationTable","Waferizer1MotorStartLogTable","Waferizer1Table","Waferizer2MotorStartLogTable","Waferizer2Table","Waferizer3MotorStartLogTable","Waferizer3Table","WaferizerDownTimeTable","WoodYardTable"],"HL":["Baghouse","BlenderDowntimeTable","BlenderIDTable","BlenderTotalizerContTable","BlenderTotalizerTable","BlendingDataTable","Boards","BookSaw","CurrentProductTable","CycleTimeLine1","Debarker1","Debarker2","DiagonalSaws","DiagSawsTable","DowntimeTable","DPCTestTable","Dryer1","Dryer2","Dryer3","DryerAbortTable","DryerDowntimeTable","DryerIDTable","EndOfShift","Energy","EnvironmentalDailySummaryTable","EnvironmentalDataTable","Finishing","FinishingDownTimeTable","FinishingRefeed700Conv","FormingDataTable","FTAEInstance","FuelSystemTable","Furnace1DataTable","Furnace2DataTable","Furnace3DataTable","Giben","GibenSawEfficiencyTable","LabCutTable","M16TotalsHistoryTESTZW","MachineCenterTable","MasterWtTable","Packaging","PaintBoothL1","PanelPointThicknessTable","PanelThicknessTable","Ponds","Preheating","PressBeltAndChains","PressBeltLubrication","PressEfficiencyTable","PressHeatTable","PressMPDataTable","PressScanTable","Product","QualityTable","RailcarBoilerTable","ReverseOsmosis1","ReverseOsmosis2","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RobotDataTable","Steam","Strander1","Strander1DataByEvent","Strander1KnifeChangeTable","Strander2","Strander2DataByEvent","Strander2KnifeChangeTable","Strander3","Strander3DataByEvent","Strander3KnifeChangeTable","StranderDowntimeTable","StranderEndOfShift","ThermalOilSystemDataTable","TriggerTable","UnitsMadeTable","VisionSmartDataTable","VPKPI","WasteManagement","WESP","WME_Testing","WoodRoomDataTable"],"HUG":["_TagManager_Test_Audit","_TO_DELETE_Barcode_InsertUnit_Audit","_to_delete_Barcode233InsertUnitAudit","_TO_DELETE_FormingDataTable","ATrainDataTable","Barcode_Error_Log","Barcode_NBInOutEntry_Insert_Audit","BiofilterDataTable","BlenderDetailDataTable","BlenderDowntimeTable","BlenderTotalizerTable","BlenderTotalizerTableCont","BlendingAtomizerTable","BlendingData10MinuteTable","BlendingDataTable","BlisterDataTable","BTrainDataTable","ConstraintsTableMPC","CraneGPS","CTrainDataTable","CurrentProductTable","dataviewDemo_StrandersData","DeltaV613DataTable","DownTimeEventsStop","DowntimeTable","Dryer","Dryer1DataByMinuteTable","Dryer1DataTable","Dryer1EfficiencyDataTable","Dryer2DataByMinuteTable","Dryer2DataTable","Dryer2EfficiencyDataTable","Dryer3DataByMinuteTable","Dryer3DataTable","Dryer3EfficiencyDataTable","DryerDowntime","DryerDowntimeTable","DryerEfficiencyTable","DryerRoutingDataTable","DustDataTable","EcoScanDataTable","EnviroShiftEndTable","EnviroTable","EpochNarrowTable","EpochTable","EWS_Area_Weight","EWSTrackStatusTable","FinishingDataTable","FinishingDownTimeTable","FinishingDowntimeTbl","FlakerKnifeChange","FlakerTransTable","FormingLine_Metric","FormingLineData10MinuteTable","FormingLineDataTable","FormingRecipeTable","FormingWigWagDataTable","FourFtUnitWeightDataTable","HeatEnergyATrainMPCDataTable","HeatEnergyBTrainMPCDataTable","HeatEnergyCTrainMPCDataTable","HeatEnergyDataTable","HeatEnergyEfficiency","HeatEnergyTableMPC","helper_MesMillVertical","JagEventID","JagEventTable","LiftStationDataTable","LogDataTable","LogLengthDataTable","LogLine1DataTable","LogLine2DataTable","MachineDowntimeTable","MachineIdTable","MatRejectTable","MatRejectTableII","MES_Mill_ProductsChanges","MES_Mill_ProductsChanges_TEST","MillStatusDataTable","MillUpdateHourly","MoistureCalibrationDataTable","Norbord40DataTable","OpacityCalibrationTable","OpacityCalibrationTypeTable","PackageBoilerTable","PanelScaleData10MinuteTable","PanelScaleDataTable","PanelThicknessTable","Parameters","PLCDataTable","Point6TGTable","PreHeaterByMinutesDataTable","PreHeaterDataTable","PreHeaterTableByMinute","PressChainLubeDataTable","PressEfficiencyTable","PressHeatDataTrans","PressInfeedDistancePressureDataTable","PressMotorsDataTable","PressOffsetsDataTable","PressOutfeedDistancePressureDataTable","PressOutfeedSawDataTable","PressRecipeTable","PressSpecificPressuresDataTable","ProcessDataTable","QualityMachineUtilizationDataTable","RCODataTable","RecipeDownLoadTime","ReleaseAgentDataTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RSSQLTriggerTable","RTOTableDryer1","RTOTableDryer2","RTOTableDryer3","SanderTG","SanderThicknessTable","SanderUnitTable","ShiftEndTable","SPCDistanceTable","SPCPressInfeedDistancePressureDataTable","SPCPressInfeedtest","SPCPressOutfeedDistancePressureDataTable","SPCPressureTable","SPCValveControlTable","SSMAbortStacks","SSMAbortTable","SSMBioFilterBypass","SSMDryerShutdownTable","SSMDryerStartUpTable","SSMOpacityHighTable","SSMPressShutdownTable","SSMPressStartUpTable","StackBinCountTable","StackBinsDataTable","Strander1DataByEvent","Strander1KnifeChangeTable","Strander2DataByEvent","Strander2KnifeChangeTable","StranderAreaTable","StranderDowntime","StranderDowntimeTable","StranderKnifeChangeTable","SuperSaw1DataTable","SuperSaw2DataTable","sysdiagrams","tblOpacityCalibrationType","TCOModeTypeTable","TempTable","Test","ThicknessBlisterTable","ThicknessModelTable","TriggerTable","UnitHeightDataTable","UnitWeightDataTable","UnitWeightTable","VibrationDataTable","WoodFuelSystemDataTable","WoodroomDataTable","WoodyardDataTable","WvCoWaxTrialTable"],"JEF":["_TEST","AROCFinishingAutoLabelPrinterTrend","BakeOutTable","Barcode_Error_Log","Barcode_InsertUnit_Audit","Barcode_PLCWIPMasterPanel_Insert_Audit","BlenderDowntimeTable","BlenderID","BlenderSnapShotTable","BlenderStopsTable","BlenderTotalizerTable","BlenderTotalizerTableCont","BlendingDataTable","BlendingSlackWax","BundleWeightDataTable","Consumption10MinuteTable","CrossCarriageTimersTable","CurrentProductTable","DowntimeRunningStatus","DowntimeTable","Dryer_CL","Dryer_SL","DryerClMPC","DryerDownTimeTable","DryerEfficiencyTable","DryerFireDumpDownTimeTable","DryerIDLookup","DryerMpcReport","DryerSlMPC","DryerSnapShotTable","EnvironmentalTable","ESPTable","EWSDataTable","EWSSummaryTable","FinishingDownTimeTable","FinishingLineHotLoadTable","FinishingTable","FinishingType","FlakerDownTimeTable","FlakerEastSLSnapshot","FlakerWestCLSnapshot","FLineSnapShotTable","FormerWeightTable","FuelTable","FurnaceCLGrateCamTable","FurnaceSLGrateCamTable","GlobeTable","GreenBin_CL","GreenBin_SL","KnifeChangeTable","LoaderMatLength","MasterWtTable","MatingChainSectionLengthTable","MillGeneral","MillSnapShotLineSpeedTable","MillSnapShotTable","NessTOH","NorthFurnanceTable","PanelBlisterTable","PanelPointThicknessTable","PanelThicknessTable","PressEfficiency","PressLoadPrecureOvercookTable","PressloadTable","PressLoadTroubleshoot","PressManTable","PressParametersTable","PressParameterTable","PressProgramTable","PressStatusDataTable","PressStepTable","PressTable","PressTemperatureTable","PressThermalOilTable","ProductTable","RawMaterialUnloadTable","RBS","RBSDowntime","RejectMatTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RobotDataTable","RTO_North","RTO_South","RTOPressTable","ShiftEndTable","SicoTable","SnapShot","SouthFurnance","SSMAbortTable","SSMDryerShutdownTable","SSMDryerStartUpTable","SSMPressShutDownTable","SSMPressStartUpTable","Strander1DataByEvent","Strander2DataByEvent","TankFarmLevels","ThomasTable","TowBarSensorTable","TransactionTriggerValuesTable","UnitWeightTable","WoodFlowTable","WoodyardSnapshotTable"],"JOA":["_tblTemp_FinishingDowntimeDetails","_to_delete_RCA_tblDowntimeOcurrences_NorbordLegacy","AtomizerDataTable","Barcode_Error_Log","Barcode_PLCWeights_Insert_Audit","BiofilterLogTable","BlenderDowntimeTable","BlenderID","BlendingTable","BlisterTable","CurrentProductTable","DebarkerDataTable","DownTimeEventsStop","Dryer1RTOTable","Dryer1Table","Dryer1VfdDataTable","Dryer2RTOTable","Dryer2Table","Dryer2VfdDataTable","Dryer3RTOTable","Dryer3Table","Dryer3VfdDataTable","DryerDowntimeTable","DryerEfficiency","dtproperties","DustTable","EndOfShift","EnergyMonitorTable","EnvironmentalTable","EpochTable","FinishingDowntimeTable","FinishingTable","FlakerTable","FormingLineTable","FourPoint0Table","HeatEnergyATrainTable","HeatEnergyBTrainTable","HeatEnergyCommonTable","HeatEnergyCTrainTable","HeatEnergyEfficiency","JagEventID","JagEventTable","KnifeChangeStrander1Table","KnifeChangeStrander2Table","LiftStation","LogDataTable","LogLengthTable","MesDataTable","MesShiftEndTable","MoistureCalibrationTable","OpacityCalibrationTable","OpacityCalibrationTypeTable","PaintBoothEventTable","PaintBoothTable1","PanelScaleTable","PreHeaterTable","PreHeaterTableByMinute","PressEfficiencyTable","PressEfficiencyTable_Temp","PressHeatTable","PressInfeedDistancePressureTable","PressMotorsTable","PressOffsetsTable","PressOutfeedDistancePressureTable","PressSpecificPressures","ProcessLabAggregatedDataTable","Products","QualityMachineUtilization","RawMaterialUsageCont","RawMaterialUsageTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankAndMeterReading_20250117","RMW_TankAndMeterReading_20250121","RMW_TankAndMeterReading_NEW_TANK_TAGS_to_delete","RMW_TankLevel","RMW_TotalizedMeterReading","RobotDataTable","ShiftEndTable","SmartWaxBlendingTable","SmartWaxRecipeTable","SmartWaxUsageCont","SSMAbortTable","StackBinTable","Strander1DataByEvent","Strander2DataByEvent","StranderArea","StranderDowntimeTable","StranderEndOfShift","Strapper1Table","Strapper2Table","TCOTable","Tisfoon","UnitHeightTable","UnitWeightTable","VariabilityBoardWeight","VibrationTable","WMSInSprocAudit","WoodyardTable"],"LAS":["Barcode_Error_Log","Barcode_StaplerData_Audit","BlenderDowntimeTable","BlenderID","BlendingDataTable","BookSawDownTime","CurrentProductTable","DebarkerDownTime","DebarkerShiftData","Dryer_1_DataTable","Dryer_2_DataTable","DryerDowntimeTable","DryerEfficiency","DryingDataTable","dtproperties","DustSysDataTable","ESPDownTimeTable","ESPTable","EventTableStrander","FillFactorStrander1","FillFactorStrander2","FinesBlendingTable","FinishingDataTable","FormingLineTable","HeatEnergyDataTable","HogBarkDataTable","LogLine_1_DataTable","LogLine_2_DataTable","LogLineDataTable","Magazine6DownTime","MasterWtTable","MillSnapShot","PanelPointThicknessTable","PanelThicknessTable","PanelWeightDataTable","PressDataTable","PressEfficency","PressloadTable","PressManTable","PressParametersTable","PressProgramTable","ProductionMaterialUsage","ProductRunAfterPressing","ProductRunBlending","ProductRunBlendingCont","ProductRunMatReject","ProductRunSetIDInfo","ProductRunSetIDInfo_Temp","ProductTrackingFinishing","PumpHouse_DataTable","RawMaterialUsage","ReportDataTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RobotDataTable","SaltonDownTimeTable","SanderDowntimeTable","SandingLineEfficiency","StackTableUpdates","Strander1DataByEvent","Strander1KnifeChangeTable","Strander2DataByEvent","Strander2KnifeChangeTable","StranderDownTime","StranderKnifeChangeCycles","StranderPowerDemand","StranderShiftData","StrapperDownTime","sysdiagrams","TableConversionLookups","tbl_Cons_Pressage_Delay","tbl_cont_effi_gaufrier","tbl_curprodsum","tbl_currentId","tbl_downtime","tbl_ecorceur_bytes","tbl_ecorceur_datapoint","tbl_ecorceur_float","tbl_ecorceur_integer","tbl_ecorceur_rapport_gau__Config","tbl_ecorceur_rapport_gau_ecor","tbl_encolage_bytes","tbl_encolage_datapoint","tbl_encolage_float","tbl_encolage_integer","tbl_energie_bytes","tbl_energie_datapoint","tbl_energie_float","tbl_energie_integer","tbl_Filtre_2_Geom_Lam_Str_1","tbl_Filtre_2_Geom_Lam_Str_2","tbl_Filtre_Geom_Lam_Str_1","tbl_FinLigne_Bytes","tbl_FinLigne_datapoint","tbl_FinLigne_Float","tbl_FinLigne_Integer","tbl_gauf_moy_pr_po_100_pas_str_2","tbl_gauf_pouss_vs_nbre_pa_Config","tbl_gauf_pouss_vs_nbre_pas","tbl_gauffrier_bytes","tbl_gauffrier_datapoint","tbl_gauffrier_float","tbl_gauffrier_integer","Tbl_Hum_Cent_0906","Tbl_Hum_Cent_0907","Tbl_Hum_Surf_0906","Tbl_Hum_Surf_0907","tbl_LigneFormation_Bytes","tbl_LigneFormation_datapoint","tbl_LigneFormation_Float","tbl_LigneFormation_Integer","tbl_Presse_bytes","tbl_presse_datapoint","tbl_Presse_Float","tbl_Presse_Integer","tbl_Presse_Rap_dech","tbl_Presse_Rap_dech_Config","tbl_Presse_Rap_Presse","tbl_Presse_Rap_Presse_Config","tbl_PresseNo","tbl_product","tbl_product_AuditChanges","tbl_product_con","tbl_product_Config","tbl_product_Config2","tbl_product_ForSchrock","tbl_rap_cons_encol","tbl_rap_cons_encol_Config","tbl_sechoir_bytes","tbl_Sechoir_datapoint","tbl_Sechoir_Float","tbl_sechoir_integer","tbl_Shift","tbl_vol_prod_dern_com","tblEncollageTable","tblFormationTable","UnitSize","UsagesDataTable","v_sechoir_Config"],"NAC":["_TagManager_Test_Audit","Barcode_Error_Log","Barcode_InsertUnit_Audit","Barcode_PLCWeights_Insert_Audit","Barcode_PLCWIPMasterPanel_Insert_Audit","BlenderDowntimeTable","BlenderID","BlenderTable","BlenderTotalsTable","BlenderTotalsTableCont","ConditionEvent","CurrentProductTable","DownTimeTable","DryerDownTimeTable","DryersTable","DunnageRobotTable","EastFlakerSnapshotTable","EnviroTable","EWSDataTable","EWSSummaryTable","FinesBlenderTable","FinishingDownTimeTable","FinishingTable","FlakerDownTimeTable","FlakerETable","FlakerWTable","FlineTable","FTAEInstance","FuelTable","IPDryerTable","IPTexasTable","KonusTable","LineSpeedTable","MasterPanelWeight","MasterPanelWeightTable","MasterWtTable","MatDumpsTable","MatDumpType","MatPositionTable","NessTable","NestecTable","OSBUnitsTable","PLC_BTable","PLC_CTable","PressEfficiencyTable","PressLoadTable","PressParametersTable","PressProgramTable","PressScanTable","PressTable","PressTemperatureTable","Product","Product_20251112","Product1","ProductAudit","ProductChangeLogTable","ProductCodeTable","ProductRatingTable","RMW_M16","RMW_TankAndMeterReading","RMW_TankLevel","RMW_TotalizedMeterReading","RobotDataTable","RTOTable","SawChassis_Table","ShiftEndTable","SimpleEvent","SSMAbortTable","StatsTable","StorageTankTable","tblChartProperties","tblFormingLineMCRange","tblProcessVariableLimits","tblTagNames","TeafordTable","TrackingEvent","WeightControlTable","wespTable","WestFlakerSnapShotTable","WoodyardTable"]}; // MILL_TABLES_PLACEHOLDER


  const SAWMILL_CODES=['NWB', 'DGM', 'MAP', 'OPE', 'HNM', 'LEO', 'ARM', 'NEW', 'JOY', 'ANG', 'RUS', 'MAN', 'FIT'];
  const SAWMILL_TABLES=['CanterDataTable', 'DowntimeDataTable', 'DrySorterDataTable', 'DryStackerDataTable', 'DryerDataTable', 'EdgerDataTable', 'EnergyDataTable', 'EnvironmentalDataTable', 'FuelEnergyDataTable', 'GreenSorterDataTable', 'GreenStackerDataTable', 'GreenTrimmerDataTable', 'KilnDataTable', 'MerchDataTable', 'PlanerDataTable', 'SawlineDataTable', 'ShiftEndDataTable', 'UtilityDataTable'];
  function isSawmill(){return SAWMILL_CODES.includes(document.getElementById("millSelect").value);}

  function getMillTables(){
    const mill=document.getElementById("millSelect").value;
    if(isSawmill()) return SAWMILL_TABLES;
    return MILL_TABLES[mill]||[];
  }

  function mkTableNameField(val){
    const wrap=document.createElement("div");wrap.style.position="relative";
    const tables=getMillTables();
    const hasTables=tables.length>0;

    // Search input (always shown)
    const inp=document.createElement("input");
    inp.type="text";inp.className="ci tableName-text";
    inp.value=val||"";
    inp.placeholder=hasTables?"Search or type table name...":"Type table name...";
    inp.setAttribute("autocomplete","off");

    // Dropdown list
    const drop=document.createElement("div");
    drop.style.cssText="position:absolute;top:100%;left:0;right:0;background:white;border:1.5px solid var(--wf);border-radius:6px;z-index:9999;max-height:180px;overflow-y:auto;display:none;box-shadow:0 4px 12px rgba(0,0,0,0.15);margin-top:2px;";

    function renderDrop(filter){
      drop.innerHTML="";
      if(!hasTables){drop.style.display="none";return;}
      const f=(filter||"").toLowerCase();
      const matches=tables.filter(t=>t.toLowerCase().includes(f));
      if(!matches.length){drop.style.display="none";return;}
      matches.forEach(t=>{
        const item=document.createElement("div");
        item.textContent=t;
        item.style.cssText="padding:7px 10px;font-size:12px;cursor:pointer;color:var(--g900);border-bottom:1px solid var(--g100);";
        item.addEventListener("mousedown",function(e){
          e.preventDefault();inp.value=t;drop.style.display="none";chkDups();
        });
        item.addEventListener("mouseenter",function(){this.style.background="var(--wf-light)";});
        item.addEventListener("mouseleave",function(){this.style.background="";});
        drop.appendChild(item);
      });
      // Show temporarily to measure, then position
      drop.style.display="block";
      drop.style.top="100%";
      drop.style.bottom="auto";
      const rect=inp.getBoundingClientRect();
      const spaceBelow=window.innerHeight-rect.bottom;
      const spaceAbove=rect.top;
      const dropH=Math.min(drop.scrollHeight,180);
      if(spaceBelow<dropH&&spaceAbove>spaceBelow){
        drop.style.top="auto";
        drop.style.bottom="100%";
        drop.style.marginTop="0";
        drop.style.marginBottom="2px";
        drop.style.boxShadow="0 -4px 12px rgba(0,0,0,0.15)";
      } else {
        drop.style.top="100%";
        drop.style.bottom="auto";
        drop.style.marginTop="2px";
        drop.style.marginBottom="0";
        drop.style.boxShadow="0 4px 12px rgba(0,0,0,0.15)";
      }
    }

    inp.addEventListener("focus",function(){renderDrop(this.value);});
    inp.addEventListener("input",function(){renderDrop(this.value);chkDups();});
    inp.addEventListener("blur",function(){setTimeout(()=>{drop.style.display="none";},200);});
    inp.addEventListener("keydown",function(e){
      const items=[...drop.querySelectorAll("div")];
      const activeIdx=items.findIndex(i=>i.classList.contains("tn-active"));
      if(e.key==="ArrowDown"){
        e.preventDefault();
        if(activeIdx>=0)items[activeIdx].classList.remove("tn-active");
        const next=items[activeIdx+1]||items[0];
        if(next){next.classList.add("tn-active");next.style.background="var(--wf-light)";}
      } else if(e.key==="ArrowUp"){
        e.preventDefault();
        if(activeIdx>=0)items[activeIdx].classList.remove("tn-active");
        const prev=items[activeIdx-1]||items[items.length-1];
        if(prev){prev.classList.add("tn-active");prev.style.background="var(--wf-light)";}
      } else if(e.key==="Enter"){
        const active=drop.querySelector(".tn-active");
        if(active){e.preventDefault();inp.value=active.textContent;drop.style.display="none";chkDups();}
      } else if(e.key==="Escape"){drop.style.display="none";}
    });

    wrap.appendChild(inp);wrap.appendChild(drop);
    return wrap;
  }
  function getTableName(r){return r.querySelector(".tableName-text")?.value||"";}

  function addRow(data={}){
    const tb=document.getElementById("tbody"),tr=document.createElement("tr");
    function td(cls,el){const c=document.createElement("td");if(cls)c.className=cls;if(el)c.appendChild(el);return c;}
    function inp(cls,val,ph){const i=document.createElement("input");i.type="text";i.className="ci "+cls;i.value=val||"";i.placeholder=ph||"";return i;}
    const tn=mkTableNameField(data.tableName||"");
    const tg=inp("tagName",data.tagName,"e.g. BladeSpeedRpm");
    const ip=inp("plcAddress",data.plcAddress,"192.168.x.x");
    const pp=inp("plcPath",data.plcPath,"PLC tag path");
    const ds=mkDTSelect(data.dataType);
    const uw=mkUnitsSearch(data.units||"");
    const fr=inp("transactionFrequency",data.transactionFrequency,"e.g. 5 s, 1 min, Unscheduled");
    const dc=inp("tagDescription",data.tagDescription,"What does this tag measure?");
    const mn=inp("min",data.min,"e.g. 0");
    const mx=inp("max",data.max,"e.g. 100");
    const ntd=document.createElement("td");ntd.className="cbtd";
    const cb=document.createElement("input");cb.type="checkbox";cb.className="newTable";if(data.newTable)cb.checked=true;ntd.appendChild(cb);
    const rtd=document.createElement("td");
    const rb=document.createElement("button");rb.className="rmbtn";rb.textContent="x";
    rb.onclick=function(){tr.remove();upEmpty();upJust();chkDups();};rtd.appendChild(rb);
    tr.appendChild(td("col-tbl",tn));tr.appendChild(td("col-tag",tg));tr.appendChild(td("col-ip",ip));
    tr.appendChild(td("col-path",pp));tr.appendChild(td("col-dt",ds));tr.appendChild(td("col-u",uw));
    tr.appendChild(td("col-fr",fr));tr.appendChild(td("col-desc",dc));tr.appendChild(td("col-mn",mn));
    tr.appendChild(td("col-mx",mx));tr.appendChild(ntd);tr.appendChild(rtd);
    tg.addEventListener("input",chkDups);

    ip.addEventListener("blur",()=>valIP(ip));tg.addEventListener("blur",()=>{valTag(tg);chkDups();});
    mn.addEventListener("blur",()=>valNum(mn,"Min"));mx.addEventListener("blur",()=>valNum(mx,"Max"));
    fr.addEventListener("input",upJust);
    tb.appendChild(tr);upEmpty();
    // Run validation immediately so warnings show on import
    valIP(ip);valTag(tg);valNum(mn,"Min");valNum(mx,"Max");
  }

  const IPR=/^(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?::\d{1,5})?$/;
  function clrMsg(el){el.parentElement.querySelectorAll(".msg,.dup-msg").forEach(m=>m.remove());el.classList.remove("err","dup","warn");}
  function addMsg(el,txt,t){clrMsg(el);const d=document.createElement("div");d.className="msg warn";d.textContent=txt;el.parentElement.appendChild(d);el.classList.add("warn");}
  function valIP(el){
    clrMsg(el);
    if(!el.value)return true;
    const v=el.value.trim();
    if(/,/.test(v)){addMsg(el,'Use . instead of , in IP address (e.g. 192.168.1.10)',"warn");return true;}
    if(!IPR.test(v)){addMsg(el,"Prefer a direct IP address (e.g. 192.168.1.10)","warn");return true;}
    return true;
  }
  function valTag(el){
    clrMsg(el);
    const v=el.value.trim();
    if(!v)return true;
    if(/_/.test(v)){addMsg(el,"Avoid underscores in tag name (e.g. BladeSpeedRpm)","warn");return true;}
    return true;
  }
  function valNum(el,lbl){
    clrMsg(el);
    if(!el.value)return true;
    const v=el.value.trim();
    if(/,/.test(v)){addMsg(el,'Use . instead of , (e.g. 1.5)',"warn");return true;}
    if(!/^-?\d*\.?\d*$/.test(v)){addMsg(el,lbl+" must be a number","warn");return true;}
    return true;
  }
  function chkDups(){
    const rows=document.querySelectorAll("#tbody tr");
    const ct={};
    rows.forEach(r=>{
      const t=getTableName(r).trim().toLowerCase();
      const n=(r.querySelector(".tagName")?.value||"").trim().toLowerCase();
      if(n){const k=t+"|"+n;ct[k]=(ct[k]||0)+1;}
    });
    rows.forEach(r=>{
      const tg=r.querySelector(".tagName");
      const t=getTableName(r).trim().toLowerCase();
      const n=(tg?.value||"").trim().toLowerCase();
      const isDup=!!(n&&ct[t+"|"+n]>1);
      if(tg){
        // Remove existing dup messages
        tg.parentElement.querySelectorAll(".dup-msg").forEach(m=>m.remove());
        tg.classList.toggle("dup",isDup);
        if(isDup){
          tg.classList.add("warn");
          const d=document.createElement("div");
          d.className="msg warn dup-msg";
          d.textContent="Duplicate tag name in this table";
          tg.parentElement.appendChild(d);
        } else {
          // Not a dup — re-run valTag so the underscore warning (and its yellow box) is preserved
          valTag(tg);
        }
      }
    });
    return Object.values(ct).every(c=>c<=1);
  }
  function upJust(){const h=Array.from(document.querySelectorAll(".transactionFrequency")).some(s=>s.value.trim().toLowerCase().startsWith("1 s")||s.value.trim()==="1s");document.getElementById("justBox").style.display=h?"block":"none";}
  function valAll(){
    let ok=true;
    document.querySelectorAll("#tbody tr").forEach(r=>{
      const tagEl=r.querySelector(".tagName");
      const minEl=r.querySelector(".min");
      const maxEl=r.querySelector(".max");
      const ipEl=r.querySelector(".plcAddress");
      // Force clear and re-run all validations so messages show immediately
      if(tagEl){clrMsg(tagEl);ok=valTag(tagEl)&&ok;}
      if(minEl){clrMsg(minEl);ok=valNum(minEl,"Min")&&ok;}
      if(maxEl){clrMsg(maxEl);ok=valNum(maxEl,"Max")&&ok;}
      if(ipEl){clrMsg(ipEl);valIP(ipEl);}
    });
    return ok;
  }
  function clearAll(){if(!confirm("Clear all rows?"))return;document.getElementById("tbody").innerHTML="";document.getElementById("pasteArea").value="";upEmpty();upJust();}
  function toggleNT(v){document.querySelectorAll(".newTable").forEach(cb=>cb.checked=v);}
  function esc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
  function getSubj(){const mill=document.getElementById("millSelect").value;const n=document.querySelectorAll("#tbody tr").length;const d=new Date().toLocaleDateString();const u=document.getElementById("urgFlag").checked;return mill+" Tag Request: "+n+" tags on "+d+(u?" [URGENT]":"");}
  function copyTable(){
    valAll();chkDups(); // show warnings but don't block
    const rows=document.querySelectorAll("#tbody tr");
    let html='<table border="1" style="border-collapse:collapse;font-family:sans-serif;font-size:13px"><thead><tr><th>TableName</th><th>TagName</th><th>PLCIPAddress</th><th>PLCPath</th><th>DataType</th><th>Units</th><th>TransactionFrequency</th><th>TagDescription</th><th>Min</th><th>Max</th><th>NewTable?</th></tr></thead><tbody>';
    rows.forEach(r=>{const dt=normDT(r.querySelector(".dataType-input")?.value||"");const uv=(()=>{const v=r.querySelector(".units-value")?.value||"";const m=UOM.find(u=>(u.u+" ("+u.d+")").toLowerCase()===v.toLowerCase());return m?m.u:v;})();html+="<tr><td>"+esc(getTableName(r))+"</td><td>"+esc(r.querySelector(".tagName")?.value||"")+"</td><td>"+esc(r.querySelector(".plcAddress")?.value||"")+"</td><td>"+esc(r.querySelector(".plcPath")?.value||"")+"</td><td>"+esc(dt)+"</td><td>"+esc(uv)+"</td><td>"+esc(r.querySelector(".transactionFrequency")?.value||"")+"</td><td>"+esc(r.querySelector(".tagDescription")?.value||"")+"</td><td>"+esc(r.querySelector(".min")?.value||"")+"</td><td>"+esc(r.querySelector(".max")?.value||"")+"</td><td>"+(r.querySelector(".newTable")?.checked?"Yes":"No")+"</td></tr>";});
    html+="</tbody></table>";
    const a1=Array.from(document.querySelectorAll(".transactionFrequency")).some(s=>s.value.trim().toLowerCase().startsWith("1 s")||s.value.trim()==="1s");
    if(a1){const j=document.getElementById("just1s").value||"";html+="<br><b>1s justification:</b> "+esc(j);}
    if(navigator.clipboard&&navigator.clipboard.write){navigator.clipboard.write([new ClipboardItem({"text/html":new Blob([html],{type:"text/html"})})]).then(()=>alert("Copied! Paste into your email.")).catch(()=>fb(html));}else fb(html);
  }
  function fb(html){const d=document.createElement("div");d.style.cssText="position:fixed;left:-9999px";d.contentEditable="true";d.innerHTML=html;document.body.appendChild(d);const r=document.createRange();r.selectNodeContents(d);const s=window.getSelection();s.removeAllRanges();s.addRange(r);try{document.execCommand("copy");alert("Copied! Paste into your email.");}catch(e){alert("Copy failed.");}document.body.removeChild(d);}
  function doSubmit(){
    valAll();chkDups(); // run to show warnings/errors but don't block
    if(!confirm("Have all tags been confirmed against the TagsRUs standard?"))return;
    const subj=getSubj();window.location.href="mailto:wfm-tagsrus@westfraser.com?subject="+encodeURIComponent(subj)+"&body="+encodeURIComponent("Please paste your copied table into the body of this email.");
  }
  window.addEventListener("DOMContentLoaded",()=>{updateUI();upEmpty();});
</script>