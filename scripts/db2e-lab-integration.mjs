import assert from "node:assert/strict";
import fs from "node:fs";
import zlib from "node:zlib";
import XLSX from "xlsx";

const patch=JSON.parse(zlib.gunzipSync(Buffer.from(["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"].join(""),"base64")).toString("utf8"));
const workbookPath="data/f1_db.xlsx";
const wb=XLSX.readFile(workbookPath,{cellFormula:true,cellStyles:true});
const unwrap=(v)=>v&&typeof v==="object"&&!Array.isArray(v)?unwrap(v.result??v.value??v):v;
const str=(v)=>String(unwrap(v)??"").trim();
const num=(v)=>Number(unwrap(v));
const rows=(n)=>XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:"",raw:true});
const headers=(n)=>XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,blankrows:false,raw:true})[0]||[];
const did=(r)=>str(r.driver_id||r.person_id||r.id);
const tid=(r)=>str(r.team_id||r.constructor_id||r.team||r.constructor);
const yr=(r)=>num(r.year||r.season_year||r.season);

const pDrivers=patch["drivers_patch.json"],pProfiles=patch["driver_rating_profiles_patch.json"],pSnapshots=patch["historical_rating_snapshots_1980_patch.json"],pOpening=patch["driver_opening_state_1980_patch.json"];
assert.ok(pDrivers&&pProfiles&&pSnapshots&&pOpening,"DB2E patch payload incomplete");
const expectedIds=Array.from({length:26},(_,i)=>"d_"+String(867+i).padStart(4,"0"));
for(const data of [pDrivers,pProfiles,pSnapshots,pOpening]){
  assert.equal(data.length,26);
  assert.deepEqual(data.map(did),expectedIds);
}
for(const n of ["drivers","driver_rating_profiles","historical_rating_snapshots","contracts"])assert.ok(wb.Sheets[n],`Missing canonical sheet ${n}`);

function appendUnique(name,data,keyFn){
  const existing=rows(name),byKey=new Map(existing.map(r=>[keyFn(r),r]).filter(([k])=>k)),add=[];
  for(const row of data){
    const key=keyFn(row),prior=byKey.get(key);
    if(prior){
      const a=str(prior.display_name||prior.driver_name||prior.name),b=str(row.display_name||row.driver_name||row.name);
      assert.ok(!a||!b||a===b,`${name}: collision at ${key}: ${a} vs ${b}`);
      continue;
    }
    add.push(row);byKey.set(key,row);
  }
  if(add.length)XLSX.utils.sheet_add_json(wb.Sheets[name],add,{header:headers(name),skipHeader:true,origin:-1});
  return add.length;
}
const addedDrivers=appendUnique("drivers",pDrivers,did);
const addedProfiles=appendUnique("driver_rating_profiles",pProfiles,did);
const addedSnapshots=appendUnique("historical_rating_snapshots",pSnapshots,r=>`${yr(r)}|${did(r)}`);

const allDrivers=rows("drivers"),driverById=new Map(allDrivers.map(r=>[did(r),r]).filter(([id])=>id));
const snapshots1980=rows("historical_rating_snapshots").filter(r=>yr(r)===1980);
const contracts1980=rows("contracts").filter(r=>yr(r)===1980&&did(r));
const raceRole=(r)=>/main|second|race/i.test(str(r.role||r.position||r.contract_role));
const contractByDriver=new Map();
for(const c of contracts1980){const id=did(c),prior=contractByDriver.get(id);if(!prior||(!raceRole(prior)&&raceRole(c)))contractByDriver.set(id,c);}

const contractOpening=(id,c,name)=>{
  const race=raceRole(c),role=str(c.role||c.position||c.contract_role||"Driver");
  return {year:1980,opening_date:"1980-01-01",driver_id:id,display_name:name,
    opening_world_status:race?"F1_CONTRACTED_RACE_SEAT":"F1_TEAM_COMMITMENT",
    opening_availability:race?"F1_CONTRACTED":"F1_TEAM_COMMITMENT_NON_RACE",
    opening_team_id:tid(c)||null,opening_team_name:str(c.team_name||c.name)||null,opening_role:role||null,
    series_context:"FORMULA_1",runtime_visibility:"ACTIVE_WORLD",runtime_market_policy:"TRANSFER_RULES",
    confidence:"HIGH",qa_status:"BASELINE_CONTRACT_DERIVED",source_ids:"CANONICAL_CONTRACTS_1980",
    season_world_status_reference:"PRESEASON_CONTRACT_SEED",notes:"Derived only from canonical 1980 contract seed; no race result is used."};
};

const generated=new Map();
for(const s of snapshots1980){
  const id=did(s);if(!id)continue;
  const d=driverById.get(id)||{},name=str(s.display_name||d.display_name||id),death=str(d.death_date);
  if(death&&death<"1980-01-01"){
    generated.set(id,{year:1980,opening_date:"1980-01-01",driver_id:id,display_name:name,opening_world_status:"DECEASED",opening_availability:"DECEASED_UNAVAILABLE",series_context:null,runtime_visibility:"EXCLUDE_ACTIVE_WORLD",runtime_market_policy:"BLOCKED",confidence:"HIGH",qa_status:"BASELINE_DERIVED",source_ids:"CANONICAL_MASTER",season_world_status_reference:"PRESEASON_MASTER",notes:"Excluded because canonical death_date is before opening date."});continue;
  }
  const c=contractByDriver.get(id);if(c){generated.set(id,contractOpening(id,c,name));continue;}
  const w=str(s.world_status).toUpperCase();
  if(w.startsWith("PROSPECT_"))generated.set(id,{year:1980,opening_date:"1980-01-01",driver_id:id,display_name:name,opening_world_status:"PROSPECT",opening_availability:"ACTIVE_OTHER_SERIES_NO_F1_SEAT",series_context:"LOWER_OR_EXTERNAL_SERIES",runtime_visibility:"ACTIVE_WORLD",runtime_market_policy:"APPROACHABLE",confidence:"MEDIUM",qa_status:"BASELINE_CONSERVATIVE_MAP",source_ids:"HISTORICAL_RATING_SNAPSHOT_1980",season_world_status_reference:w,notes:"Conservative pre-season mapping; no 1980 team/result assignment inferred."});
  else if(w.startsWith("RECENT_F1_EXIT_")||w==="F1_GAP_RETURN_CANDIDATE")generated.set(id,{year:1980,opening_date:"1980-01-01",driver_id:id,display_name:name,opening_world_status:"OUT_OF_F1_NO_SEAT",opening_availability:"AVAILABLE",series_context:"F1_RECENT_EXIT",runtime_visibility:"ACTIVE_WORLD",runtime_market_policy:"APPROACHABLE",confidence:"MEDIUM",qa_status:"BASELINE_CONSERVATIVE_MAP",source_ids:"HISTORICAL_RATING_SNAPSHOT_1980",season_world_status_reference:w,notes:"Conservative pre-season mapping; no 1980 team/result assignment inferred."});
  else generated.set(id,{year:1980,opening_date:"1980-01-01",driver_id:id,display_name:name,opening_world_status:"NEEDS_RESEARCH",opening_availability:"MARKET_STATUS_RESEARCH",series_context:"FORMULA_1",runtime_visibility:"ACTIVE_WORLD",runtime_market_policy:"VISIBLE_BLOCK_DIRECT_NEGOTIATION",confidence:"LOW",qa_status:"BASELINE_RESEARCH_REQUIRED",source_ids:"HISTORICAL_RATING_SNAPSHOT_1980",season_world_status_reference:w||"UNRESOLVED",notes:"No canonical pre-season contract found; kept visible but direct negotiation is blocked."});
}
for(const c of contracts1980){const id=did(c);if(!id||generated.has(id))continue;const d=driverById.get(id)||{};generated.set(id,contractOpening(id,c,str(d.display_name||c.driver_name||id)));}
for(const r of pOpening)generated.set(did(r),{...r});

// DB2E opening-state correction: Shadow entered David Kennedy as car #18 from
// the opening 1980 rounds. Geoff Lees only joined the #17 later, so neither
// the stale test-driver label nor later-season substitution may define Jan 1.
if(driverById.has("d_0225")){
  generated.set("d_0225",{
    year:1980,
    opening_date:"1980-01-01",
    driver_id:"d_0225",
    display_name:"David Kennedy",
    opening_world_status:"F1_CONTRACTED_RACE_SEAT",
    opening_availability:"F1_CONTRACTED",
    opening_team_id:"t_0015",
    opening_team_name:"Shadow",
    opening_role:"second_driver",
    series_context:"FORMULA_1",
    runtime_visibility:"ACTIVE_WORLD",
    runtime_market_policy:"TRANSFER_RULES",
    confidence:"MEDIUM",
    qa_status:"PRESEASON_SHADOW_OPENING_REVIEW",
    source_ids:"SHADOW_1980_OPENING_REVIEW",
    season_world_status_reference:"OPENING_ENTRY",
    notes:"Opening Shadow #18 race seat. Geoff Lees is a later-season replacement and is not backdated to Jan 1."
  });
}

assert.equal(pOpening.filter(r=>str(r.opening_team_id)).length,0,"DB2E newcomers must not receive invented F1 team commitments");
assert.equal(pSnapshots.filter(r=>str(r.data_cutoff)!=="1979-12-31").length,0,"DB2E snapshots must use 1979-12-31 cutoff");
const opening1980=[...generated.values()].sort((a,b)=>did(a).localeCompare(did(b)));
const openingIds=new Set(opening1980.map(did));
for(const id of expectedIds)assert.ok(openingIds.has(id),`Opening state missing ${id}`);
for(const s of snapshots1980)assert.ok(openingIds.has(did(s)),`Opening state missing baseline driver ${did(s)}`);

let preserved=[];
if(wb.Sheets.driver_opening_state){
  preserved=rows("driver_opening_state").filter(r=>yr(r)!==1980);
  delete wb.Sheets.driver_opening_state;
  const i=wb.SheetNames.indexOf("driver_opening_state");if(i>=0)wb.SheetNames.splice(i,1);
}
const openingHeaders=["year","opening_date","driver_id","display_name","opening_world_status","opening_availability","opening_team_id","opening_team_name","opening_role","series_context","runtime_visibility","runtime_market_policy","confidence","qa_status","source_ids","season_world_status_reference","notes"];
XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet([...preserved,...opening1980],{header:openingHeaders}),"driver_opening_state");

const daniel=opening1980.find(r=>did(r)==="d_0880");
assert.equal(str(daniel.opening_world_status),"NEEDS_RESEARCH");
assert.equal(str(daniel.opening_availability),"MARKET_STATUS_RESEARCH");
assert.ok(str(daniel.runtime_market_policy).includes("BLOCK_DIRECT_NEGOTIATION"));

if(process.argv.includes("--overlay-json")){
  const dataDir="public/data";
  const readJson=(name)=>JSON.parse(fs.readFileSync(`${dataDir}/${name}`,"utf8"));
  const writeJson=(name,data)=>fs.writeFileSync(`${dataDir}/${name}`,JSON.stringify(data,null,2)+"\n","utf8");
  const appendBy=(base,add,keyFn)=>{
    const out=[...base],seen=new Set(base.map(keyFn).filter(Boolean));
    for(const row of add){
      const key=keyFn(row);
      if(!key||seen.has(key))continue;
      out.push(row);seen.add(key);
    }
    return out;
  };
  const normalizeDriver=(row)=>({
    ...row,
    id:row.driver_id,
    name:row.display_name,
    nationality:row.country_name||row.country_code||null,
    birthdate:row.dob||null,
    birthdate_iso:row.dob||null,
    prefered_number:row.prefered_number==null?null:Number(row.prefered_number),
  });
  const jsonDrivers=appendBy(readJson("drivers.json"),pDrivers.map(normalizeDriver),did);
  const jsonProfiles=appendBy(readJson("driver_rating_profiles.json"),pProfiles,did);
  const jsonSnapshots=appendBy(readJson("historical_rating_snapshots.json"),pSnapshots,r=>`${yr(r)}|${did(r)}`);
  let jsonOpening=[];
  try{jsonOpening=readJson("driver_opening_state.json").filter(r=>yr(r)!==1980);}catch{}
  jsonOpening.push(...opening1980);
  writeJson("drivers.json",jsonDrivers);
  writeJson("driver_rating_profiles.json",jsonProfiles);
  writeJson("historical_rating_snapshots.json",jsonSnapshots);
  writeJson("driver_opening_state.json",jsonOpening);
  console.log("DB2E JSON overlay applied",JSON.stringify({
    drivers:jsonDrivers.length,
    profiles:jsonProfiles.length,
    snapshots1980:jsonSnapshots.filter(r=>yr(r)===1980).length,
    opening1980:opening1980.length,
    openingRaceSeats:opening1980.filter(r=>str(r.opening_world_status)==="F1_CONTRACTED_RACE_SEAT").length
  }));
  process.exit(0);
}

for(const id of expectedIds){
  assert.ok(new Set(rows("drivers").map(did)).has(id),`drivers missing ${id}`);
  assert.ok(new Set(rows("driver_rating_profiles").map(did)).has(id),`profiles missing ${id}`);
  assert.ok(new Set(rows("historical_rating_snapshots").map(r=>`${yr(r)}|${did(r)}`)).has(`1980|${id}`),`snapshot missing ${id}`);
}
if(process.argv.includes("--apply"))XLSX.writeFile(wb,workbookPath,{bookType:"xlsx",compression:true});
console.log(JSON.stringify({db2e:"v21",addedDrivers,addedProfiles,addedSnapshots,drivers:rows("drivers").length,profiles:rows("driver_rating_profiles").length,snapshots1980:snapshots1980.length,opening1980:opening1980.length,canonicalContracts1980:contracts1980.length,curatedReady:pOpening.filter(r=>str(r.qa_status)==="READY_CANONICAL_OPENING").length,curatedBlocked:pOpening.filter(r=>str(r.qa_status)==="OPENING_RESEARCH_BLOCKED").length,wroteWorkbook:process.argv.includes("--apply")},null,2));
