import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { isRaceDriverContract } from "../src/domain/contractRoles.js";
import { openingStateNegotiationBlockReason } from "../src/domain/driverOpeningState.js";

const root=process.cwd();
const dataDir=path.join(root,"public","data");
const read=(name)=>JSON.parse(fs.readFileSync(path.join(dataDir,name),"utf8"));

const expected=new Map([
  ["d_0867","Tom Gloy"],
  ["d_0868","Phil Dowsett"],
  ["d_0869","Roberto Ravaglia"],
  ["d_0870","Bruno Eichmann"],
  ["d_0871","Carlo Rossi"],
  ["d_0872","Enzo Coloni"],
  ["d_0873","Franz Konrad"],
  ["d_0874","Helmut Bross"],
  ["d_0875","Michael Korten"],
  ["d_0876","Hans-Georg Bürger"],
  ["d_0877","Alain Ferté"],
  ["d_0878","Bernard Devaney"],
  ["d_0879","Brett Riley"],
  ["d_0880","Daniel Hintzy"],
  ["d_0881","Daniele Albertin"],
  ["d_0882","David Sears"],
  ["d_0883","Eddie Jordan"],
  ["d_0884","Frank Jelinski"],
  ["d_0885","Guido Pardini"],
  ["d_0886","Harald Brutschin"],
  ["d_0887","Herbert Lingmann"],
  ["d_0888","Kurt Thiim"],
  ["d_0889","Mike White"],
  ["d_0890","Peter Kroeber"],
  ["d_0891","Rob Wilson"],
  ["d_0892","Thierry Tassin"],
]);

const ids=[...expected.keys()];
const idSet=new Set(ids);
const nameOf=(row)=>String(row?.display_name||row?.driver_name||row?.name||"").trim();
const byId=(rows)=>new Map(rows.map((r)=>[String(r.driver_id||r.id||""),r]).filter(([id])=>id));

test("DB2E v21 canonical runtime surfaces contain all 26 staged drivers exactly once",()=>{
  const drivers=read("drivers.json");
  const profiles=read("driver_rating_profiles.json");
  const snapshots=read("historical_rating_snapshots.json").filter((r)=>Number(r.year)===1980);
  const opening=read("driver_opening_state.json").filter((r)=>Number(r.year)===1980);

  for(const [label,rows] of [["drivers",drivers],["profiles",profiles],["1980 snapshots",snapshots],["1980 opening state",opening]]){
    const hits=rows.filter((r)=>idSet.has(String(r.driver_id||r.id||"")));
    assert.equal(hits.length,26,`${label}: expected 26 DB2E rows, found ${hits.length}`);
    assert.equal(new Set(hits.map((r)=>String(r.driver_id||r.id))).size,26,`${label}: duplicate staged driver IDs`);
  }

  const driversById=byId(drivers);
  const profilesById=byId(profiles);
  const snapshotsById=byId(snapshots);
  const openingById=byId(opening);

  for(const [id,name] of expected){
    assert.equal(nameOf(driversById.get(id)),name,`${id}: driver master name drift`);
    assert.equal(nameOf(profilesById.get(id)),name,`${id}: rating profile name drift`);
    assert.equal(nameOf(snapshotsById.get(id)),name,`${id}: 1980 snapshot name drift`);
    assert.equal(nameOf(openingById.get(id)),name,`${id}: opening-state name drift`);

    const p=profilesById.get(id);
    assert.ok(String(p.development_curve||"").trim(),`${id}: missing development curve`);
    assert.ok(Number.isFinite(Number(p.peak_ability)),`${id}: missing peak ability`);
    for(const field of ["peak_pace","peak_qualifying","peak_racecraft","peak_consistency","base_aggression","base_crash_likelihood"]){
      assert.ok(Number.isFinite(Number(p[field])),`${id}: missing ${field}`);
    }

    const s=snapshotsById.get(id);
    assert.equal(String(s.data_cutoff),"1979-12-31",`${id}: temporal firewall cutoff changed`);
    const ca=Number(s.current_ability),pa=Number(s.peak_ability);
    assert.ok(Number.isFinite(ca)&&Number.isFinite(pa),`${id}: CA/PA must be numeric`);
    assert.ok(ca<=pa,`${id}: current ability cannot exceed peak ability`);
    assert.ok(ca>=45&&ca<=66,`${id}: current ability outside DB2E calibration envelope: ${ca}`);
    assert.ok(pa>=60&&pa<=85,`${id}: peak ability outside DB2E calibration envelope: ${pa}`);
  }
});

test("DB2E opening state preserves the Jan-1 temporal firewall and Daniel Hintzy research block",()=>{
  const opening=read("driver_opening_state.json").filter((r)=>Number(r.year)===1980&&idSet.has(String(r.driver_id)));
  assert.equal(opening.length,26);

  for(const row of opening){
    const id=String(row.driver_id);
    assert.equal(String(row.opening_date),"1980-01-01",`${id}: opening date must be Jan 1 1980`);
    assert.equal(String(row.opening_team_id||""),"",`${id}: feeder candidate received an invented F1 opening team`);
    assert.notEqual(String(row.runtime_visibility||"").toUpperCase(),"EXCLUDE_ACTIVE_WORLD",`${id}: staged driver unexpectedly hidden`);
  }

  const blocked=opening.filter((r)=>String(r.qa_status)==="OPENING_RESEARCH_BLOCKED");
  assert.deepEqual(blocked.map((r)=>String(r.driver_id)),["d_0880"],"Daniel Hintzy must be the only unresolved Batch-1 opening status");

  const daniel=blocked[0];
  assert.equal(String(daniel.opening_world_status),"NEEDS_RESEARCH");
  assert.equal(String(daniel.opening_availability),"MARKET_STATUS_RESEARCH");
  assert.equal(openingStateNegotiationBlockReason(daniel,1980),"market_status_research");

  const ready=opening.filter((r)=>String(r.qa_status)==="READY_CANONICAL_OPENING");
  assert.equal(ready.length,25,"25 Batch-1 drivers should be opening-state ready");
});

test("DB2E 1980 Season Pack materializes all staged drivers without F1 backdating",()=>{
  const pack=JSON.parse(fs.readFileSync(path.join(dataDir,"seasons","1980","season.json"),"utf8"));
  assert.equal(pack.validation?.ok,true,JSON.stringify(pack.validation));
  assert.equal(pack.ratingModel,"R2B");

  const state=pack.state||{};
  assert.equal(state.teams.length,15,"1980 must expose 15 managerial teams");
  assert.equal(state.calendar.length,14,"1980 must expose 14 GPs");
  assert.equal(state.drivers.length,109,"1980 DB2E opening world should expose 109 drivers");
  assert.equal(state.driverOpeningState.length,109,"1980 opening state must cover the full visible opening world");

  const raceOpening=(state.driverOpeningState||[]).filter((r)=>String(r.opening_world_status)==="F1_CONTRACTED_RACE_SEAT");
  assert.equal(raceOpening.length,28,"1980 opening state should contain 28 historical race entries");

  const teamIds=state.teams.map((t)=>String(t.team_id||""));
  assert.equal(new Set(teamIds).size,15,"1980 contains duplicate managerial team identities");
  assert.ok(teamIds.every((id)=>/^t_\d+$/.test(id)),`invalid/ghost team identity: ${teamIds.find((id)=>!/^t_\d+$/.test(id))}`);

  const driverById=byId(state.drivers||[]);
  const ratingById=byId(state.driverRatings||[]);
  const openingById=byId(state.driverOpeningState||[]);
  const stagedContracts=(state.contracts||[]).filter((c)=>idSet.has(String(c.driver_id||"")));

  assert.equal(stagedContracts.length,0,"DB2E feeder candidates must not be backdated into F1 contracts");

  for(const [id,name] of expected){
    const d=driverById.get(id);
    const r=ratingById.get(id);
    const o=openingById.get(id);
    assert.ok(d,`${id}: missing from materialized 1980 driver world`);
    assert.ok(r,`${id}: missing materialized R2B rating`);
    assert.ok(o,`${id}: missing materialized opening state`);
    assert.equal(nameOf(d),name,`${id}: materialized name drift`);
    assert.equal(Number(r.year),1980,`${id}: rating snapshot is not exact 1980`);
    assert.equal(String(r.source),"historical_rating_snapshot_r2b",`${id}: rating did not come from the R2B 1980 snapshot`);
  }

  const daniel=driverById.get("d_0880");
  assert.equal(daniel.canHireF1,false,"Daniel Hintzy must remain blocked from direct F1 negotiation until researched");

  const shadow=(state.teams||[]).find((t)=>String(t.team_name||t.name)==="Shadow");
  assert.ok(shadow,"1980 Shadow missing");
  const shadowRace=(state.contracts||[])
    .filter((c)=>String(c.team_id)===String(shadow.team_id)&&isRaceDriverContract(c))
    .map((c)=>String(c.driver_id))
    .sort();
  assert.deepEqual(shadowRace,["d_0140","d_0225"],"Shadow opening seats must be Johansson + Kennedy; later Geoff Lees must not be backdated");
});
