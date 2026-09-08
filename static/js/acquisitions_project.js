// Tracts carry a county NAME; the override is keyed by FIPS.
const COUNTY_FIPS = {
  Harris: '48201', 'Fort Bend': '48157', Montgomery: '48339', Brazoria: '48039',
  Galveston: '48167', Liberty: '48291', Waller: '48473', Chambers: '48071',
  'San Jacinto': '48407', Walker: '48471', Grimes: '48185', Washington: '48477',
  Austin: '48015', Madison: '48313', Brazos: '48041',
};

/* Ported from the standalone Acquisitions GIS project page.
 * Static file, so Jinja never parses it. Server values arrive on window.ACQ_*.
 */
const PROJECT_ID = window.ACQ_PROJECT_ID;
function esc(s){return String(s||'').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function fmt(n, d){d=d==null?1:d;return (+n||0).toLocaleString('en-US',{maximumFractionDigits:d});}

// Centroid of the project's tracts — the origin for every directions link.
// Computed once and cached; `projectData` is a module-level let assigned in
// loadProject(), so guard against being called before that resolves.
let _projCentroid = null;
function projectCentroid() {
  if (_projCentroid) return _projCentroid;
  const p = (typeof projectData !== 'undefined' && projectData) ? projectData : null;
  if (!p || !p.tracts || !p.tracts.length) return null;
  try {
    const b = L.latLngBounds([]);
    for (const t of p.tracts) {
      if (t.geometry) b.extend(L.geoJSON(t.geometry).getBounds());
    }
    if (!b.isValid()) return null;
    const c = b.getCenter();
    _projCentroid = { lat: c.lat, lon: c.lng };
    return _projCentroid;
  } catch { return null; }
}

// Clickable distance → Google Maps driving directions from the project.
// Degrades to a plain figure when the target has no coordinates.
function distLink(miles, dir, lat, lon, label) {
  if (miles == null) return '<span style="color:#A5ADB7">—</span>';
  const txt = `${(+miles).toFixed(2)} mi${dir ? ' ' + esc(dir) : ''}`;
  const from = projectCentroid();
  if (lat == null || lon == null || !from) {
    return `<span style="color:#F25929;font-weight:600">${txt}</span>`;
  }
  const url = `https://www.google.com/maps/dir/?api=1&origin=${from.lat},${from.lon}`
            + `&destination=${lat},${lon}&travelmode=driving`;
  return `<a href="${url}" target="_blank" rel="noopener"
             title="Driving directions from the project${label ? ' to ' + esc(label) : ''}"
             style="color:#F25929;font-weight:600;text-decoration:none;border-bottom:1px dotted #F25929">${txt} ↗</a>`;
}

const map = L.map('map').setView([30.0, -95.8], 10);
L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
  { maxZoom: 19, attribution: 'Imagery © Esri' }).addTo(map);
L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
  { maxZoom: 19 }).addTo(map);

let projectData = null;
let tractsLayer = L.layerGroup().addTo(map);
let constraintLayerGroup = L.layerGroup().addTo(map);   // floodplain, wetlands, streams, etc

const CONSTRAINT_STYLES = {
  floodplain:   { color: '#0040aa', fillColor: '#3380ff', fillOpacity: 0.40, weight: 1,  label: 'Floodplain (100-yr)' },
  wetlands:     { color: '#005c2e', fillColor: '#33a06f', fillOpacity: 0.45, weight: 1,  label: 'Wetlands (NWI)' },
  transmission: { color: '#9c27b0', weight: 3, opacity: 0.85, dashArray: '4,3',         label: 'Transmission' },
  streams:      { color: '#1976d2', weight: 2, opacity: 0.85,                            label: 'Streams' },
  pipelines:    { color: '#F25929', weight: 2, opacity: 0.85,                            label: 'Pipelines' },
};

function renderConstraintGeoms(constraintGeoms) {
  constraintLayerGroup.clearLayers();
  const legend = [];
  for (const [key, fc] of Object.entries(constraintGeoms || {})) {
    if (!fc) continue;
    const style = CONSTRAINT_STYLES[key] || { color: '#333', weight: 1 };
    L.geoJSON(fc, { style }).addTo(constraintLayerGroup);
    legend.push(`<div style="display:flex;align-items:center;gap:6px;margin:2px 0"><span style="display:inline-block;width:12px;height:12px;background:${style.fillColor || style.color};border:1px solid ${style.color}"></span><span>${style.label}</span></div>`);
  }
  // Show/hide legend
  const legendEl = document.getElementById('map-legend');
  if (legend.length) {
    document.getElementById('map-legend-rows').innerHTML = legend.join('');
    legendEl.style.display = 'block';
  } else {
    legendEl.style.display = 'none';
  }
}

async function loadProject() {
  const r = await fetch(`/api/acq/projects/${PROJECT_ID}`);
  const d = await r.json();
  if (d.error) { alert('Failed to load: ' + d.error); return; }
  projectData = d.project;
  render();
}

function render() {
  const p = projectData;
  document.getElementById('proj-name').textContent = p.name;
  document.getElementById('proj-name-input').value = p.name;
  document.title = p.name + ' — Acquisitions GIS';

  // Restore saved net-out switches so a re-analysis repeats the same basis.
  const na = p.netout_assumptions || {};
  for (const [id, key] of [['no-flood', 'flood'], ['no-wetlands', 'wetlands'],
                           ['no-transmission', 'transmission'], ['no-streams', 'streams'],
                           ['no-pipelines', 'pipelines']]) {
    const el = document.getElementById(id);
    if (el && na[key] !== undefined) el.checked = na[key] !== false;
  }
  const _ip = document.getElementById('infra-pct');
  if (_ip && na.infrastructure_pct !== undefined) _ip.value = na.infrastructure_pct;

  // A measured constraint is only as good as the layer on the day. An
  // engineer's exhibit or a LOMR beats FEMA's polygon, so any figure typed
  // here is deducted exactly as given and the row is marked as stated.
  const no = p.netout_overrides || {};
  document.querySelectorAll('.netout-acres').forEach(el => {
    if (no[el.dataset.key] !== undefined && no[el.dataset.key] !== null) {
      el.value = no[el.dataset.key];
    }
  });

  // Yield mix — 4 product types with density + allocation %
  const ya = p.yield_assumptions || {};
  let lotTypes = (ya.lot_types && ya.lot_types.length) ? ya.lot_types : [
    {label: "40 FF", units_per_acre: 6.0, allocation_pct: 25},
    {label: "50 FF", units_per_acre: 5.0, allocation_pct: 25},
    {label: "60 FF", units_per_acre: 4.0, allocation_pct: 25},
    {label: "70 FF", units_per_acre: 3.5, allocation_pct: 25},
  ];
  // Auto-clean legacy labels from older projects — strip the trailing
  // "Small/Standard/Premium/Estate" descriptors that were in earlier defaults.
  lotTypes = lotTypes.map(lt => ({
    ...lt,
    label: (lt.label || '').replace(/^(\d+\s*FF)\s*(Small|Standard|Premium|Estate)\b.*$/i, '$1').trim(),
  }));
  renderLotTypeRows(lotTypes);
  window._lotTypes = lotTypes;
  updateImpliedLot();

  // Tract list
  const tlist = document.getElementById('tract-list');
  tlist.innerHTML = (p.tracts||[]).map(t => `
    <div class="tract-row">
      <span><b>${esc(t.owner_name || '(no owner)')}</b><br><span style="color:#6B7B8B">${esc(t.county||'')} · Prop ${esc(t.prop_id)}</span></span>
      <span style="white-space:nowrap">${fmt(t.acres)} ac
        <button class="acre-edit" title="Set the acreage you know to be correct"
                data-pid="${esc(t.prop_id)}" data-fips="${esc(t.county_fips || '')}"
                data-county="${esc(t.county || '')}" data-acres="${t.acres || ''}"
                style="background:none;border:0;color:#6B7B8B;cursor:pointer;font-size:11px;padding:0 2px">&#9998;</button>
      </span>
    </div>
  `).join('') || '<div class="placeholder">No tracts in this project.</div>';

  // Neither StratMap nor the appraisal district is the last word -- a survey
  // or a deed beats both, so the acreage is editable here and the override
  // applies everywhere the app reports that parcel.
  tlist.querySelectorAll('.acre-edit').forEach(btn => {
    btn.addEventListener('click', async () => {
      const pid = btn.dataset.pid;
      const cur = btn.dataset.acres;
      const val = prompt(`Acreage for Prop ${pid}`
        + ' (blank to clear an override and go back to the surveyed sources):', cur);
      if (val === null) return;
      const note = val.trim() ? (prompt('Source, for the record (optional):', '') || '') : '';
      try {
        const r = await fetch(`/api/acq/parcel/${encodeURIComponent(pid)}/acres`, {
          method: val.trim() ? 'POST' : 'DELETE',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            county_fips: btn.dataset.fips || COUNTY_FIPS[btn.dataset.county] || '',
            acres: val.trim() ? Number(val) : null, note }),
        });
        const d = await r.json();
        if (!r.ok || d.error) throw new Error(d.error || ('HTTP ' + r.status));
        alert('Saved. Re-run the analysis to apply it to the acreage and yield.');
        location.reload();
      } catch (e) {
        alert('Could not save the acreage: ' + e.message);
      }
    });
  });

  // Map: draw tracts
  tractsLayer.clearLayers();
  const bounds = L.latLngBounds([]);
  for (const t of (p.tracts||[])) {
    if (!t.geometry) continue;
    const layer = L.geoJSON(t.geometry, {
      style: { color: '#F25929', weight: 2, fillColor: '#F25929', fillOpacity: 0.20 }
    }).addTo(tractsLayer);
    layer.bindPopup(`<b>${esc(t.owner_name)}</b><br>${fmt(t.acres)} ac · ${esc(t.county)}`);
    bounds.extend(layer.getBounds());
  }
  if (bounds.isValid()) map.fitBounds(bounds.pad(0.2));

  // Render cached analysis if present
  if (p.analysis_cache) {
    renderAnalysis(p.analysis_cache);
    if (p.analysis_cache.constraint_geoms) renderConstraintGeoms(p.analysis_cache.constraint_geoms);

    // Then load the context cards. Opening a saved project used to show the
    // acreage ladder and nothing else — schools, CBAS, amenities, roads, news,
    // market and FRED only ever loaded as a side effect of pressing Run
    // Acquisition Analysis. So a project you had already analysed came back
    // looking half-empty, and the fix was to re-run a minute of GIS queries you
    // had already paid for.
    //
    // These are the same calls the analyse handler makes; each updates its own
    // card independently and handles its own failure.
    if (p.analysis_cache.union_geometry) loadElevation(p.analysis_cache.union_geometry);
    loadSchools();
    loadCbas();
    loadAmenities();
    loadRoads();
    loadNews();
    loadMarket();
    loadFred();
  }
}

// Add-more-tracts button: pulls latest results from the main map's last
// search via /api/last-search and lets user pick which tracts to add.
document.getElementById('btn-add-tracts').addEventListener('click', async () => {
  try {
    const r = await fetch('/api/acq/last-search');
    const d = await r.json();
    if (!d.ok || !d.layers || !d.layers.tracts || !d.layers.tracts.features) {
      alert("No recent search results found.\\n\\nGo to the Map page, run a search, " +
            "select tracts (shift+click on map or checkboxes), then come back and click 'Add more from map' " +
            "— OR use the bulk-bar '+ Add to existing project' from the search results page.");
      return;
    }
    const tracts = d.layers.tracts.features;
    const existing = new Set((projectData.tracts || []).map(t => t.prop_id));
    const available = tracts.filter(f => !existing.has(String((f.properties||{}).Prop_ID||'')));
    if (!available.length) {
      alert('All current search-result tracts are already in this project.');
      return;
    }
    // Build a checkbox list in a prompt-style modal (simple)
    const html = '<div style="max-height:50vh;overflow-y:auto;border:1px solid #DDE3E8;padding:8px;border-radius:6px">' +
      available.map((f, i) => {
        const x = f.properties || {};
        return `<label style="display:block;padding:4px 0;font-size:12px;border-bottom:1px solid #F0F0F0"><input type="checkbox" data-idx="${i}" checked style="margin-right:6px">${esc(x.OWNER_NAME || '?')} · ${fmt(x.Acres,1)} ac · ${esc(x._county || '?')} · Prop ${esc(x.Prop_ID)}</label>`;
      }).join('') + '</div>';
    // Quick modal
    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.4);z-index:10000;display:flex;align-items:center;justify-content:center';
    overlay.innerHTML = `<div style="background:#FFF;border-radius:8px;padding:20px;max-width:600px;width:90%;max-height:80vh;overflow:auto">
      <h3 style="margin:0 0 12px">Add tracts to "${esc(projectData.name)}"</h3>
      <p style="margin:0 0 8px;font-size:13px;color:#6B7B8B">${available.length} tracts available from the last search. Uncheck any you don't want to add.</p>
      ${html}
      <div style="margin-top:12px;display:flex;gap:8px;justify-content:flex-end">
        <button id="modal-cancel" class="ghost">Cancel</button>
        <button id="modal-add" class="primary">Add selected</button>
      </div>
    </div>`;
    document.body.appendChild(overlay);
    document.getElementById('modal-cancel').onclick = () => overlay.remove();
    document.getElementById('modal-add').onclick = async () => {
      const checked = Array.from(overlay.querySelectorAll('input[type=checkbox]:checked'))
                          .map(cb => available[parseInt(cb.dataset.idx)]);
      if (!checked.length) { overlay.remove(); return; }
      const payload = {
        tracts: checked.map(f => ({
          prop_id:    String(f.properties.Prop_ID || ''),
          owner_name: f.properties.OWNER_NAME || '',
          acres:      f.properties.Acres || 0,
          county:     f.properties._county || '',
          geometry:   f.geometry,
        })),
      };
      overlay.remove();
      const r2 = await fetch(`/api/acq/projects/${PROJECT_ID}/add-tracts`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const d2 = await r2.json();
      if (!r2.ok || d2.error) { alert('Add failed: ' + (d2.error || r2.status)); return; }
      alert(`Added ${d2.added} tract${d2.added===1?'':'s'} (${d2.skipped} duplicates skipped).\\nRe-running analysis…`);
      projectData = d2.project;
      render();
      // Auto re-run analysis since geometry changed
      document.getElementById('btn-analyze').click();
    };
  } catch (e) {
    alert('Could not load last-search results: ' + e.message);
  }
});

function renderAnalysis(a) {
  // Reached from both the cached-analysis path on load and the fresh run, so
  // the button's state follows the analysis wherever it came from.
  _syncSummaryBtn(!!a);
  const c = a.constraints || {};
  const body = document.getElementById('analysis-body');
  memoPut('analysis', a);

  let constraintsHtml = '';
  const rows = [
    ['Floodplain (100-yr)', c.floodplain],
    ['Wetlands (NWI)',      c.wetlands],
    ['Transmission ROW',    c.transmission_row],
    ['Stream buffers (50ft)', c.stream_buffers],
    ['Pipeline easements (50ft)', c.pipeline_easements],
  ];
  let anyConstraintFailed = false;
  for (const [name, x] of rows) {
    if (!x) continue;
    if (x.error) {
      anyConstraintFailed = true;
      // Make a failed layer unmistakable — a silent 0 would read as "no
      // floodplain here" and quietly inflate net developable acreage.
      constraintsHtml += `<div class="constraint-row" style="background:#FEF2F2">
          <span>${esc(name)} <span style="background:#C62828;color:#FFF;padding:1px 6px;border-radius:3px;font-size:9px;font-weight:700">DATA UNAVAILABLE</span></span>
          <span class="err" style="font-size:10px" title="${esc(x.error)}">not counted in net developable</span>
        </div>`;
      continue;
    }
    constraintsHtml += `<div class="constraint-row"><span>${esc(name)}</span><span class="acres">${fmt(x.acres)} ac · ${fmt(x.pct,1)}%</span></div>`;
  }
  if (anyConstraintFailed) {
    constraintsHtml += `<div style="background:#FEF2F2;border-left:3px solid #C62828;padding:8px 10px;margin-top:8px;font-size:11px;border-radius:4px">
        <b>One or more constraint layers failed to load.</b>Net developable acreage is
        <b>overstated</b> — the missing layers were not deducted. Re-run the analysis; these
        public services (FEMA NFHL especially) intermittently drop connections under load.
      </div>`;
  }

  // Gross -> constraints -> infrastructure -> saleable, shown as a ladder so the
  // number the yield is built on is traceable rather than asserted.
  // A layer that failed to load and a layer you deliberately kept both show as
  // "not deducted", but they mean opposite things - one overstates the acreage
  // without you knowing. Call the failure out in red.
  // The deducted figure is what the layer removes BEYOND the layers above it,
  // not its own footprint. A wetland inside the floodplain is real acreage on
  // both layers, but it only leaves the site once. Showing both footprints
  // under a minus sign made the column read as double-dipping and it never
  // summed to the net developable figure underneath it. Net developable itself
  // was always right -- the constraints are unioned before the difference.
  const ladder = (a.netout_detail || []).map(n => {
    const marg = (n.acres_marginal == null) ? n.acres : n.acres_marginal;
    const overlap = (+n.acres || 0) - (+marg || 0);
    const shown = n.applied ? marg : n.acres;
    const note = n.error
      ? ` <span style="font-size:10px;color:#C62828;font-weight:700" title="${esc(n.reason || 'the service did not respond')}">layer unavailable — not deducted</span>`
      : (n.stated
          ? ` <span style="font-size:10px;color:#C08A2E;font-weight:700">stated</span>`
            + ` <span style="font-size:10px;color:#6B7B8B">layers measure ${fmt(n.measured_acres)} ac</span>`
          : n.applied
            ? (overlap > 0.05
                ? ` <span style="font-size:10px;color:#6B7B8B">${fmt(n.acres)} ac on site · ${fmt(overlap)} ac already deducted above</span>`
                : '')
            : ' <span style="font-size:10px;color:#6B7B8B">(kept by override)</span>');
    return `
      <div class="constraint-row" style="${n.applied ? '' : 'opacity:.65'}">
        <span>${esc(n.label)}${note}</span>
        <span class="acres">${n.applied ? '-' : ''}${n.error ? '?' : fmt(shown)} ac</span>
      </div>`;
  }).join('');

  // Say where the headline acreage came from. It is deliberately NOT a fresh
  // measurement of the polygon -- the boundary a county draws and the acreage
  // it states for the same account differ by a few tenths, and the analysis
  // has to agree with the tract listed above it, not with the map.
  const basisNote = {
    override: 'per your override',
    stated:   'per appraisal district',
    measured: 'measured from the boundary',
  }[a.gross_basis] || '';
  const grossNote = basisNote
    ? `<div style="font-size:10px;color:#6B7B8B;margin-top:2px">${esc(basisNote)}${
        (a.gross_basis !== 'measured' && a.measured_acres) ?
          ` &middot; boundary measures ${fmt(a.measured_acres)} ac` : ''}</div>`
    : '';

  body.innerHTML = `
    <div class="kpi-grid">
      <div class="kpi big"><div class="label">Gross</div><div class="value">${fmt(a.gross_acres)} ac</div>${grossNote}</div>
      <div class="kpi big"><div class="label">Net developable</div><div class="value">${fmt(a.net_developable_acres)} ac</div></div>
      <div class="kpi big" style="background:#FFF4ED;border:1px solid #F25929">
        <div class="label">Net saleable</div>
        <div class="value" style="color:#F25929">${fmt(a.net_saleable_acres)} ac</div>
      </div>
      <div class="kpi"><div class="label">Saleable %</div><div class="value">${fmt(a.net_saleable_pct,1)}%</div></div>
    </div>

    <h3>Gross to saleable</h3>
    <div class="constraint-row" style="font-weight:600"><span>Gross acreage</span><span class="acres">${fmt(a.gross_acres)} ac</span></div>
    ${ladder}
    <div class="constraint-row" style="font-weight:600;border-top:1px solid #DDE3E8">
      <span>Net developable</span><span class="acres">${fmt(a.net_developable_acres)} ac</span></div>
    <div class="constraint-row">
      <span>Infrastructure &amp; landscaping (${fmt(a.infrastructure_pct,0)}%)
        <span style="font-size:10px;color:#6B7B8B">roads, plans, detention, landscaping</span></span>
      <span class="acres">-${fmt(a.infrastructure_acres)} ac</span></div>
    <div class="constraint-row" style="font-weight:700;border-top:2px solid #F25929;color:#F25929">
      <span>Net saleable</span><span class="acres">${fmt(a.net_saleable_acres)} ac</span></div>
    <div style="font-size:11px;color:#6B7B8B;margin-top:6px">
      Each layer shows what it removes beyond the layers above it, so the column adds up
      to net developable. Where layers overlap — a wetland inside the floodplain — the
      shared acreage is deducted once and charged to the layer listed first; the note on
      each row gives that layer's full footprint on the site.
      Yield below runs off net saleable acreage.
    </div>

    <h3>Constraint breakdown</h3>
    ${constraintsHtml}

    <h3>Yield (pro-forma)</h3>
    <div class="kpi-grid" style="grid-template-columns:2fr 1fr 1fr">
      <div class="kpi" style="background:#FFF4ED;border:1px solid #F25929">
        <div class="label">TOTAL LOTS</div>
        <div class="value" style="font-size:32px;color:#F25929">${fmt(a.yield_estimates.total_lots, 0)}</div>
        <div class="label" style="margin-top:2px">${fmt(a.net_developable_acres)} net dev ac · weighted ${a.yield_estimates.weighted_density} u/ac</div>
      </div>
      <div class="kpi"><div class="label">Net developable</div><div class="value">${fmt(a.net_developable_acres)} ac</div></div>
      <div class="kpi"><div class="label">Weighted density</div><div class="value">${a.yield_estimates.weighted_density} u/ac</div></div>
    </div>

    <h3>Mix breakdown</h3>
    <table style="width:100%;border-collapse:collapse;font-size:12px">
      <thead>
        <tr style="background:#13344E;color:#FFF">
          <th style="text-align:left;padding:6px 8px">Product</th>
          <th style="text-align:right;padding:6px 8px">Density</th>
          <th style="text-align:right;padding:6px 8px">Allocation</th>
          <th style="text-align:right;padding:6px 8px">Acres</th>
          <th style="text-align:right;padding:6px 8px">Lots</th>
        </tr>
      </thead>
      <tbody>
        ${(a.yield_estimates.breakdown || []).map(b => `
          <tr style="border-bottom:1px solid #DDE3E8">
            <td style="padding:5px 8px">${esc(b.label)}</td>
            <td style="padding:5px 8px;text-align:right">${b.units_per_acre} u/ac</td>
            <td style="padding:5px 8px;text-align:right">${fmt(b.allocation_pct,0)}%</td>
            <td style="padding:5px 8px;text-align:right">${fmt(b.acres,1)}</td>
            <td style="padding:5px 8px;text-align:right"><b>${fmt(b.lots,0)}</b></td>
          </tr>
        `).join('')}
        <tr style="background:#FFF4ED;font-weight:600">
          <td style="padding:6px 8px">Total</td>
          <td style="padding:6px 8px"></td>
          <td style="padding:6px 8px;text-align:right">${fmt(a.yield_estimates.total_allocation_pct, 0)}%</td>
          <td style="padding:6px 8px;text-align:right">${fmt(a.net_developable_acres)}</td>
          <td style="padding:6px 8px;text-align:right">${fmt(a.yield_estimates.total_lots, 0)}</td>
        </tr>
      </tbody>
    </table>

    <div style="font-size:11px;color:#6B7B8B;margin-top:8px">Computed ${esc(a.computed_at)}. Re-run if you change tract list or assumptions.</div>
  `;

}

document.getElementById('btn-rename').addEventListener('click', async () => {
  const name = document.getElementById('proj-name-input').value.trim();
  if (!name) return;
  const r = await fetch(`/api/acq/projects/${PROJECT_ID}`, {
    method: 'PATCH', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name }),
  });
  const d = await r.json();
  if (d.ok) { projectData = d.project; render(); }
});

document.getElementById('btn-delete').addEventListener('click', async () => {
  if (!confirm('Delete this project? The underlying tracts remain.')) return;
  const r = await fetch(`/api/acq/projects/${PROJECT_ID}`, { method: 'DELETE' });
  if (r.ok) { window.location.href = '/'; }
});

function renderLotTypeRows(lotTypes) {
  const rowsEl = document.getElementById('lot-types-rows');
  rowsEl.innerHTML = lotTypes.map((lt, i) => `
    <div style="display:grid;grid-template-columns:1.6fr 1fr 1fr;gap:4px;margin:2px 0">
      <input autocomplete="off" data-lpignore="true" data-1p-ignore type="text" data-lt-idx="${i}" data-lt-field="label" value="${(lt.label||'').replace(/"/g,'&quot;')}" style="padding:3px 6px">
      <input autocomplete="off" data-lpignore="true" data-1p-ignore type="number" data-lt-idx="${i}" data-lt-field="units_per_acre" value="${lt.units_per_acre||0}" step="0.1" min="0" max="50" style="padding:3px 6px;text-align:right">
      <input autocomplete="off" data-lpignore="true" data-1p-ignore type="number" data-lt-idx="${i}" data-lt-field="allocation_pct" value="${lt.allocation_pct||0}" step="5" min="0" max="100" style="padding:3px 6px;text-align:right">
    </div>
  `).join('');
  rowsEl.querySelectorAll('input').forEach(inp => {
    inp.addEventListener('input', (e) => {
      const i = parseInt(e.target.dataset.ltIdx, 10);
      const field = e.target.dataset.ltField;
      const v = field === 'label' ? e.target.value : (parseFloat(e.target.value) || 0);
      window._lotTypes[i][field] = v;
      updateImpliedLot();
    });
  });
  updateImpliedLot();
}

function updateImpliedLot() {
  const lots = window._lotTypes || [];
  const totalAlloc = lots.reduce((s, x) => s + (x.allocation_pct || 0), 0);
  const weighted = lots.reduce((s, x) => s + ((x.allocation_pct || 0) / 100) * (x.units_per_acre || 0), 0);
  document.getElementById('lot-alloc-total').textContent = totalAlloc.toFixed(0) + '%';
  document.getElementById('lot-alloc-total').style.color = (totalAlloc === 100) ? '#2E7D32' : '#C62828';
  document.getElementById('yield-implied').innerHTML =
    `Weighted density: <b>${weighted.toFixed(2)} u/ac</b> across the mix. ` +
    (totalAlloc !== 100 ? `<span style="color:#C62828">Allocation totals ${totalAlloc.toFixed(0)}% — should equal 100%</span>` : '');
}

function _readNetoutOverrides() {
  const out = {};
  document.querySelectorAll('.netout-acres').forEach(el => {
    const v = (el.value || '').trim().replace(/,/g, '');
    if (v === '') return;
    const n = Number(v);
    if (isFinite(n) && n >= 0) out[el.dataset.key] = n;
  });
  return out;
}

function _readNetouts() {
  const g = (id, d) => { const el = document.getElementById(id); return el ? el.checked : d; };
  const pctEl = document.getElementById('infra-pct');
  let pct = pctEl ? parseFloat(pctEl.value) : 30;
  if (!isFinite(pct)) pct = 30;
  return {
    flood:        g('no-flood', true),
    wetlands:     g('no-wetlands', true),
    transmission: g('no-transmission', true),
    streams:      g('no-streams', true),
    pipelines:    g('no-pipelines', true),
    infrastructure_pct: Math.max(0, Math.min(pct, 90)),
  };
}


// The summary report is built from the cached analysis, so it only becomes
// available once one has run. Enabling it before that just produces a 400.
function _syncSummaryBtn(hasAnalysis) {
  const b = document.getElementById('btn-exec-pdf');
  if (!b) return;
  b.disabled = !hasAnalysis;
  b.title = hasAnalysis
    ? 'Ten-page report: executive summary, constraint map, yield, topography, submarket and tract roster'
    : 'Run the acquisition analysis first — the report is built from it';
}

// Executive summary. Fetched rather than window.open'd so the button can show
// a real progress state -- the report draws a satellite map and several charts
// and takes a few seconds -- and so a failure surfaces as a readable message
// instead of a browser tab containing raw JSON.
document.getElementById('btn-exec-pdf')?.addEventListener('click', async (ev) => {
  const btn = ev.currentTarget;
  const label = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Building report…';
  try {
    const r = await fetch(`/api/acq/projects/${encodeURIComponent(PROJECT_ID)}/executive-pdf`);
    if (!r.ok) {
      let msg = `HTTP ${r.status}`;
      try { msg = (await r.json()).error || msg; } catch (e) {}
      throw new Error(msg);
    }
    const blob = await r.blob();
    const cd = r.headers.get('content-disposition') || '';
    const m = /filename="?([^"]+)"?/.exec(cd);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = m ? m[1] : 'EMBER_Acquisition_Summary.pdf';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  } catch (e) {
    console.error('[executive pdf]', e);
    alert('Could not build the executive report: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.textContent = label;
  }
});

document.getElementById('btn-analyze').addEventListener('click', async () => {
  await fetch(`/api/acq/projects/${PROJECT_ID}`, {
    method: 'PATCH', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ yield_assumptions: { lot_types: window._lotTypes },
                           netout_assumptions: _readNetouts(),
                           netout_overrides: _readNetoutOverrides() }),
  });
  const statusEl = document.getElementById('analyze-status');
  const btn = document.getElementById('btn-analyze');
  btn.disabled = true; btn.textContent = '⏳ Analyzing…';
  statusEl.textContent = 'Computing constraints + yield estimates (~10-30 sec)…';
  try {
    const r = await fetch(`/api/acq/projects/${PROJECT_ID}/analyze`, { method: 'POST' });
    const d = await r.json();
    if (!r.ok || d.error) throw new Error(d.error || `HTTP ${r.status}`);
    renderAnalysis(d.analysis);
    if (d.analysis.constraint_geoms) renderConstraintGeoms(d.analysis.constraint_geoms);
    statusEl.textContent = 'Done. Loading elevation + submarket + communities + amenities + builders + news + market…';
    // Kick off all data loads in parallel — they update their cards independently
    // CBAS supplies communities + builders (real market data), so the old
    // parcel-derived versions of those cards are gone.
    loadElevation(d.analysis.union_geometry);
    loadSchools();
    loadCbas();
    loadAmenities();
    loadRoads();
    loadNews();
    loadMarket();
    loadFred();
  } catch (e) {
    statusEl.innerHTML = `<span class="err">Analysis failed: ${esc(e.message)}</span>`;
  } finally {
    btn.disabled = false;
    btn.textContent = 'Re-run Analysis';
  }
});

let elevContourLayer = null;   // drawn contour lines on the map
async function loadElevation(geom) {
  const card = document.getElementById('elev-card');
  const body = document.getElementById('elev-body');
  card.style.display = 'block';
  body.innerHTML = '<div class="placeholder">Sampling USGS 3DEP elevation across the project (~30 sec)…</div>';
  try {
    const r = await fetch('/api/acq/elevation-profile', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ geometry: geom }),
    });
    const d = await r.json();
    if (!r.ok || d.error) throw new Error(d.error || `HTTP ${r.status}`);

    // Render contour lines on the project map (semi-transparent gray, varying by level)
    if (elevContourLayer) { try { map.removeLayer(elevContourLayer); } catch {} }
    elevContourLayer = L.layerGroup().addTo(map);
    if (d.contours && Array.isArray(d.contours)) {
      const minE = d.min_ft, maxE = d.max_ft, range = maxE - minE || 1;
      for (const c of d.contours) {
        // Color: shade from light blue (low) to bright red (high)
        const t = (c.level_ft - minE) / range;
        const r2 = Math.round(60 + t * 195), g = Math.round(120 - t * 80), b = Math.round(200 - t * 180);
        const color = `rgb(${r2},${g},${b})`;
        const lines = c.lines || [];
        for (const line of lines) {
          L.polyline(line.map(([x, y]) => [y, x]), {
            color, weight: 1.2, opacity: 0.75, interactive: false
          }).addTo(elevContourLayer);
        }
      }
    }

    // Mark highest + lowest points
    if (d.highest_latlon) {
      L.marker(d.highest_latlon, {
        icon: L.divIcon({ className: '', html: `<div style="background:#C62828;color:#FFF;font-size:10px;padding:2px 5px;border-radius:3px;white-space:nowrap;box-shadow:0 1px 3px rgba(0,0,0,0.3)">▲ HIGH ${fmt(d.highest_ft,0)}ft</div>`, iconAnchor: [0, 10] })
      }).addTo(elevContourLayer);
    }
    if (d.lowest_latlon) {
      L.marker(d.lowest_latlon, {
        icon: L.divIcon({ className: '', html: `<div style="background:#1976D2;color:#FFF;font-size:10px;padding:2px 5px;border-radius:3px;white-space:nowrap;box-shadow:0 1px 3px rgba(0,0,0,0.3)">▼ LOW ${fmt(d.lowest_ft,0)}ft</div>`, iconAnchor: [0, 10] })
      }).addTo(elevContourLayer);
    }

    const slope = d.slope_mean_pct ?? 0;
    const slopeLabel = slope < 0.5 ? 'very flat (engineered grading needed)'
                      : slope < 2.0 ? 'gentle (standard grading)'
                      : slope < 5.0 ? 'moderate (engineered grading)'
                      : 'significant relief';
    body.innerHTML = `
      <div class="kpi-grid">
        <div class="kpi"><div class="label">Elevation range</div><div class="value">${fmt(d.range_ft)} ft</div></div>
        <div class="kpi"><div class="label">Min / Max</div><div class="value" style="font-size:14px">${fmt(d.min_ft)} – ${fmt(d.max_ft)} ft</div></div>
        <div class="kpi"><div class="label">Mean slope</div><div class="value">${fmt(slope, 2)}%</div></div>
        <div class="kpi"><div class="label">Max slope</div><div class="value">${fmt(d.slope_max_pct ?? 0, 2)}%</div></div>
        <div class="kpi"><div class="label">Drainage</div><div class="value" style="font-size:13px">${esc(d.drainage_dir || '—')}</div></div>
        <div class="kpi"><div class="label">Median elev</div><div class="value">${fmt(d.median_ft)} ft</div></div>
      </div>
      <div style="background:#FFF4ED;border-left:3px solid #F25929;padding:8px 12px;margin-top:8px;border-radius:4px;font-size:13px">
        <b>Site character:</b> ${esc(slopeLabel)}. Contour interval ${d.contour_interval_ft}ft.
      </div>
      <div style="font-size:11px;color:#6B7B8B;margin-top:8px">Sampled ${d.sample_count || '?'} points via USGS 3DEP. Contours + high/low markers now visible on the project map above. Toggle them off by reloading the analysis.</div>
    `;
  } catch (e) {
    body.innerHTML = `<div class="placeholder err">Elevation failed: ${esc(e.message)}</div>`;
  }
}

async function loadMarket() {
  const card = document.getElementById('market-card');
  const body = document.getElementById('market-body');
  card.style.display = 'block';
  body.innerHTML = '<div class="placeholder">Pulling Census ACS county-level demographics…</div>';
  try {
    const r = await fetch(`/api/acq/projects/${PROJECT_ID}/market`);
    const d = await r.json();
    if (!r.ok || d.error) throw new Error(d.error || `HTTP ${r.status}`);
    const cur = d.current || {};
    const gro = d.growth || {};
    function dollar(v) { return v != null ? '$' + fmt(v, 0) : '—'; }
    function num(v) { return v != null ? fmt(v, 0) : '—'; }
    function pct(v) { return v != null ? fmt(v, 1) + '%' : '—'; }
    function trend(t) { return t != null ? ` <span style="color:${t>=0?'#2E7D32':'#C62828'};font-size:10px">${t>=0?'▲':'▼'} ${Math.abs(t).toFixed(1)}% 5yr</span>` : ''; }
    body.innerHTML = `
      <div style="font-size:12px;color:#6B7B8B;margin-bottom:8px">${esc(d.county_name)} County · Census ACS ${cur.year}</div>

      <h3>Population & households</h3>
      <div class="kpi-grid">
        <div class="kpi"><div class="label">Population</div><div class="value">${num(cur.population)}</div><div class="label" style="margin-top:2px">${trend(gro.population_total_pct)} ${gro.population_cagr_pct!=null?'('+fmt(gro.population_cagr_pct,1)+'% CAGR)':''}</div></div>
        <div class="kpi"><div class="label">Median age</div><div class="value">${cur.median_age != null ? cur.median_age.toFixed(1) : '—'}</div></div>
        <div class="kpi"><div class="label">Households</div><div class="value">${num(cur.households)}</div></div>
        <div class="kpi"><div class="label">Avg household size</div><div class="value">${cur.avg_household_size != null ? cur.avg_household_size.toFixed(2) : '—'}</div></div>
      </div>

      <h3>Income</h3>
      <div class="kpi-grid">
        <div class="kpi"><div class="label">Median household income</div><div class="value">${dollar(cur.median_household_income)}</div></div>
        <div class="kpi"><div class="label">Per capita income</div><div class="value">${dollar(cur.per_capita_income)}</div></div>
      </div>

      <h3>Housing</h3>
      <div class="kpi-grid">
        <div class="kpi"><div class="label">Median home value</div><div class="value">${dollar(cur.median_home_value)}</div><div class="label" style="margin-top:2px">${trend(gro.home_value_total_pct)} ${gro.home_value_cagr_pct!=null?'('+fmt(gro.home_value_cagr_pct,1)+'% CAGR)':''}</div></div>
        <div class="kpi"><div class="label">Median rent</div><div class="value">${dollar(cur.median_rent)}</div></div>
        <div class="kpi"><div class="label">Housing units</div><div class="value">${num(cur.housing_units)}</div></div>
        <div class="kpi"><div class="label">Owner-occupancy</div><div class="value">${pct(cur.owner_occupancy_pct)}</div></div>
        <div class="kpi"><div class="label">Vacancy rate</div><div class="value">${pct(cur.vacancy_pct)}</div></div>
        <div class="kpi"><div class="label">Owner-occupied units</div><div class="value">${num(cur.owner_occupied)}</div></div>
      </div>

      <h3>Workforce & education</h3>
      <div class="kpi-grid">
        <div class="kpi"><div class="label">Employed</div><div class="value">${num(cur.employed)}</div><div class="label" style="margin-top:2px">${trend(gro.employed_total_pct)}</div></div>
        <div class="kpi"><div class="label">% Bachelor's or higher (25+)</div><div class="value">${pct(cur.pct_bachelors_or_higher)}</div></div>
      </div>

      <div style="font-size:11px;color:#6B7B8B;margin-top:8px">Sources: ${esc((d.sources || []).join(' · '))}.</div>
    `;
  } catch (e) {
    body.innerHTML = `<div class="placeholder err">Market data failed: ${esc(e.message)}</div>`;
  }
}


// Unfinished sections stay off until their assumptions are settled.
const SHOW_DEMAND_SECTIONS = false;
const SHOW_ACQUISITION_MEMO = false;

async function loadAmenities() {
  const card = document.getElementById('amenities-card');
  const body = document.getElementById('amenities-body');
  card.style.display = 'block';
  body.innerHTML = '<div class="placeholder">Querying OpenStreetMap for nearby grocery, schools, hospitals, parks…</div>';
  try {
    const r = await fetch(`/api/acq/projects/${PROJECT_ID}/amenities?radius_mi=5`);
    const d = await r.json();
    if (!r.ok || d.error) {
      body.innerHTML = `<div class="placeholder err">Amenities lookup failed: ${esc(d.error || ('HTTP ' + r.status))}</div>`;
      return;
    }
    function renderList(title, items) {
      items = (items || []).filter(x => {
        const nm = String(x.name || '').trim().toLowerCase();
        return nm && nm !== '(unnamed)' && nm !== 'unnamed' && nm !== 'unknown';
      });
      if (!items || !items.length) {
        return `<div style="margin-bottom:12px"><h3>${esc(title)}</h3>
                <div style="font-size:11px;color:#A5ADB7">None found.</div></div>`;
      }
      return `
        <div style="margin-bottom:14px">
          <h3>${esc(title)}</h3>
          <div style="font-size:11px;max-height:200px;overflow-y:auto;border:1px solid #DDE3E8;border-radius:4px">
            ${items.map(x => {
              const nm = esc(String(x.name || '').trim());
              // OSM tags brand alongside name and they're usually identical
              // ("H-E-B (H-E-B)"); only show brand when it says something new.
              const b = (x.brand || '').trim();
              const brandLine = (b && b.toLowerCase() !== (x.name || '').toLowerCase())
                ? ` <span style="color:#6B7B8B">${esc(b)}</span>` : '';
              return `
                <div style="padding:5px 8px;border-bottom:1px solid #F0F0F0;display:flex;justify-content:space-between;align-items:center">
                  <span><b>${nm}</b>${brandLine}</span>
                  <span>${distLink(x.distance_mi, x.direction, x.lat, x.lon, x.name)}</span>
                </div>
              `;
            }).join('')}
          </div>
        </div>
      `;
    }
    // Commute & access leads the card: for an outer-ring tract, distance to the
    // freeway moves lot pricing more than any store does, so it shouldn't sit
    // below a list of parks.
    function accessStrip() {
      const hw = d.highways || [];
      if (!hw.length && d.nearest_ramp_mi == null) return '';
      const primary = hw[0];
      const tile = (label, value, sub) => `
        <div class="kpi">
          <div class="label">${label}</div>
          <div class="value" style="font-size:18px">${value}</div>
          ${sub ? `<div class="label" style="margin-top:2px">${sub}</div>` : ''}
        </div>`;
      return `
        <div style="margin-bottom:14px">
          <h3>Commute &amp; access</h3>
          <div class="kpi-grid" style="grid-template-columns:repeat(3,1fr)">
            ${primary ? tile('Nearest freeway',
                `<span style="color:#F25929">${esc(primary.ref)}</span>`,
                `${primary.distance_mi} mi ${esc(primary.direction)}${primary.name ? ' · ' + esc(primary.name) : ''}`) : ''}
            ${d.nearest_ramp_mi != null ? tile('Nearest on-ramp', d.nearest_ramp_mi + ' mi', 'closest interchange') : ''}
            ${(d.airports && d.airports.length) ? tile('Airport',
                esc(d.airports[0].iata || '—'),
                `${d.airports[0].distance_mi} mi · ${esc(d.airports[0].name)}`) : ''}
          </div>
          ${hw.length ? `
            <div style="font-size:11px;max-height:150px;overflow-y:auto;border:1px solid #DDE3E8;border-radius:4px;margin-top:8px">
              ${hw.map(h => `
                <div style="padding:5px 8px;border-bottom:1px solid #F0F0F0;display:flex;justify-content:space-between;align-items:center">
                  <span><b>${esc(h.ref)}</b>${h.name ? ` <span style="color:#6B7B8B">${esc(h.name)}</span>` : ''}
                    ${h.toll ? '<span style="background:#F59E0B;color:#FFF;padding:1px 5px;border-radius:3px;font-size:9px;font-weight:700;margin-left:4px">TOLL</span>' : ''}</span>
                  <span>${distLink(h.distance_mi, h.direction, h.lat, h.lon, h.ref)}</span>
                </div>`).join('')}
            </div>` : ''}
          ${(d.park_ride && d.park_ride.length) ? `
            <div style="font-size:11px;color:#6B7B8B;margin-top:4px">
              Nearest park &amp; ride: <b>${esc(d.park_ride[0].name)}</b> ${distLink(d.park_ride[0].distance_mi, d.park_ride[0].direction, d.park_ride[0].lat, d.park_ride[0].lon, d.park_ride[0].name)}
            </div>` : ''}
        </div>`;
    }

    // At full buildout: what the trade area becomes, and which commercial uses
    // that population would support that aren't here yet.
    function buildoutStrip() {
      const b = d.buildout || {};
      if (!b.population_now && !b.rooftops_at_buildout) return '';
      const dm = b.demographics || {};
      return `
        <div style="margin-bottom:14px">
          <h3>At full buildout</h3>
          <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr)">
            <div class="kpi" style="background:#FFF4ED;border:1px solid #F25929">
              <div class="label">Population at buildout</div>
              <div class="value" style="color:#F25929">${fmt(b.population_at_buildout,0)}</div>
              <div class="label" style="margin-top:2px">${fmt(b.population_now,0)} today${b.growth_pct != null ? ` · +${b.growth_pct}%` : ''}</div>
            </div>
            <div class="kpi">
              <div class="label">Rooftops at buildout</div>
              <div class="value">${fmt(b.rooftops_at_buildout,0)}</div>
              <div class="label" style="margin-top:2px">${fmt(b.rooftops_now,0)} today · ${fmt(b.lots_remaining,0)} to build</div>
            </div>
            <div class="kpi">
              <div class="label">Years to buildout</div>
              <div class="value">${b.years_to_buildout != null ? b.years_to_buildout : '—'}</div>
              <div class="label" style="margin-top:2px">at ${fmt(b.annual_closings,0)} closings/yr</div>
            </div>
            <div class="kpi">
              <div class="label">Avg household</div>
              <div class="value">${b.avg_household_size != null ? b.avg_household_size : '—'}</div>
              <div class="label" style="margin-top:2px">${esc(b.household_size_source || '')}</div>
            </div>
          </div>
          ${(dm.median_household_income || dm.median_home_value) ? `
            <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr);margin-top:6px">
              <div class="kpi"><div class="label">Median HH income</div><div class="value">$${fmt(dm.median_household_income,0)}</div></div>
              <div class="kpi"><div class="label">Median home value</div><div class="value">$${fmt(dm.median_home_value,0)}</div></div>
              <div class="kpi"><div class="label">Median age</div><div class="value">${dm.median_age != null ? dm.median_age : '—'}</div></div>
              <div class="kpi"><div class="label">Households today</div><div class="value">${fmt(dm.households,0)}</div></div>
            </div>` : ''}
          <div style="font-size:11px;color:#6B7B8B;margin-top:6px">
            Rooftops today are ${esc(b.rooftops_source || 'n/a')} within ${d.radius_mi} mi
            ${b.cbas_occupied ? `(CBAS tracks ${fmt(b.cbas_occupied,0)} of these in actively-selling communities)` : ''};
            growth is the remaining CBAS lot pipeline. Population from ${esc(b.population_basis || 'n/a')}${dm.tracts ? ` across ${dm.tracts} tracts, ACS ${dm.year}` : ''}
            at ${b.avg_household_size} per home.
          </div>
        </div>`;
    }

    // Commercial gap: what this population supports vs what exists.
    function opportunityBlock() {
      const ops = d.opportunities || [];
      if (!ops.length) return '';
      const real = ops.filter(o => o.opportunity && !o.unknown);
      return `
        <div style="margin-bottom:14px;border:1px solid #13344E;border-radius:6px;padding:12px">
          <h3 style="margin-top:0">Commercial gaps &amp; opportunities</h3>
          ${real.length ? `
            <div style="font-size:12px;margin-bottom:10px;line-height:1.6">
              At full buildout this trade area is short:
              ${real.slice(0,4).map(o =>
                `<b>${esc(o.use.toLowerCase())}</b> (supports ${o.supported_at_buildout}, has ${o.existing})`).join(' · ')}.
            </div>` : `
            <div style="font-size:12px;margin-bottom:10px">
              No material retail shortfall at buildout — the existing commercial base covers projected demand.
            </div>`}
          <table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead>
              <tr style="background:#13344E;color:#FFF">
                <th style="text-align:left;padding:6px 8px">Use</th>
                <th style="text-align:right;padding:6px 8px">Within ${d.radius_mi} mi</th>
                <th style="text-align:right;padding:6px 8px">Supported today</th>
                <th style="text-align:right;padding:6px 8px">Supported at buildout</th>
                <th style="text-align:right;padding:6px 8px">Gap</th>
                <th style="text-align:center;padding:6px 8px" title="How consistent the ratio was across the six reference areas">Consistency</th>
                <th style="text-align:left;padding:6px 8px">Measured benchmark</th>
              </tr>
            </thead>
            <tbody>
              ${ops.map(o => {
                const cc = o.confidence === 'high' ? '#2E7D32'
                         : o.confidence === 'medium' ? '#F59E0B' : '#A5ADB7';
                return `
                <tr style="border-bottom:1px solid #DDE3E8;${o.opportunity && !o.unknown ? 'background:#FFF9F5' : ''}">
                  <td style="padding:5px 8px"><b>${esc(o.use)}</b></td>
                  <td style="padding:5px 8px;text-align:right">${o.unknown ? '<span style="color:#A5ADB7">partial</span>' : fmt(o.existing,0)}</td>
                  <td style="padding:5px 8px;text-align:right">${o.supported_now}</td>
                  <td style="padding:5px 8px;text-align:right"><b>${o.supported_at_buildout}</b></td>
                  <td style="padding:5px 8px;text-align:right">
                    ${o.gap_at_buildout > 0
                      ? `<b style="color:${o.confidence === 'low' ? '#A5ADB7' : '#F25929'}">+${o.gap_at_buildout}</b>`
                      : `<span style="color:#2E7D32">${o.gap_at_buildout}</span>`}
                  </td>
                  <td style="padding:5px 8px;text-align:center">
                    <span style="color:${cc};font-size:10px;font-weight:700">${esc(String(o.confidence).toUpperCase())}</span>
                  </td>
                  <td style="padding:5px 8px;font-size:11px;color:#6B7B8B">
                    1 per ${fmt(o.rooftops_per_facility,0)} rooftops · ${o.spread}× spread · ${esc(o.note)}
                  </td>
                </tr>`;
              }).join('')}
            </tbody>
          </table>
          <div style="font-size:11px;color:#6B7B8B;margin-top:6px">
            Benchmarks are measured, not assumed: six built-out Houston suburbs (Cinco Ranch, Bridgeland, Riverstone,
            Spring/Klein, Shadow Creek, Aliana), counting the same OSM categories within 5 mi against the OSM building
            footprint count in that submarket, taking the median. <b>Spread</b> is max ÷ min across those six — low spread
            means the ratio is a real requirement, high spread means it varies too much to act on, which is why
            hospitals carry no gap flag. Rooftops here are building footprints, the same denominator used above.
          </div>
        </div>`;
    }

    // NOTE: schools intentionally omitted here — the School District Profile
    // card covers them properly (district roster + distance + TEA ratings).
    body.innerHTML =
      accessStrip() +
      // At-full-buildout and Commercial-gaps are built but not shown: the
      // population and retail-ratio assumptions behind them still need work and
      // would invite questions in a presentation. Flip SHOW_DEMAND_SECTIONS to
      // bring them back - the code and the API fields are untouched.
      (SHOW_DEMAND_SECTIONS ? buildoutStrip() + opportunityBlock() : '') +
      renderList('Grocery stores', d.grocery_stores) +
      renderList('Retail anchors', d.retail_anchors) +
      renderList('Hospitals', d.hospitals) +
      renderList('Pharmacies', d.pharmacies) +
      renderList('Parks', d.parks) +
      renderList('Fuel', d.fuel) +
      `<div style="font-size:11px;color:#6B7B8B;margin-top:6px">Source: ${esc(d.source)}</div>`;
    memoPut('amenities', d);
  } catch (e) {
    body.innerHTML = `<div class="placeholder err">Amenities load failed: ${esc(e.message)}</div>`;
  }
}

// ---------------------------------------------------------------------------
// Acquisition memo
// ---------------------------------------------------------------------------
// Every card fetches its own slice; the memo synthesises them. Each loader
// stashes its payload here on success and asks for a re-render, so the memo
// fills in progressively and never blocks on the slowest source.
const MEMO = {};
let _memoTimer = null;
function memoPut(key, data) {
  MEMO[key] = data;
  clearTimeout(_memoTimer);
  _memoTimer = setTimeout(renderMemo, 400);
}

// Print the memo as its own document rather than window.print() on the page —
// the project page carries maps, nav and a dozen other cards that would either
// need print-hiding rules or come out as noise. This writes just the memo into
// a new window with print styling, so "Save as PDF" produces a clean deliverable.
window.printMemo = function () {
  const body = document.getElementById('memo-body');
  if (!body) return;
  const p = projectData || {};
  const stamp = new Date().toLocaleDateString('en-US',
    { year: 'numeric', month: 'long', day: 'numeric' });
  const w = window.open('', '_blank');
  if (!w) { alert('Allow pop-ups for this site to print the memo.'); return; }
  w.document.write(`<!doctype html><html><head><meta charset="utf-8">
    <title>Acquisition Memo — ${esc(p.name || 'Project')}</title>
    <style>
      @page { size: letter; margin: 0.6in; }
      body { font-family: 'Inter', system-ui, -apple-system, sans-serif; color: #13344E;
             font-size: 11px; line-height: 1.5; margin: 0; }
      .doc-head { border-bottom: 3px solid #F25929; padding-bottom: 10px; margin-bottom: 16px;
                  display: flex; justify-content: space-between; align-items: flex-end; }
      .doc-head .brand { font-size: 10px; letter-spacing: .14em; text-transform: uppercase;
                         color: #F25929; font-weight: 700; }
      .doc-head h1 { font-size: 20px; margin: 4px 0 0; }
      .doc-head .meta { font-size: 10px; color: #6B7B8B; text-align: right; }
      h3 { font-size: 13px; margin: 16px 0 6px; padding-bottom: 3px;
           border-bottom: 1px solid #DDE3E8; page-break-after: avoid; }
      table { width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 6px; }
      th { background: #13344E !important; color: #FFF !important; padding: 5px 7px; text-align: left;
           -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      td { padding: 4px 7px; border-bottom: 1px solid #E8ECEF; }
      tr, table, .kpi-grid { page-break-inside: avoid; }
      .kpi-grid { display: grid; gap: 7px; margin-bottom: 10px; }
      .kpi { border: 1px solid #DDE3E8; border-radius: 4px; padding: 7px 9px; }
      .kpi .label { font-size: 8px; text-transform: uppercase; letter-spacing: .05em; color: #6B7B8B; }
      .kpi .value { font-size: 15px; font-weight: 700; }
      a { color: #13344E; text-decoration: none; }
      * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .doc-foot { margin-top: 18px; padding-top: 8px; border-top: 1px solid #DDE3E8;
                  font-size: 9px; color: #6B7B8B; }
    </style></head><body>
    <div class="doc-head">
      <div>
        <div class="brand">Ember Group — Acquisitions</div>
        <h1>${esc(p.name || 'Project')}</h1>
      </div>
      <div class="meta">Acquisition memo<br>${esc(stamp)}</div>
    </div>
    ${body.innerHTML}
    <div class="doc-foot">
      Generated from the Acquisitions GIS platform. Figures derive from CBAS, Census ACS, TxDOT,
      NCES/TEA, OpenStreetMap and FEMA/USGS as cited above.
    </div>
  </body></html>`);
  w.document.close();
  // Give the new document a beat to lay out before the print dialog opens.
  w.onload = () => setTimeout(() => { w.focus(); w.print(); }, 250);
  setTimeout(() => { try { w.focus(); w.print(); } catch (e) {} }, 700);
};

function renderMemo() {
  const card = document.getElementById('memo-card');
  const body = document.getElementById('memo-body');
  if (!card || !body) return;
  // The memo synthesises sections that are still being worked on, so it stays
  // hidden rather than half-right. memoPut still collects the data, so turning
  // SHOW_ACQUISITION_MEMO back on renders it immediately.
  if (!SHOW_ACQUISITION_MEMO) { card.style.display = 'none'; return; }
  card.style.display = 'block';

  const p = projectData || {};
  const an = MEMO.analysis || {};
  const cb = MEMO.cbas || {};
  const am = MEMO.amenities || {};
  const rd = MEMO.roads || {};
  const sm = MEMO.submarket || {};
  const nw = MEMO.news || {};
  const me = cb.market_entry || {};
  const bo = am.buildout || {};
  const agg = cb.aggregate || {};

  const money = v => v == null ? '—' : '$' + fmt(Math.round(v), 0);
  const mm = v => v == null ? '—'
    : v >= 1e9 ? '$' + (v / 1e9).toFixed(2) + 'B'
    : '$' + fmt(Math.round(v / 1e6), 0) + 'M';

  // ---- Site ----------------------------------------------------------
  const acres = an.gross_acres != null ? an.gross_acres : (p.total_acres || null);
  const net = an.net_developable_acres;
  const netPct = (acres && net != null) ? (net / acres * 100) : null;
  const y = an.yield_estimates || {};
  const lots = y.total_lots != null ? y.total_lots : (y.by_units_per_acre || null);

  // Lot absorption runs off STARTS, not closings — a start is a builder pulling
  // a lot, which is the transaction we're underwriting. Base capture defaults to
  // what a strong community in this ring actually achieves (p75 of observed
  // shares), so the assumption is anchored to demonstrated performance.
  const cap = me.capture || {};
  const baseCapture = cap.base_capture_pct != null ? cap.base_capture_pct : null;
  // Prefer the addressable denominator — starts in the lot widths this project
  // actually builds — since that's what it competes for.
  const captureBase = cap.addressable_starts || cap.ring_annual_starts;
  const lotsPerYear = (captureBase && baseCapture)
    ? Math.round(captureBase * baseCapture / 100) : null;
  const sellout = (lots && lotsPerYear) ? lots / lotsPerYear : null;

  // ---- Signals: transparent, rule-based, each one auditable ----------
  const sig = [];
  const add = (dir, label, detail) => sig.push({ dir, label, detail });

  if (me.months_lot_supply != null) {
    if (me.months_lot_supply < 12) add('+', 'Tight finished-lot supply',
      `${me.months_lot_supply} months of VDL at the current closing pace — lots are scarce here.`);
    else if (me.months_lot_supply > 24) add('-', 'Long finished-lot supply',
      `${me.months_lot_supply} months of VDL. Competitors already hold more finished inventory than the market absorbs.`);
    else add('~', 'Balanced lot supply', `${me.months_lot_supply} months of VDL.`);
  }
  if (cb.submarket_annual_closings != null) {
    if (cb.submarket_annual_closings >= 750) add('+', 'Deep absorption',
      `${fmt(cb.submarket_annual_closings,0)} closings/yr across ${cb.active_count} selling communities.`);
    else if (cb.submarket_annual_closings < 200) add('-', 'Thin absorption',
      `Only ${fmt(cb.submarket_annual_closings,0)} closings/yr in the submarket.`);
  }
  if (me.builder_concentration === 'fragmented') add('+', 'Fragmented builder field',
    `Top 3 take just ${fmt(me.top3_start_share_pct,0)}% of starts — no gatekeeper to displace.`);
  else if (me.builder_concentration === 'concentrated') add('-', 'Concentrated builder field',
    `Top 3 control ${fmt(me.top3_start_share_pct,0)}% of starts — harder to place lots.`);
  if (me.years_of_pipeline != null && me.years_of_pipeline > 10) add('-', 'Heavy future pipeline',
    `${me.years_of_pipeline} yrs of lots already planned in the submarket — future supply is not scarce.`);
  if (netPct != null) {
    if (netPct >= 80) add('+', 'Clean site', `${netPct.toFixed(0)}% of gross acreage is net developable.`);
    else if (netPct < 60) add('-', 'Constrained site',
      `Only ${netPct.toFixed(0)}% net developable after floodplain, slopes and easements.`);
  }
  const fp = an.floodplain_pct;
  if (fp != null && fp > 20) add('-', 'Significant floodplain', `${fp.toFixed(0)}% of the site is in the 100-yr floodplain.`);
  if (am.highways && am.highways.length) {
    const h = am.highways[0];
    if (h.distance_mi <= 3) add('+', 'Strong highway access', `${h.ref} is ${h.distance_mi} mi away.`);
    else if (h.distance_mi > 8) add('-', 'Remote from the freeway', `Nearest is ${h.ref} at ${h.distance_mi} mi.`);
  }
  const nearRoad = (rd.planned || []).filter(x => x.horizon_rank <= 1);
  if (nearRoad.length) add('+', 'Road capacity coming',
    `${nearRoad.length} TxDOT project${nearRoad.length===1?'':'s'} within 4 yrs, incl. ${esc(nearRoad[0].roadway)} ${esc(nearRoad[0].work||'')}.`);
  const sch = MEMO.schools || {};
  if (sch.growth && sch.growth.total_pct != null) {
    if (sch.growth.total_pct > 10) add('+', 'District enrolment growing',
      `${esc(sch.name || 'District')} +${sch.growth.total_pct.toFixed(1)}% over ${sch.growth.span_years} yrs.`);
    else if (sch.growth.total_pct < 0) add('-', 'District enrolment shrinking',
      `${esc(sch.name || 'District')} ${sch.growth.total_pct.toFixed(1)}% over ${sch.growth.span_years} yrs.`);
  }
  if (sch.tea && sch.tea.overall_rating) {
    const g = String(sch.tea.overall_rating).charAt(0);
    if (g === 'A' || g === 'B') add('+', `District rated ${sch.tea.overall_rating}`, 'TEA accountability rating supports pricing.');
    else if (g === 'D' || g === 'F') add('-', `District rated ${sch.tea.overall_rating}`, 'Weak TEA rating is a pricing headwind.');
  }
  if (bo.growth_pct != null && bo.growth_pct > 50) add('+', 'Large population runway',
    `Trade-area population grows ${bo.growth_pct}% to ${fmt(bo.population_at_buildout,0)} at full buildout.`);

  const pos = sig.filter(s => s.dir === '+').length;
  const neg = sig.filter(s => s.dir === '-').length;
  const verdict = (pos - neg >= 3) ? { v: 'PURSUE', c: '#2E7D32' }
                : (neg - pos >= 2) ? { v: 'PASS', c: '#C62828' }
                : { v: 'PURSUE WITH CONDITIONS', c: '#F59E0B' };

  // ---- Commercial opportunity ----------------------------------------
  const ops = (am.opportunities || []).filter(o => o.opportunity && !o.unknown);
  const parkGap = (am.opportunities || []).find(o => o.key === 'park');

  body.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;margin-bottom:10px">
      <div>
        <div style="font-size:18px;font-weight:700">${esc(p.name || 'Project')}</div>
        <div style="font-size:12px;color:#6B7B8B">
          ${fmt(acres,0)} gross ac · ${(p.tracts||[]).length} tract${(p.tracts||[]).length===1?'':'s'}
          ${sm.place ? ' · ' + esc(sm.place.name) : ''}
          ${cb.district_name ? ' · ' + esc(cb.district_name) : ''}
        </div>
      </div>
      <div style="text-align:right">
        <div style="background:${verdict.c};color:#FFF;padding:6px 14px;border-radius:4px;font-weight:700;font-size:14px">${verdict.v}</div>
        <div style="font-size:11px;color:#6B7B8B;margin-top:3px">${pos} supporting · ${neg} opposing signals</div>
      </div>
    </div>

    <div class="kpi-grid" style="grid-template-columns:repeat(5,1fr)">
      <div class="kpi"><div class="label">Net developable</div><div class="value">${fmt(net,0)} ac</div>
        <div class="label" style="margin-top:2px">${netPct != null ? netPct.toFixed(0) + '% of gross' : ''}</div></div>
      <div class="kpi"><div class="label">Est. lots</div><div class="value">${fmt(lots,0)}</div>
        <div class="label" style="margin-top:2px">at your yield mix</div></div>
      <div class="kpi" title="New-home starts in the trailing 12 months across every CBAS community whose centroid falls within ${cb.radius_mi} miles of this site. Each start consumes one lot.">
        <div class="label">Submarket lot demand</div>
        <div class="value">${fmt(cap.ring_annual_starts,0)}</div>
        <div class="label" style="margin-top:2px">new home starts / yr</div></div>
      <div class="kpi" style="${baseCapture ? 'background:#FFF4ED;border:1px solid #F25929' : ''}"
           title="Our share of those starts. Derived from starts, not home closings — a start is when the builder takes the lot.">
        <div class="label">Lots sold / yr</div>
        <div class="value" style="color:#F25929">${fmt(lotsPerYear,0)}</div>
        <div class="label" style="margin-top:2px">at ${baseCapture || '—'}% of starts</div></div>
      <div class="kpi"><div class="label">Sell-out</div>
        <div class="value">${sellout != null ? sellout.toFixed(1) : '—'}</div>
        <div class="label" style="margin-top:2px">years at that pace</div></div>
    </div>
    <div style="font-size:11px;color:#6B7B8B;margin-top:6px">
      <b>Submarket</b> = a <b>${cb.radius_mi}-mile radius from this site's centroid</b> — a distance radius, not a
      school district and not a CBAS submarket, because competitors don't stop at those lines. It covers
      ${cb.community_count} CBAS-tracked communities${cb.district_name ? ` spanning ${esc(cb.district_name)}` : ''}.
      <b>Lots sold / yr</b> is ${baseCapture}% of those starts — a home start is a builder taking down a lot, so
      starts (not home closings) are the lot demand this project competes for.
    </div>

    <h3>Recommendation</h3>
    <div style="font-size:12px;line-height:1.7">
      ${sig.length ? sig.map(s => `
        <div style="padding:3px 0">
          <span style="display:inline-block;width:16px;font-weight:700;color:${s.dir==='+'?'#2E7D32':s.dir==='-'?'#C62828':'#A5ADB7'}">${s.dir==='+'?'▲':s.dir==='-'?'▼':'●'}</span>
          <b>${esc(s.label)}</b> — ${s.detail}
        </div>`).join('') : '<div style="color:#A5ADB7">Run the analysis to generate signals.</div>'}
    </div>
    <h3>Market entry</h3>
    <div style="font-size:12px;line-height:1.7">
      ${cb.community_count ? `
        • The ${cb.radius_mi}-mile submarket holds <b>${cb.community_count} communities</b> (${cb.active_count} actively closing)
          absorbing <b>${fmt(cb.submarket_annual_closings,0)} homes/yr</b>, worth <b>${mm(cb.annual_sales_volume)}</b> in annual sales.<br>
        • Dominant product is <b>${esc(me.dominant_product || '—')}</b>${me.dominant_product_price ? ` at ${money(me.dominant_product_price)}` : ''};
          best $/SF is <b>${esc(me.highest_ppsf_band || '—')}</b>${me.highest_ppsf ? ` at $${me.highest_ppsf}/SF` : ''}.
          Prices run ${money(me.price_band && me.price_band.min)}–${money(me.price_band && me.price_band.max)},
          submarket median ${money(me.price_band && me.price_band.median)}.<br>
        • Lot market is <b>${esc(me.lot_market || '—')}</b> (${me.months_lot_supply} mos supply, ${me.years_of_pipeline} yrs pipeline).
          Builder field is <b>${esc(me.builder_concentration || '—')}</b>: ${(me.top3_builders||[]).join(', ')}
          hold ${fmt(me.top3_start_share_pct,0)}% of starts.<br>
        ${lots && cap.ring_annual_starts ? `• At ${fmt(lots,0)} lots this project is
          <b>${(lots / cap.ring_annual_starts).toFixed(1)}×</b> one year of submarket-wide lot demand.` : ''}
      ` : '<span style="color:#A5ADB7">Competition data not loaded.</span>'}
    </div>

    ${cap.ring_annual_starts ? `
      <h3>Lot absorption — capture of submarket demand</h3>
      <div style="font-size:12px;line-height:1.7;margin-bottom:8px">
        The submarket starts <b>${fmt(cap.ring_annual_starts,0)} homes/yr</b> across
        ${cap.active_communities} communities. Each start is a lot sold, so that — not closings — is the demand
        this project competes for.
        Communities here capture a median of <b>${cap.share_median_pct}%</b> of that; a strong one takes
        <b>${cap.share_p75_pct}%</b>, and the best performer, <b>${esc(cap.top_community || '—')}</b> at
        ${fmt(cap.top_community_starts,0)} starts/yr, takes <b>${cap.share_max_pct}%</b>.
        ${cap.addressable_starts ? `<br>You compete for a slice of that, not all of it: your target widths
          (${(cap.target_bands||[]).join(', ')}) are <b>${cap.segment_share_pct}%</b> of the submarket's lots, so the
          addressable demand is roughly <b>${fmt(cap.addressable_starts,0)} starts/yr</b>. A 20–25% capture reads as
          unrealistic against all starts but is ordinary against that.` : ''}
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead>
          <tr style="background:#13344E;color:#FFF">
            <th style="text-align:left;padding:6px 8px">Capture</th>
            <th style="text-align:right;padding:6px 8px">of all submarket starts</th>
            ${cap.addressable_starts ? `<th style="text-align:right;padding:6px 8px">of your product segment</th>` : ''}
            <th style="text-align:right;padding:6px 8px">Years to sell ${lots ? fmt(lots,0) + ' lots' : ''}</th>
            <th style="text-align:left;padding:6px 8px">vs what this submarket demonstrates</th>
          </tr>
        </thead>
        <tbody>
          ${(cap.scenarios || []).map(s => {
            const seg = s.addressable_lots_per_year;
            const yrs = lots && seg ? (lots / seg) : (lots && s.lots_per_year ? lots / s.lots_per_year : null);
            const isBase = s.capture_pct === baseCapture;
            return `
              <tr style="border-bottom:1px solid #DDE3E8;${isBase ? 'background:#FFF9F5' : s.above_ceiling ? 'opacity:.7' : ''}">
                <td style="padding:5px 8px"><b>${s.capture_pct}%</b>${isBase ? ' <span style="color:#F25929;font-size:10px;font-weight:700">BASE</span>' : ''}</td>
                <td style="padding:5px 8px;text-align:right">${fmt(s.lots_per_year,0)}</td>
                ${cap.addressable_starts ? `<td style="padding:5px 8px;text-align:right"><b>${fmt(seg,0)}</b></td>` : ''}
                <td style="padding:5px 8px;text-align:right">${yrs != null ? yrs.toFixed(1) : '—'}</td>
                <td style="padding:5px 8px;font-size:11px;color:${s.above_ceiling ? '#C62828' : '#6B7B8B'}">
                  ${s.above_ceiling ? `above the ${cap.share_max_pct}% ceiling seen here` : esc(s.label)}
                </td>
              </tr>`;
          }).join('')}
        </tbody>
      </table>
      <div style="font-size:11px;color:#6B7B8B;margin-top:4px">
        Base case ${baseCapture}% is the 75th percentile of capture actually achieved in this submarket.
        ${cap.addressable_starts ? 'Sell-out years use the product-segment column.' : ''}
      </div>` : ''}

    <h3>Demand at full buildout</h3>
    <div style="font-size:12px;line-height:1.7">
      ${bo.population_at_buildout ? `
        • Trade area grows from <b>${fmt(bo.population_now,0)}</b> to <b>${fmt(bo.population_at_buildout,0)}</b> people
          (${bo.growth_pct != null ? '+' + bo.growth_pct + '%' : ''}) as ${fmt(bo.lots_remaining,0)} remaining lots build out —
          roughly <b>${bo.years_to_buildout || '—'} years</b> at ${fmt(bo.annual_closings,0)} closings/yr.<br>
        • ${fmt(bo.rooftops_now,0)} rooftops today → <b>${fmt(bo.rooftops_at_buildout,0)}</b> at buildout,
          at ${bo.avg_household_size} people per home.
      ` : '<span style="color:#A5ADB7">Buildout projection not available.</span>'}
    </div>

    <h3>Commercial &amp; amenity opportunity</h3>
    <div style="font-size:12px;line-height:1.7">
      ${ops.length ? `
        • That population supports commercial this trade area does not yet have. The clearest gaps:
          ${ops.slice(0,5).map(o => `<b>${esc(o.use.toLowerCase())}</b> (supports ${o.supported_at_buildout}, has ${o.existing})`).join(' · ')}.<br>
        • A pad or two of neighbourhood retail — grocery-anchored where the gap supports it — turns a
          commodity subdivision into a destination and pulls forward absorption. Worth reserving frontage for.<br>
      ` : '<span style="color:#A5ADB7">No material commercial shortfall, or amenity data not loaded.</span><br>'}
      ${parkGap && parkGap.gap_at_buildout > 0 ? `
        • Parks and open space run <b>${parkGap.gap_at_buildout}</b> short of what the buildout population supports
          (${parkGap.existing} within ${am.radius_mi} mi). In a ${esc(me.dominant_product || 'competitive')} market where
          product is near-identical, amenity is the differentiator — trails, a lake, a pool amenity centre.<br>` : ''}
      ${(am.grocery_stores && am.grocery_stores.length && !am.grocery_stores[0].within_radius) ? `
        • Nearest grocery is <b>${esc(am.grocery_stores[0].name)}</b> at ${am.grocery_stores[0].distance_mi} mi —
          outside the ${am.radius_mi}-mi submarket. Daily-needs retail is a real gap for residents today.` : ''}
    </div>

    <h3>Access &amp; infrastructure</h3>
    <div style="font-size:12px;line-height:1.7">
      ${(am.highways && am.highways.length) ? `
        • Nearest freeway <b>${esc(am.highways[0].ref)}</b> at ${am.highways[0].distance_mi} mi${am.nearest_ramp_mi != null ? `, closest ramp ${am.nearest_ramp_mi} mi` : ''}.
          ${am.airports && am.airports.length ? `${esc(am.airports[0].iata)} ${am.airports[0].distance_mi} mi.` : ''}<br>` : ''}
      ${(rd.planned && rd.planned.length) ? `
        • TxDOT has <b>${rd.upcoming_count}</b> upcoming projects within ${rd.radius_mi} mi
          (${(rd.counts||{}).near_term || 0} breaking ground inside 4 yrs).
          ${nearRoad.slice(0,3).map(x => `<b>${esc(x.roadway)}</b> ${esc(x.work||'')}${x.let_year ? ` (${x.let_year})` : ''}`).join(' · ')}
      ` : '<span style="color:#A5ADB7">Road data not loaded.</span>'}
    </div>

    ${(nw.stories && nw.stories.length) ? `
      <h3>What's in the news</h3>
      <div style="font-size:12px;line-height:1.6">
        ${nw.stories.slice(0,6).map(s => `
          <div style="padding:3px 0">
            • <a href="${esc(s.link)}" target="_blank" style="color:#1976D2;text-decoration:none">${esc(s.title)}</a>
            <span style="color:#A5ADB7;font-size:10px">${esc(s.source || '')}</span>
          </div>`).join('')}
      </div>` : ''}

    <h3>Assumptions</h3>
    <div style="font-size:12px;line-height:1.7;color:#13344E">
      • Yield of <b>${fmt(lots,0)} lots</b> on ${fmt(net,0)} net acres uses the product mix set on this page, not an engineered plan.<br>
      • Lot absorption of <b>${fmt(lotsPerYear,0)} lots/yr</b> assumes <b>${baseCapture}%</b> capture of
        ${cap.addressable_starts
          ? `the ${fmt(cap.addressable_starts,0)} annual starts in your target lot widths (${cap.segment_share_pct}% of submarket product)`
          : `the submarket's ${fmt(cap.ring_annual_starts,0)} annual starts`}.
        ${baseCapture}% is the 75th percentile of capture actually achieved here, not a round number.
        Starts rather than closings, because a start is when the builder buys the lot.
        A new community ramps below its run rate in year one.<br>
      • Sales volume prices each community at its own CBAS midpoint${me.unpriced_communities ? `; ${me.unpriced_communities} communities without CBAS pricing use the ring median` : ''}.<br>
      • Buildout population applies ${bo.avg_household_size || '—'} people per home (${esc(bo.household_size_source || 'n/a')})
        to the remaining lot pipeline; today's base is ${esc(bo.population_basis || 'n/a')}.<br>
      • Commercial gaps use trade-area planning ratios to rank under-supply — they are not a market study.<br>
      • Constraints come from public GIS (FEMA, USGS, NWI, TxGIO); none of it substitutes for survey or geotech.
    </div>

    <div style="font-size:11px;color:#6B7B8B;margin-top:10px">
      Sources: CBAS new-home survey · Census ACS ${(bo.demographics||{}).year || ''} · TxDOT · NCES/TEA · OpenStreetMap · FEMA/USGS · Google News.
      ${['analysis','cbas','amenities','roads','submarket','news'].filter(k => !MEMO[k]).length
        ? `Still loading: ${['analysis','cbas','amenities','roads','submarket','news'].filter(k => !MEMO[k]).join(', ')}.` : ''}
    </div>
  `;
}

// Search the whole CBAS universe, not just the ring — for pulling up a specific
// comparable or checking where a builder is active outside the 8-mi radius.
window.searchCbas = async function () {
  const inp = document.getElementById('cbas-q');
  const out = document.getElementById('cbas-search-results');
  if (!inp || !out) return;
  const q = (inp.value || '').trim();
  if (q.length < 2) { out.innerHTML = '<div style="font-size:11px;color:#A5ADB7">Enter at least 2 characters.</div>'; return; }
  out.innerHTML = '<div style="font-size:11px;color:#6B7B8B">Searching…</div>';
  try {
    const r = await fetch(`/api/acq/projects/${PROJECT_ID}/cbas/search?q=${encodeURIComponent(q)}`);
    const d = await r.json();
    if (!r.ok || d.error) {
      out.innerHTML = `<div style="font-size:11px;color:#C62828">${esc(d.error || ('HTTP ' + r.status))}</div>`;
      return;
    }
    if (!d.community_count && !d.builder_count) {
      out.innerHTML = `<div style="font-size:11px;color:#A5ADB7">No CBAS community or builder matches "${esc(q)}" (searched ${fmt(d.universe,0)}).</div>`;
      return;
    }
    const dl = (mi, lat, lon, nm) => distLink(mi, null, lat, lon, nm);
    out.innerHTML = `
      <div style="border:1px solid #13344E;border-radius:6px;padding:10px">
        <div style="font-size:11px;color:#6B7B8B;margin-bottom:6px">
          "${esc(d.query)}" — ${d.community_count} communities, ${d.builder_count} builders
          <span style="color:#A5ADB7">(searched all ${fmt(d.universe,0)} CBAS communities)</span>
          <button class="ghost" onclick="document.getElementById('cbas-search-results').innerHTML='';document.getElementById('cbas-q').value=''"
                  style="padding:2px 8px;font-size:10px;margin-left:6px">Clear</button>
        </div>
        ${d.builders.length ? `
          <table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:8px">
            <thead><tr style="background:#13344E;color:#FFF">
              <th style="text-align:left;padding:5px 8px">Builder</th>
              <th style="text-align:right;padding:5px 8px">Communities</th>
              <th style="text-align:right;padding:5px 8px">Lots</th>
              <th style="text-align:right;padding:5px 8px">Avg price</th>
              <th style="text-align:right;padding:5px 8px">$/SF</th>
              <th style="text-align:right;padding:5px 8px">Nearest</th>
              <th style="text-align:left;padding:5px 8px">Closest communities</th>
            </tr></thead>
            <tbody>
              ${d.builders.map(b => `
                <tr style="border-bottom:1px solid #DDE3E8">
                  <td style="padding:5px 8px"><b>${esc(b.name || '')}</b></td>
                  <td style="padding:5px 8px;text-align:right">${fmt(b.community_count,0)}</td>
                  <td style="padding:5px 8px;text-align:right">${fmt(b.lots,0)}</td>
                  <td style="padding:5px 8px;text-align:right">${b.avg_price ? '$' + fmt(b.avg_price,0) : '—'}</td>
                  <td style="padding:5px 8px;text-align:right">${b.avg_ppsf ? '$' + b.avg_ppsf.toFixed(0) : '—'}</td>
                  <td style="padding:5px 8px;text-align:right">${b.nearest_mi != null ? b.nearest_mi + ' mi' : '—'}</td>
                  <td style="padding:5px 8px;font-size:11px;color:#6B7B8B">
                    ${(b.communities||[]).slice(0,3).map(c => esc(c.name) + (c.distance_mi != null ? ` (${c.distance_mi}mi)` : '')).join(', ')}
                  </td>
                </tr>`).join('')}
            </tbody>
          </table>` : ''}
        ${d.communities.length ? `
          <table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead><tr style="background:#13344E;color:#FFF">
              <th style="text-align:left;padding:5px 8px">Community</th>
              <th style="text-align:right;padding:5px 8px">Distance</th>
              <th style="text-align:left;padding:5px 8px">Builders</th>
              <th style="text-align:left;padding:5px 8px">Lot widths</th>
              <th style="text-align:right;padding:5px 8px">Price</th>
              <th style="text-align:right;padding:5px 8px">Ann. closings</th>
              <th style="text-align:right;padding:5px 8px">VDL</th>
              <th style="text-align:right;padding:5px 8px">Total lots</th>
            </tr></thead>
            <tbody>
              ${d.communities.map(c => `
                <tr style="border-bottom:1px solid #DDE3E8">
                  <td style="padding:5px 8px"><b>${esc(c.name || '')}</b>
                    ${c.status ? `<span style="color:#6B7B8B;font-size:9px;font-weight:700;margin-left:4px">${esc(String(c.status).toUpperCase())}</span>` : ''}
                    <div style="color:#A5ADB7;font-size:10px">${esc(c.developer || '')}${c.city ? ' · ' + esc(c.city) : ''}${c.county ? ' · ' + esc(c.county) : ''}</div>
                  </td>
                  <td style="padding:5px 8px;text-align:right">${dl(c.distance_mi, c.lat, c.lon, c.name)}</td>
                  <td style="padding:5px 8px;font-size:11px">${(c.builders||[]).slice(0,3).map(esc).join(', ') || '—'}</td>
                  <td style="padding:5px 8px;font-size:11px">${(c.lot_types_ff||[]).join(', ') || '—'}</td>
                  <td style="padding:5px 8px;text-align:right;font-size:11px">${c.price_min ? '$' + fmt(Math.round(c.price_min/1000),0) + 'K–$' + fmt(Math.round(c.price_max/1000),0) + 'K' : '—'}</td>
                  <td style="padding:5px 8px;text-align:right"><b>${fmt(c.annual_closings,0)}</b></td>
                  <td style="padding:5px 8px;text-align:right">${fmt(c.vdls,0)}</td>
                  <td style="padding:5px 8px;text-align:right">${fmt(c.total_lots,0)}</td>
                </tr>`).join('')}
            </tbody>
          </table>` : ''}
      </div>`;
  } catch (e) {
    out.innerHTML = `<div style="font-size:11px;color:#C62828">Search failed: ${esc(e.message)}</div>`;
  }
};

let _roadsMap = null;
async function loadRoads() {
  const card = document.getElementById('roads-card');
  const body = document.getElementById('roads-body');
  card.style.display = 'block';
  body.innerHTML = '<div class="placeholder">Querying TxDOT planned &amp; programmed projects…</div>';
  try {
    const r = await fetch(`/api/acq/projects/${PROJECT_ID}/roads?radius_mi=15`);
    const d = await r.json();
    if (!r.ok || d.error) {
      body.innerHTML = `<div class="placeholder">${esc(d.error || ('HTTP ' + r.status))}</div>`;
      return;
    }
    const cn = d.counts || {};
    const hColor = h => h === 0 ? '#C62828' : h === 1 ? '#F25929'
                      : h === 2 ? '#F59E0B' : h === 3 ? '#1976D2' : '#A5ADB7';
    const upcoming = (d.planned || []).filter(p => p.horizon_rank <= 3);

    // Lead with capacity work only. Of ~48 scheduled projects in a 15-mi box,
    // most are overlays, seal coats, signal upgrades and safety lighting — none
    // of which change what a tract is worth. The rest stays available but folded
    // away rather than crowding out the handful that matter.
    const capSched = upcoming.filter(p => p.capacity);
    const otherSched = upcoming.filter(p => !p.capacity);
    const capProg = (d.programmed || []).filter(p => p.capacity);
    // Summarised server-side over the full capacity set — `programmed` is
    // truncated in the response, so counting it here would undercount.
    const roadSummary = d.capacity_by_roadway || [];

    body.innerHTML = `
      <div class="kpi-grid" style="grid-template-columns:repeat(3,1fr)">
        <div class="kpi" style="background:#FFF4ED;border:1px solid #F25929">
          <div class="label">Capacity work, next 4 yrs</div>
          <div class="value" style="color:#F25929">${fmt(capSched.filter(p => p.horizon_rank <= 1).length,0)}</div>
          <div class="label" style="margin-top:2px">funded with a date</div>
        </div>
        <div class="kpi"><div class="label">Capacity in the programme</div><div class="value">${fmt(cn.capacity,0)}</div>
          <div class="label" style="margin-top:2px">${fmt(cn.capacity_scheduled,0)} of them dated</div></div>
        <div class="kpi"><div class="label">Future toll corridors</div><div class="value">${fmt(cn.future_tolls,0)}</div></div>
      </div>

      <h3>Capacity work with a construction date</h3>
      ${capSched.length ? `
        <table style="width:100%;border-collapse:collapse;font-size:12px">
          <thead>
            <tr style="background:#13344E;color:#FFF">
              <th style="text-align:left;padding:6px 8px">Roadway</th>
              <th style="text-align:left;padding:6px 8px">Work</th>
              <th style="text-align:left;padding:6px 8px">Limits</th>
              <th style="text-align:right;padding:6px 8px" title="TxDOT letting year — the year the construction contract is bid and awarded. Dirt typically moves within months of it.">Construction start</th>
            </tr>
          </thead>
          <tbody>
            ${capSched.slice(0,10).map(p => `
              <tr style="border-bottom:1px solid #DDE3E8">
                <td style="padding:5px 8px"><b>${esc(p.roadway || '')}</b></td>
                <td style="padding:5px 8px">${esc(p.description || p.work || '')}</td>
                <td style="padding:5px 8px;font-size:11px;color:#6B7B8B">${esc(p.from || '')}${p.to && p.to !== '.' ? ' → ' + esc(p.to) : ''}</td>
                <td style="padding:5px 8px;text-align:right">
                  <b style="color:${hColor(p.horizon_rank)}">${p.let_year || '—'}</b>
                </td>
              </tr>`).join('')}
          </tbody>
        </table>` : '<div style="font-size:12px;color:#A5ADB7">No funded capacity work scheduled within ' + d.radius_mi + ' mi.</div>'}

      ${roadSummary.length ? `
        <div style="font-size:12px;margin-top:10px;line-height:1.6">
          <b>Programmed but undated:</b> ${fmt(cn.capacity,0)} more capacity projects sit in TxDOT's inventory with no
          construction date published — ${roadSummary.slice(0,5).map(([r, n]) => `<b>${esc(r)}</b> ${n}`).join(' · ')}${roadSummary.length > 5 ? ` · +${roadSummary.length - 5} more roads` : ''}.
        </div>` : ''}

      <details style="margin-top:10px">
        <summary style="font-size:11px;color:#6B7B8B;cursor:pointer;font-weight:600">
          Show the full project list (${upcoming.length} scheduled, ${capProg.length} programmed capacity)
        </summary>
        <div style="max-height:340px;overflow-y:auto;margin-top:6px">
          <table style="width:100%;border-collapse:collapse;font-size:11px">
            <thead>
              <tr style="background:#13344E;color:#FFF">
                <th style="text-align:left;padding:5px 7px">Roadway</th>
                <th style="text-align:left;padding:5px 7px">Work</th>
                <th style="text-align:left;padding:5px 7px">Limits</th>
                <th style="text-align:right;padding:5px 7px">Let</th>
              </tr>
            </thead>
            <tbody>
              ${otherSched.map(p => `
                <tr style="border-bottom:1px solid #F0F0F0">
                  <td style="padding:4px 7px">${esc(p.roadway || '')}</td>
                  <td style="padding:4px 7px;color:#6B7B8B">${esc(p.description || p.work || '')}</td>
                  <td style="padding:4px 7px;color:#A5ADB7">${esc(p.from || '')}${p.to && p.to !== '.' ? ' → ' + esc(p.to) : ''}</td>
                  <td style="padding:4px 7px;text-align:right">${p.let_year || '—'}</td>
                </tr>`).join('')}
              ${capProg.map(p => `
                <tr style="border-bottom:1px solid #F0F0F0">
                  <td style="padding:4px 7px">${esc(p.highway || '')}</td>
                  <td style="padding:4px 7px;color:#6B7B8B">${esc(p.klass || '')}</td>
                  <td style="padding:4px 7px;color:#A5ADB7">${esc(p.from || '')}${p.to && p.to !== '.' ? ' → ' + esc(p.to) : ''}</td>
                  <td style="padding:4px 7px;text-align:right;color:#A5ADB7">undated</td>
                </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </details>

      ${(d.future_tolls || []).length ? `
        <h3>Future toll corridors</h3>
        <div style="font-size:12px">
          ${d.future_tolls.map(t => `
            <div style="padding:5px 8px;border-bottom:1px solid #F0F0F0">
              <b>${esc(t.name || '')}</b>
              ${t.year_open ? `<span style="color:#F25929;font-weight:700;margin-left:6px">opens ${esc(String(t.year_open))}</span>` : ''}
              ${t.operator ? `<div style="font-size:11px;color:#6B7B8B">${esc(t.operator)}${t.note ? ' · ' + esc(t.note) : ''}</div>` : ''}
            </div>`).join('')}
        </div>` : ''}

      <h3>Project map</h3>
      <div id="roads-map" style="height:400px;border:1px solid #DDE3E8;border-radius:6px"></div>
      <div style="font-size:11px;color:#6B7B8B;margin-top:6px">
        Source: ${esc(d.source)}. Capacity means widening, added lanes, interchange or new location —
        overlays, seal coats, signals and striping are excluded because they don't change what a tract is worth.
      </div>`;

    memoPut('roads', d);
    setTimeout(() => {
      const el = document.getElementById('roads-map');
      if (!el || typeof L === 'undefined') return;
      try {
        if (_roadsMap) { _roadsMap.remove(); _roadsMap = null; }
        // See renderCbasMap — setView before adding layers, or Leaflet throws.
        const rc = d.center || {};
        _roadsMap = L.map('roads-map', { scrollWheelZoom: false })
                     .setView([rc.lat != null ? rc.lat : 29.76, rc.lon != null ? rc.lon : -95.37], 11);
        L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}',
          { maxZoom: 19, attribution: 'Esri' }).addTo(_roadsMap);
        const layers = [];
        const tg = ((projectData && projectData.tracts) || []).filter(t => t.geometry)
          .map(t => ({ type: 'Feature', properties: {}, geometry: t.geometry }));
        if (tg.length) {
          layers.push(L.geoJSON({ type: 'FeatureCollection', features: tg },
            { style: { color: '#F25929', weight: 3, fillColor: '#F25929', fillOpacity: 0.4 } })
            .bindPopup('<b>Your tract</b>').addTo(_roadsMap));
        } else if (d.center) {
          layers.push(L.circleMarker([d.center.lat, d.center.lon],
            { radius: 9, color: '#FFF', weight: 2, fillColor: '#F25929', fillOpacity: 1 })
            .bindPopup('<b>Your tract</b>').addTo(_roadsMap));
        }
        // Capacity work only — plotting every overlay and signal upgrade turns
        // the map into a road atlas of the county.
        capSched.forEach(p => {
          if (!p.geometry) return;
          const ly = L.geoJSON(p.geometry, { style: { color: hColor(p.horizon_rank), weight: 4, opacity: 0.85 } })
            .bindPopup(`<div style="font-size:12px">
                <b>${esc(p.roadway || '')}</b> — ${esc(p.description || p.work || '')}<br>
                <span style="color:#6B7B8B">${esc(p.from || '')}${p.to && p.to !== '.' ? ' → ' + esc(p.to) : ''}</span><br>
                <b style="color:${hColor(p.horizon_rank)}">${esc(p.horizon)}</b>${p.let_year ? ` · let ${p.let_year}` : ''}
              </div>`).addTo(_roadsMap);
          layers.push(ly);
        });
        (d.programmed || []).filter(p => p.capacity && p.geometry).slice(0,60).forEach(p => {
          layers.push(L.geoJSON(p.geometry, { style: { color: '#7B1FA2', weight: 3, opacity: 0.7, dashArray: '6,4' } })
            .bindPopup(`<div style="font-size:12px"><b>${esc(p.highway||'')}</b> — ${esc(p.klass||'')}<br>
              <span style="color:#6B7B8B">${esc(p.from||'')}${p.to && p.to !== '.' ? ' → ' + esc(p.to) : ''}</span></div>`)
            .addTo(_roadsMap));
        });
        if (layers.length) _roadsMap.fitBounds(L.featureGroup(layers).getBounds().pad(0.08));
        else if (d.center) _roadsMap.setView([d.center.lat, d.center.lon], 11);
        setTimeout(() => { try { _roadsMap.invalidateSize(); } catch (e) {} }, 120);
      } catch (e) {
        el.innerHTML = `<div class="placeholder">Map failed: ${esc(e.message)}</div>`;
      }
    }, 60);
  } catch (e) {
    body.innerHTML = `<div class="placeholder err">Road data failed: ${esc(e.message)}</div>`;
  }
}

// Competitor map: every CBAS community in the ring, sized by annual closings,
// against the project outline. Built after innerHTML lands so the div exists.
let _cbasMap = null;
function renderCbasMap(d) {
  const el = document.getElementById('cbas-map');
  if (!el || typeof L === 'undefined') return;
  const c = d.center || {};
  if (c.lat == null) return;
  try {
    if (_cbasMap) { _cbasMap.remove(); _cbasMap = null; }
    // setView must happen before any layer is added: Leaflet projects each
    // layer on add, and with no view it has no pixel origin — the failure is
    // an opaque "Cannot read properties of undefined (reading
    // 'layerPointToLatLng')". fitBounds below then refines this.
    _cbasMap = L.map('cbas-map', { scrollWheelZoom: false }).setView([c.lat, c.lon], 11);
    L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}',
      { maxZoom: 19, attribution: 'Esri' }).addTo(_cbasMap);

    const layers = [];
    // The project itself — real tract outlines when loaded, else the centroid.
    const tractGeoms = ((projectData && projectData.tracts) || [])
      .filter(t => t.geometry)
      .map(t => ({ type: 'Feature', properties: { prop_id: t.prop_id }, geometry: t.geometry }));
    if (tractGeoms.length) {
      const pg = L.geoJSON({ type: 'FeatureCollection', features: tractGeoms },
        { style: { color: '#F25929', weight: 3, fillColor: '#F25929', fillOpacity: 0.35 } }).addTo(_cbasMap);
      pg.bindPopup('<b>Your tract</b>');
      layers.push(pg);
    } else {
      layers.push(L.circleMarker([c.lat, c.lon], { radius: 10, color: '#FFF', weight: 2,
        fillColor: '#F25929', fillOpacity: 1 }).bindPopup('<b>Your tract</b>').addTo(_cbasMap));
    }
    // The search ring, so "within 8 mi" is visible rather than asserted.
    layers.push(L.circle([c.lat, c.lon], { radius: (d.radius_mi || 8) * 1609.34,
      color: '#F25929', weight: 1, dashArray: '5,5', fill: false }).addTo(_cbasMap));

    const maxCl = Math.max(1, ...(d.communities || []).map(x => x.annual_closings || 0));
    (d.communities || []).forEach(x => {
      if (x.lat == null || x.lon == null) return;
      const col = x.status === 'Active' ? '#2E7D32'
                : x.status === 'Future' ? '#1976D2' : '#9AA5B1';
      // sqrt keeps a 500-closing community from swallowing the map
      const r = 5 + 16 * Math.sqrt((x.annual_closings || 0) / maxCl);
      const mk = L.circleMarker([x.lat, x.lon], { radius: r, color: '#FFF', weight: 1.5,
        fillColor: col, fillOpacity: 0.8 }).addTo(_cbasMap);
      const gmap = `https://www.google.com/maps/dir/?api=1&origin=${c.lat},${c.lon}`
                 + `&destination=${x.lat},${x.lon}&travelmode=driving`;
      mk.bindPopup(`
        <div style="font-size:12px;min-width:210px">
          <b>${esc(x.name || '')}</b>${x.status ? ` <span style="color:#6B7B8B">${esc(x.status)}</span>` : ''}
          ${x.developer ? `<div style="color:#6B7B8B">${esc(x.developer)}</div>` : ''}
          <div style="margin-top:5px">
            <a href="${gmap}" target="_blank" rel="noopener">${x.distance_mi} mi ${esc(x.direction||'')} ↗</a>
          </div>
          <table style="margin-top:5px;font-size:11px;border-collapse:collapse">
            <tr><td style="padding-right:8px;color:#6B7B8B">Annual closings</td><td><b>${fmt(x.annual_closings,0)}</b></td></tr>
            <tr><td style="padding-right:8px;color:#6B7B8B">Annual starts</td><td>${fmt(x.annual_starts,0)}</td></tr>
            <tr><td style="padding-right:8px;color:#6B7B8B">VDL</td><td>${fmt(x.vdls,0)}</td></tr>
            <tr><td style="padding-right:8px;color:#6B7B8B">Total lots</td><td>${fmt(x.total_lots,0)}</td></tr>
            ${x.annual_volume ? `<tr><td style="padding-right:8px;color:#6B7B8B">Sales volume</td><td>$${fmt(Math.round(x.annual_volume/1e6),0)}M/yr</td></tr>` : ''}
            ${x.price_min ? `<tr><td style="padding-right:8px;color:#6B7B8B">Price</td><td>$${fmt(Math.round(x.price_min/1000),0)}K–$${fmt(Math.round(x.price_max/1000),0)}K</td></tr>` : ''}
            ${(x.lot_types_ff||[]).length ? `<tr><td style="padding-right:8px;color:#6B7B8B">Lot widths</td><td>${x.lot_types_ff.join(', ')} FF</td></tr>` : ''}
          </table>
          ${(x.builders||[]).length ? `<div style="margin-top:5px;font-size:11px;color:#6B7B8B">${x.builders.slice(0,4).map(esc).join(', ')}</div>` : ''}
        </div>`);
      layers.push(mk);
    });

    if (layers.length) {
      _cbasMap.fitBounds(L.featureGroup(layers).getBounds().pad(0.08));
    } else {
      _cbasMap.setView([c.lat, c.lon], 11);
    }
    // Leaflet mis-sizes a map created inside a freshly-written container.
    setTimeout(() => { try { _cbasMap.invalidateSize(); } catch (e) {} }, 120);
  } catch (e) {
    el.innerHTML = `<div class="placeholder">Map failed: ${esc(e.message)}</div>`;
  }
}

async function loadCbas() {
  const card = document.getElementById('cbas-card');
  const body = document.getElementById('cbas-body');
  card.style.display = 'block';
  body.innerHTML = '<div class="placeholder">Pulling CBAS new-home survey for the submarket…</div>';
  try {
    const r = await fetch(`/api/acq/projects/${PROJECT_ID}/cbas?radius_mi=8`);
    const d = await r.json();
    if (!r.ok || d.error) {
      body.innerHTML = `<div class="placeholder">${esc(d.error || ('HTTP ' + r.status))}</div>`;
      return;
    }
    const a = d.aggregate || {};
    // Absorption pill colouring — annual closings is the demand signal
    function absPill(v) {
      if (v == null) return '<span style="color:#A5ADB7">—</span>';
      const c = v >= 300 ? '#2E7D32' : v >= 100 ? '#7CB342' : v >= 25 ? '#F59E0B' : '#A5ADB7';
      return `<span style="background:${c};color:#FFF;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700">${fmt(v,0)}</span>`;
    }
    function money(v) { return v != null ? '$' + fmt(v, 0) : '—'; }
    function numDash(v, d) { return v != null ? fmt(v, d == null ? 0 : d) : '—'; }
    function priceRange(min, max) {
      return min ? '$' + fmt(Math.round(min / 1000), 0) + 'K–$' + fmt(Math.round(max / 1000), 0) + 'K' : '—';
    }
    function communityDetailHtml(c) {
      const detail = c.detail || {};
      const widths = detail.lot_widths || [];
      const builderLots = detail.builder_lot_widths || [];
      const inventory = `
        <div class="kpi-grid" style="grid-template-columns:repeat(5,1fr);margin:0 0 8px">
          <div class="kpi"><div class="label">Community avg price</div><div class="value">${money(c.avg_price)}</div></div>
          <div class="kpi"><div class="label">Price range</div><div class="value" style="font-size:15px">${priceRange(c.price_min, c.price_max)}</div></div>
          <div class="kpi"><div class="label">VDL inventory</div><div class="value">${numDash(c.vdls, 0)}</div></div>
          <div class="kpi"><div class="label">Under construction</div><div class="value">${numDash(c.under_construction, 0)}</div></div>
          <div class="kpi"><div class="label">Future lots</div><div class="value">${numDash(c.futures, 0)}</div></div>
        </div>
        <div style="font-size:10px;color:#6B7B8B;margin-bottom:8px">
          Builder and lot-width lot counts come from CBAS section detail. Est. VDL / starts / closings are apportioned by each row's share of those section lots.
        </div>`;
      const widthTable = widths.length ? `
        <h4 style="margin:4px 0 5px;color:#083763;font-size:12px">By lot width</h4>
        <table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:10px">
          <thead>
            <tr style="background:#EFF4F8;color:#13344E">
              <th style="text-align:left;padding:5px 7px">Lot width</th>
              <th style="text-align:right;padding:5px 7px">Lots</th>
              <th style="text-align:right;padding:5px 7px">Est. VDL</th>
              <th style="text-align:right;padding:5px 7px">Est. pipeline</th>
              <th style="text-align:right;padding:5px 7px">Est. ann. starts</th>
              <th style="text-align:left;padding:5px 7px">Builders</th>
              <th style="text-align:right;padding:5px 7px">Avg price</th>
              <th style="text-align:right;padding:5px 7px">Range</th>
              <th style="text-align:right;padding:5px 7px">Avg SF</th>
              <th style="text-align:right;padding:5px 7px">$/SF</th>
              <th style="text-align:right;padding:5px 7px">Plans</th>
            </tr>
          </thead>
          <tbody>
            ${widths.map(w => `
              <tr style="border-bottom:1px solid #E8ECEF">
                <td style="padding:5px 7px"><b>${numDash(w.lot_width_ff, 0)} FF</b></td>
                <td style="padding:5px 7px;text-align:right"><b>${numDash(w.lots, 0)}</b></td>
                <td style="padding:5px 7px;text-align:right">${numDash(w.est_vdls, 0)}</td>
                <td style="padding:5px 7px;text-align:right">${numDash(w.est_futures, 0)}</td>
                <td style="padding:5px 7px;text-align:right">${w.est_annual_starts != null ? fmt(w.est_annual_starts, 1) : '—'}</td>
                <td style="padding:5px 7px;color:#6B7B8B">${(w.builders || []).slice(0,4).map(b => esc(b.name) + ' (' + numDash(b.lots, 0) + ')').join(', ') || '—'}${w.builder_count > 4 ? ' +' + (w.builder_count - 4) : ''}</td>
                <td style="padding:5px 7px;text-align:right">${money(w.avg_price)}</td>
                <td style="padding:5px 7px;text-align:right;color:#6B7B8B">${priceRange(w.min_price, w.max_price)}</td>
                <td style="padding:5px 7px;text-align:right">${numDash(w.avg_sqft, 0)}</td>
                <td style="padding:5px 7px;text-align:right">${w.avg_ppsf ? '$' + w.avg_ppsf.toFixed(0) : '—'}</td>
                <td style="padding:5px 7px;text-align:right">${numDash(w.plans, 0)}</td>
              </tr>`).join('')}
          </tbody>
        </table>` : '<div style="font-size:11px;color:#A5ADB7;margin-bottom:8px">No lot-width section detail returned by CBAS for this community.</div>';
      const builderTable = builderLots.length ? `
        <h4 style="margin:4px 0 5px;color:#083763;font-size:12px">By builder and lot width</h4>
        <table style="width:100%;border-collapse:collapse;font-size:11px">
          <thead>
            <tr style="background:#EFF4F8;color:#13344E">
              <th style="text-align:left;padding:5px 7px">Builder</th>
              <th style="text-align:left;padding:5px 7px">Lot width</th>
              <th style="text-align:right;padding:5px 7px">Lots</th>
              <th style="text-align:right;padding:5px 7px">Est. VDL</th>
              <th style="text-align:right;padding:5px 7px">Est. ann. starts</th>
              <th style="text-align:right;padding:5px 7px">Est. ann. closings</th>
              <th style="text-align:right;padding:5px 7px">Avg price</th>
              <th style="text-align:right;padding:5px 7px">Range</th>
              <th style="text-align:right;padding:5px 7px">Avg SF</th>
              <th style="text-align:right;padding:5px 7px">$/SF</th>
              <th style="text-align:right;padding:5px 7px">Plans</th>
            </tr>
          </thead>
          <tbody>
            ${builderLots.map(b => `
              <tr style="border-bottom:1px solid #E8ECEF">
                <td style="padding:5px 7px"><b>${esc(b.name)}</b></td>
                <td style="padding:5px 7px">${numDash(b.lot_width_ff, 0)} FF</td>
                <td style="padding:5px 7px;text-align:right"><b>${numDash(b.lots, 0)}</b></td>
                <td style="padding:5px 7px;text-align:right">${numDash(b.est_vdls, 0)}</td>
                <td style="padding:5px 7px;text-align:right">${b.est_annual_starts != null ? fmt(b.est_annual_starts, 1) : '—'}</td>
                <td style="padding:5px 7px;text-align:right">${b.est_annual_closings != null ? fmt(b.est_annual_closings, 1) : '—'}</td>
                <td style="padding:5px 7px;text-align:right">${money(b.avg_price)}</td>
                <td style="padding:5px 7px;text-align:right;color:#6B7B8B">${priceRange(b.min_price, b.max_price)}</td>
                <td style="padding:5px 7px;text-align:right">${numDash(b.avg_sqft, 0)}</td>
                <td style="padding:5px 7px;text-align:right">${b.avg_ppsf ? '$' + b.avg_ppsf.toFixed(0) : '—'}</td>
                <td style="padding:5px 7px;text-align:right">${numDash(b.plans, 0)}</td>
              </tr>`).join('')}
          </tbody>
        </table>` : '<div style="font-size:11px;color:#A5ADB7">No builder-by-lot-width detail returned by CBAS for this community.</div>';
      return `<div style="padding:10px;background:#F8FAFC;border-top:1px solid #DDE3E8;border-bottom:1px solid #DDE3E8">${inventory}${widthTable}${builderTable}</div>`;
    }
    body.innerHTML = `
      <div style="font-size:11px;color:#6B7B8B;margin-bottom:8px">
        <b>${esc(d.quarter_label || '')}</b> · ${d.community_count} communities (${d.active_count} actively closing)
        and ${d.builder_count} builders within ${d.radius_mi} mi of the project centre
        ${d.district_community_count ? ` · ${d.district_community_count} in ${esc(d.district_name || 'district')}` : ''}
        <span style="color:#A5ADB7">· screened from ${fmt(d.tracked_universe, 0)} CBAS-tracked communities</span>
      </div>

      <div style="display:flex;gap:6px;align-items:center;margin-bottom:4px">
        <input autocomplete="off" data-lpignore="true" data-1p-ignore id="cbas-q" type="text" placeholder="Look up any community or builder — anywhere CBAS tracks"
               onkeydown="if(event.key==='Enter')searchCbas()"
               style="flex:1;padding:6px 9px;border:1px solid #DDE3E8;border-radius:4px;font-size:12px">
        <button class="ghost" onclick="searchCbas()" style="padding:6px 12px;font-size:12px">Search</button>
      </div>
      <div id="cbas-search-results" style="margin-bottom:12px"></div>

      <h3>Competitor map</h3>
      <div id="cbas-map" style="height:420px;border:1px solid #DDE3E8;border-radius:6px;margin-bottom:6px"></div>
      <div style="font-size:11px;color:#6B7B8B;margin-bottom:12px">
        <span style="display:inline-block;width:10px;height:10px;background:#F25929;border-radius:50%;vertical-align:middle"></span> your tract
        <span style="display:inline-block;width:10px;height:10px;background:#2E7D32;border-radius:50%;vertical-align:middle;margin-left:10px"></span> active
        <span style="display:inline-block;width:10px;height:10px;background:#1976D2;border-radius:50%;vertical-align:middle;margin-left:10px"></span> future
        <span style="display:inline-block;width:10px;height:10px;background:#9AA5B1;border-radius:50%;vertical-align:middle;margin-left:10px"></span> built out
        · circle size scales with annual closings · click a marker for detail
      </div>

      <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr)">
        <div class="kpi" style="background:#FFF4ED;border:1px solid #F25929" title="A start is a builder pulling a lot and breaking ground — the transaction a lot seller competes for">
          <div class="label">Annual starts — lot demand</div>
          <div class="value" style="color:#F25929">${fmt(d.submarket_annual_starts, 0)}</div>
          <div class="label" style="margin-top:2px">trailing 12 months</div>
        </div>
        <div class="kpi"><div class="label">Annual closings</div><div class="value">${fmt(d.submarket_annual_closings, 0)}</div>
          <div class="label" style="margin-top:2px">homes delivered</div></div>
        <div class="kpi" title="75th percentile of each community's share of submarket starts">
          <div class="label">Strong-community capture</div>
          <div class="value">${(d.market_entry && d.market_entry.capture && d.market_entry.capture.share_p75_pct != null) ? d.market_entry.capture.share_p75_pct + '%' : '—'}</div>
          <div class="label" style="margin-top:2px">of submarket starts</div></div>
        <div class="kpi"><div class="label">Finished lot supply (VDL)</div><div class="value">${fmt(a.vdls, 0)}</div></div>
      </div>
      <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr);margin-top:6px">
        <div class="kpi"><div class="label">Under construction</div><div class="value">${fmt(a.under_construction, 0)}</div></div>
        <div class="kpi"><div class="label">Finished vacant</div><div class="value">${fmt(a.complete_vacant, 0)}</div></div>
        <div class="kpi"><div class="label">Future lots (pipeline)</div><div class="value">${fmt(a.futures, 0)}</div></div>
        <div class="kpi" title="Finished lots ÷ monthly closing pace — under ~6 months is a tight lot market">
          <div class="label">Months of lot supply</div>
          <div class="value" style="color:${a.months_lot_supply == null ? '#13344E' : a.months_lot_supply < 6 ? '#2E7D32' : a.months_lot_supply < 18 ? '#F59E0B' : '#C62828'}">
            ${a.months_lot_supply != null ? a.months_lot_supply.toFixed(1) : '—'}
          </div>
        </div>
      </div>
      ${(a.price_min || a.price_max) ? `
        <div style="font-size:11px;color:#6B7B8B;margin-top:6px">
          Home prices across the submarket range
          <b style="color:#13344E">$${fmt(a.price_min,0)}</b> – <b style="color:#13344E">$${fmt(a.price_max,0)}</b>
          · ${fmt(a.occupied,0)} occupied homes · ${fmt(a.total_lots,0)} total lots
        </div>` : ''}

      ${(() => {
        const qs = d.quarter_series || [];
        if (qs.length < 2) return '';
        // Grouped bars: starts vs closings per quarter — the demand trend.
        const W = 720, H = 190, PADL = 52, PADR = 12, PADT = 16, PADB = 34;
        const pw = W - PADL - PADR, ph = H - PADT - PADB;
        const maxV = Math.max(1, ...qs.map(p => Math.max(p.starts || 0, p.closings || 0)));
        const niceMax = Math.ceil(maxV / 500) * 500 || maxV;
        const slot = pw / qs.length;
        const bw = Math.min(26, slot / 3);
        const y = v => PADT + ph - (v / niceMax) * ph;
        let grid = '';
        for (let k = 0; k <= 4; k++) {
          const v = niceMax * k / 4, yy = y(v);
          grid += `<line x1="${PADL}" y1="${yy.toFixed(1)}" x2="${PADL+pw}" y2="${yy.toFixed(1)}" stroke="#E8ECEF"/>
                   <text x="${PADL-8}" y="${(yy+3.5).toFixed(1)}" text-anchor="end" font-size="10" fill="#6B7B8B">${Math.round(v).toLocaleString()}</text>`;
        }
        let bars = '';
        qs.forEach((p, i) => {
          const cx = PADL + slot * i + slot / 2;
          const s = p.starts || 0, c = p.closings || 0;
          bars += `<rect x="${(cx-bw-1).toFixed(1)}" y="${y(s).toFixed(1)}" width="${bw}" height="${(PADT+ph-y(s)).toFixed(1)}" fill="#13344E"><title>${esc(p.label||'')} starts: ${s.toLocaleString()}</title></rect>
                   <rect x="${(cx+1).toFixed(1)}" y="${y(c).toFixed(1)}" width="${bw}" height="${(PADT+ph-y(c)).toFixed(1)}" fill="#F25929"><title>${esc(p.label||'')} closings: ${c.toLocaleString()}</title></rect>
                   <text x="${cx.toFixed(1)}" y="${H-18}" text-anchor="middle" font-size="9" fill="#6B7B8B">${esc(p.label||'')}</text>`;
        });
        const last = qs[qs.length-1], prev = qs[qs.length-2];
        const dc = (last.closings||0) - (prev.closings||0);
        return `
          <h3>Quarterly trend <span style="font-weight:400;font-size:11px;color:#6B7B8B">(submarket starts vs closings)</span></h3>
          <div style="font-size:11px;margin-bottom:4px">
            <span style="display:inline-block;width:10px;height:10px;background:#13344E;vertical-align:middle"></span> starts
            <span style="display:inline-block;width:10px;height:10px;background:#F25929;vertical-align:middle;margin-left:10px"></span> closings
            <span style="margin-left:12px;color:${dc>=0?'#2E7D32':'#C62828'};font-weight:600">
              ${dc>=0?'▲':'▼'} ${Math.abs(dc).toLocaleString()} closings vs prior quarter
            </span>
          </div>
          <svg viewBox="0 0 ${W} ${H}" style="width:100%;max-width:${W}px;height:auto">${grid}${bars}</svg>`;
      })()}

      <h3>Product mix by lot width</h3>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead>
          <tr style="background:#13344E;color:#FFF">
            <th style="text-align:left;padding:6px 8px">Lot width</th>
            <th style="text-align:right;padding:6px 8px" title="Lots under control at this width">Lots</th>
            <th style="text-align:right;padding:6px 8px">Communities</th>
            <th style="text-align:right;padding:6px 8px">Builders</th>
            <th style="text-align:right;padding:6px 8px" title="Average across every floorplan offered at this width">Avg price</th>
            <th style="text-align:center;padding:6px 8px">Detail</th>
            <th style="text-align:right;padding:6px 8px">Avg SF</th>
            <th style="text-align:right;padding:6px 8px">$/SF</th>
            <th style="text-align:left;padding:6px 8px">Share of lots</th>
          </tr>
        </thead>
        <tbody>
          ${(() => {
            const bands = d.lot_bands || [];
            const maxLots = Math.max(1, ...bands.map(b => b.lots || 0));
            return bands.map(b => `
              <tr style="border-bottom:1px solid #DDE3E8;${(b.lots||0) === maxLots ? 'background:#FFF9F5' : ''}">
                <td style="padding:5px 8px"><b>${esc(b.label)}</b></td>
                <td style="padding:5px 8px;text-align:right"><b>${fmt(b.lots, 0)}</b></td>
                <td style="padding:5px 8px;text-align:right">${fmt(b.communities, 0)}</td>
                <td style="padding:5px 8px;text-align:right">${fmt(b.builders, 0)}</td>
                <td style="padding:5px 8px;text-align:right">${b.avg_price ? '$' + fmt(b.avg_price, 0) : '—'}</td>
                <td style="padding:5px 8px;text-align:right;font-size:11px;color:#6B7B8B">
                  ${b.min_price ? '$' + fmt(Math.round(b.min_price/1000),0) + 'K–$' + fmt(Math.round(b.max_price/1000),0) + 'K' : '—'}
                </td>
                <td style="padding:5px 8px;text-align:right">${fmt(b.avg_sqft, 0)}</td>
                <td style="padding:5px 8px;text-align:right">${b.avg_ppsf ? '$' + b.avg_ppsf.toFixed(0) : '—'}</td>
                <td style="padding:5px 8px">
                  <div style="background:#F25929;height:12px;border-radius:2px;width:${((b.lots||0)/maxLots*100).toFixed(0)}%;min-width:2px"></div>
                </td>
              </tr>`).join('');
          })()}
        </tbody>
      </table>
      <div style="font-size:11px;color:#6B7B8B;margin-top:4px">
        Lot widths in frontage feet. Pricing averages ${fmt((d.lot_bands||[]).reduce((s,b)=>s+(b.plans||0),0), 0)} floorplans priced in this ring.
      </div>

      ${(d.builders && d.builders.length) ? `
        <h3>Builders active around the project
          <span style="font-weight:400;font-size:11px;color:#6B7B8B">
            (${d.builder_count} within ${d.radius_mi} mi · by estimated annual starts)
          </span>
        </h3>
        <table style="width:100%;border-collapse:collapse;font-size:12px">
          <thead>
            <tr style="background:#13344E;color:#FFF">
              <th style="text-align:left;padding:6px 8px">Builder</th>
              <th style="text-align:right;padding:6px 8px" title="Community annual starts apportioned by this builder's share of lots">Est. ann. starts</th>
              <th style="text-align:right;padding:6px 8px" title="Estimated annual closings × this builder's average price">Est. sales vol.</th>
              <th style="text-align:right;padding:6px 8px" title="Lots this builder controls in the submarket — a stock, not a flow">Lots</th>
              <th style="text-align:right;padding:6px 8px">Comms</th>
              <th style="text-align:left;padding:6px 8px">Lot widths (FF)</th>
              <th style="text-align:right;padding:6px 8px">Avg price</th>
              <th style="text-align:right;padding:6px 8px">Avg SF</th>
              <th style="text-align:right;padding:6px 8px">$/SF</th>
              <th style="text-align:left;padding:6px 8px">Building in</th>
            </tr>
          </thead>
          <tbody>
            ${d.builders.map(b => `
              <tr style="border-bottom:1px solid #DDE3E8">
                <td style="padding:5px 8px">${b.named
                    ? '<b>' + esc(b.name) + '</b>'
                    : '<span style="color:#6B7B8B">' + esc(b.name) + '</span>'}</td>
                <td style="padding:5px 8px;text-align:right"><b>${fmt(b.est_annual_starts, 0)}</b></td>
                <td style="padding:5px 8px;text-align:right">${b.est_annual_volume ? '$' + fmt(Math.round(b.est_annual_volume/1e6),1) + 'M' : '—'}</td>
                <td style="padding:5px 8px;text-align:right">${fmt(b.lots, 0)}</td>
                <td style="padding:5px 8px;text-align:right">${fmt(b.communities, 0)}</td>
                <td style="padding:5px 8px;font-size:11px">${(b.lot_types_ff || []).join(', ') || '—'}</td>
                <td style="padding:5px 8px;text-align:right">${b.avg_price ? '$' + fmt(b.avg_price, 0) : '—'}</td>
                <td style="padding:5px 8px;text-align:right">${fmt(b.avg_sqft, 0)}</td>
                <td style="padding:5px 8px;text-align:right">${b.avg_ppsf ? '$' + b.avg_ppsf.toFixed(0) : '—'}</td>
                <td style="padding:5px 8px;font-size:11px;color:#6B7B8B">${(b.in || []).slice(0,3).map(esc).join(', ')}${b.communities > 3 ? ' +' + (b.communities - 3) : ''}</td>
              </tr>`).join('')}
          </tbody>
        </table>
        <div style="font-size:11px;color:#6B7B8B;margin-top:4px">
          <b>Lots</b> = lots controlled in the submarket (a stock). <b>Est. ann. starts</b> = homes begun in the trailing
          12 months (a flow), apportioned from each community's starts by that builder's lot share — CBAS reports
          starts per community, not per builder. Lots and pricing are exact.
        </div>` : ''}

      <h3>Competing communities
        <span style="font-weight:400;font-size:11px;color:#6B7B8B">
          ${d.district_name ? `in-district (${esc(d.district_name)}) first, then the wider submarket` : 'sorted by annual closings'}
        </span>
      </h3>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead>
          <tr style="background:#13344E;color:#FFF">
            <th style="text-align:left;padding:6px 8px">Community</th>
            <th style="text-align:right;padding:6px 8px" title="Straight-line distance from the project — click for driving directions">Distance</th>
            <th style="text-align:left;padding:6px 8px">Builders</th>
            <th style="text-align:left;padding:6px 8px" title="Lot widths offered, in frontage feet">Lot widths</th>
            <th style="text-align:right;padding:6px 8px">Price range</th>
            <th style="text-align:center;padding:6px 8px" title="Closings, trailing 12 months">Annual closings</th>
            <th style="text-align:right;padding:6px 8px" title="Annual closings × community price midpoint">Sales vol.</th>
            <th style="text-align:right;padding:6px 8px" title="Starts, trailing 12 months">Ann. starts</th>
            <th style="text-align:right;padding:6px 8px" title="${esc(d.quarter_label||'')} starts / closings">${esc(d.quarter_label||'Qtr')}</th>
            <th style="text-align:right;padding:6px 8px" title="Vacant developed lots — finished lot inventory">VDL</th>
            <th style="text-align:right;padding:6px 8px" title="Months of finished-lot supply at the trailing-12 closing pace">Mos. supply</th>
            <th style="text-align:right;padding:6px 8px" title="Future planned lots">Pipeline</th>
            <th style="text-align:right;padding:6px 8px" title="Total lots at full buildout, and how far along">Buildout</th>
          </tr>
        </thead>
        <tbody>
          ${(d.communities || []).map((c, i) => {
            const stColor = c.status === 'Active' ? '#2E7D32'
                          : c.status === 'Future' ? '#1976D2'
                          : c.status === 'Builtout' ? '#6B7B8B' : '#A5ADB7';
            return `
            <tr style="border-bottom:1px solid #DDE3E8;${c.in_district ? 'background:#FFF9F5' : ''}">
              <td style="padding:5px 8px"><b>${esc(c.name || '')}</b>
                ${c.status ? `<span style="color:${stColor};font-size:9px;font-weight:700;margin-left:4px">${esc(c.status.toUpperCase())}</span>` : ''}
                ${c.in_district ? '<span style="background:#1976D2;color:#FFF;padding:1px 6px;border-radius:3px;font-size:9px;font-weight:700;margin-left:4px">IN DISTRICT</span>' : ''}
                ${c.developer ? `<div style="color:#A5ADB7;font-size:10px">${esc(c.developer)}${c.city ? ' · ' + esc(c.city) : ''}</div>` : ''}
              </td>
              <td style="padding:5px 8px;text-align:right">${distLink(c.distance_mi, c.direction, c.lat, c.lon, c.name)}</td>
              <td style="padding:5px 8px;font-size:11px">
                ${(c.builders || []).slice(0,3).map(esc).join(', ') || '<span style="color:#A5ADB7">—</span>'}${(c.builders||[]).length > 3 ? ` <span style="color:#A5ADB7">+${c.builders.length - 3}</span>` : ''}
              </td>
              <td style="padding:5px 8px;font-size:11px">${(c.lot_types_ff || []).join(', ') || '<span style="color:#A5ADB7">—</span>'}</td>
              <td style="padding:5px 8px;text-align:center">
                <button class="ghost" data-cbas-detail="${i}" style="padding:2px 8px;font-size:10px">Expand</button>
              </td>
              <td style="padding:5px 8px;text-align:center">${absPill(c.annual_closings)}</td>
              <td style="padding:5px 8px;text-align:right">${c.annual_volume ? '$' + fmt(Math.round(c.annual_volume/1e6),1) + 'M' : '—'}</td>
              <td style="padding:5px 8px;text-align:right">${fmt(c.annual_starts, 0)}</td>
              <td style="padding:5px 8px;text-align:right;font-size:11px">
                ${fmt(c.starts_qtr,0)}<span style="color:#A5ADB7"> / </span>${fmt(c.closings_qtr,0)}
              </td>
              <td style="padding:5px 8px;text-align:right">${fmt(c.vdls, 0)}</td>
              <td style="padding:5px 8px;text-align:right">${c.months_lot_supply != null ? c.months_lot_supply.toFixed(1) : '—'}</td>
              <td style="padding:5px 8px;text-align:right">${fmt(c.futures, 0)}</td>
              <td style="padding:5px 8px;text-align:right">
                <b>${fmt(c.total_lots, 0)}</b> <span style="color:#A5ADB7;font-size:10px">lots</span>
                <div style="font-size:10px;color:#6B7B8B">${c.pct_built_out != null ? c.pct_built_out.toFixed(0) + '% built' : '—'}</div>
              </td>
            </tr>
            <tr id="cbas-community-detail-${i}" style="display:none">
              <td colspan="13" style="padding:0">${communityDetailHtml(c)}</td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
      <div style="font-size:11px;color:#6B7B8B;margin-top:8px">
        Source: ${esc(d.source)}. <b>VDL</b> = vacant developed lots. <b>Mos. supply</b> = VDL ÷ monthly closing pace
        (under ~6 is a tight lot market). <b>Pipeline</b> = future planned lots.
      </div>
    `;
    body.querySelectorAll('[data-cbas-detail]').forEach(btn => btn.addEventListener('click', () => {
      const row = document.getElementById(`cbas-community-detail-${btn.dataset.cbasDetail}`);
      if (!row) return;
      const open = row.style.display !== 'none';
      row.style.display = open ? 'none' : 'table-row';
      btn.textContent = open ? 'Expand' : 'Hide';
    }));
    renderCbasMap(d);
    memoPut('cbas', d);
  } catch (e) {
    body.innerHTML = `<div class="placeholder err">CBAS load failed: ${esc(e.message)}</div>`;
  }
}

// The school-district profile needs the district list and the project centroid.
// Both come from the submarket endpoint, which used to be called by the
// Submarket card - when that card was removed the profile lost its only caller
// and silently stopped appearing. This fetches the same endpoint purely for
// those two fields; nothing renders the removed card.
async function loadSchools() {
  const card = document.getElementById('schools-card');
  try {
    const r = await fetch(`/api/acq/projects/${PROJECT_ID}/submarket?radius_mi=3`);
    const d = await r.json();
    if (!r.ok || d.error) throw new Error(d.error || ('HTTP ' + r.status));
    if (d.school_districts && d.school_districts.length && d.centroid) {
      loadSchoolDistrict(d.school_districts, d.centroid);
    } else if (card) {
      // Say why rather than leaving a card that never appears.
      card.style.display = 'block';
      card.innerHTML = '<h2>School district profile</h2>'
        + '<div class="placeholder">No school district resolved for this location.</div>';
    }
  } catch (e) {
    if (card) {
      card.style.display = 'block';
      card.innerHTML = '<h2>School district profile</h2>'
        + `<div class="placeholder err">District lookup failed: ${esc(e.message)}</div>`;
    }
  }
}


async function loadSchoolDistrict(districts, centroid) {
  const card = document.getElementById('schools-card');
  const body = document.getElementById('schools-body');
  card.style.display = 'block';
  body.innerHTML = '<div class="placeholder">Pulling district enrollment, schools, and growth trend…</div>';
  try {
    const results = await Promise.all(districts.map(sd =>
      fetch(`/api/acq/school-district/${encodeURIComponent(sd.geoid)}?lat=${centroid.lat}&lon=${centroid.lon}&name=${encodeURIComponent(sd.name || '')}`)
        .then(r => r.json())
        .catch(() => ({ error: 'fetch failed', name: sd.name }))
    ));
    let html = '';
    // Memo uses the largest district by enrolment — that's the one setting the
    // pricing narrative when a tract straddles a boundary.
    const ranked = results.filter(x => !x.error)
      .sort((a, b) => (b.enrollment || 0) - (a.enrollment || 0));
    if (ranked.length) memoPut('schools', ranked[0]);
    for (const d of results) {
      if (d.error) {
        html += `<div class="placeholder err">${esc(d.name || 'District')}: ${esc(d.error)}</div>`;
        continue;
      }
      const g = d.growth || {};
      const trendPill = g.total_pct != null
        ? `<span style="background:${g.total_pct >= 0 ? '#2E7D32' : '#C62828'};color:#FFF;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:600">
             ${g.total_pct >= 0 ? '▲' : '▼'} ${Math.abs(g.total_pct).toFixed(1)}% over ${g.span_years}yr · ${g.added_students >= 0 ? '+' : ''}${fmt(g.added_students, 0)} students</span>`
        : '';

      // TEA A–F letter grade badge
      function gradeColor(r) {
        const c = { 'A': '#2E7D32', 'B': '#7CB342', 'C': '#F59E0B', 'D': '#EF6C00', 'F': '#C62828' };
        return c[(r || '').trim().charAt(0)] || '#A5ADB7';
      }
      function gradeBadge(rating, score, big) {
        if (!rating) return '<span style="color:#A5ADB7">—</span>';
        const sz = big ? 'font-size:26px;padding:6px 16px' : 'font-size:13px;padding:2px 9px';
        return `<span style="background:${gradeColor(rating)};color:#FFF;${sz};border-radius:5px;font-weight:700;display:inline-block">${esc(rating)}</span>` +
               (score != null ? `<span style="color:#6B7B8B;font-size:${big?'13':'10'}px;margin-left:5px">${score}</span>` : '');
      }
      function numOrDash(v) { return v == null ? '—' : fmt(v, 0); }

      const tea = d.tea || {};
      const warnings = d.warnings || [];
      const warningBlock = warnings.length ? `
        <div style="background:#FFF8E1;border-left:3px solid #F59E0B;border-radius:5px;padding:8px 10px;margin:8px 0;font-size:11px;color:#6B5A00">
          ${warnings.map(w => esc(w)).join('<br>')}
        </div>` : '';
      const teaBlock = tea.overall_rating ? `
        <div style="background:#F6F8FA;border:1px solid #DDE3E8;border-radius:8px;padding:12px;margin:10px 0">
          <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap">
            <div style="text-align:center">
              <div style="font-size:10px;color:#6B7B8B;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:4px">TEA rating ${esc(d.tea_year || '')}</div>
              ${gradeBadge(tea.overall_rating, tea.overall_score, true)}
            </div>
            <div style="display:flex;gap:18px;flex-wrap:wrap;font-size:11px">
              <div><div style="color:#6B7B8B;margin-bottom:3px">Student achievement</div>${gradeBadge(tea.student_achievement?.rating, tea.student_achievement?.score)}</div>
              <div><div style="color:#6B7B8B;margin-bottom:3px">School progress</div>${gradeBadge(tea.school_progress?.rating, tea.school_progress?.score)}</div>
              <div><div style="color:#6B7B8B;margin-bottom:3px">Closing the gaps</div>${gradeBadge(tea.closing_the_gaps?.rating, tea.closing_the_gaps?.score)}</div>
              <div><div style="color:#6B7B8B;margin-bottom:3px">Econ. disadvantaged</div><b style="font-size:14px">${tea.econ_disadvantaged_pct != null ? tea.econ_disadvantaged_pct.toFixed(1) + '%' : '—'}</b></div>
              <div><div style="color:#6B7B8B;margin-bottom:3px">English learners</div><b style="font-size:14px">${tea.el_students_pct != null ? tea.el_students_pct.toFixed(1) + '%' : '—'}</b></div>
            </div>
          </div>
        </div>` : '';

      // Full enrollment history chart — the migration indicator.
      const tr = d.enrollment_trend || [];
      let chart = '';
      if (tr.length >= 2) {
        const W = 720, H = 200, PADL = 52, PADR = 12, PADT = 14, PADB = 26;
        const plotW = W - PADL - PADR, plotH = H - PADT - PADB;
        const vals = tr.map(t => t.enrollment);
        const yrs  = tr.map(t => t.year);
        const yMax = Math.max(...vals), yMin = Math.min(...vals);
        // Round the axis to a clean top value
        const niceMax = Math.ceil(yMax / 1000) * 1000;
        const niceMin = Math.max(0, Math.floor(yMin / 1000) * 1000);
        const yRange = (niceMax - niceMin) || 1;
        const x = i => PADL + (i / (tr.length - 1)) * plotW;
        const y = v => PADT + plotH - ((v - niceMin) / yRange) * plotH;

        const linePts = tr.map((t, i) => `${x(i).toFixed(1)},${y(t.enrollment).toFixed(1)}`).join(' ');
        const areaPts = `${PADL},${PADT + plotH} ${linePts} ${(PADL + plotW).toFixed(1)},${PADT + plotH}`;

        // Y gridlines (4 bands)
        let grid = '';
        for (let k = 0; k <= 4; k++) {
          const v = niceMin + (yRange * k / 4);
          const yy = y(v);
          grid += `<line x1="${PADL}" y1="${yy.toFixed(1)}" x2="${PADL+plotW}" y2="${yy.toFixed(1)}" stroke="#E8ECEF" stroke-width="1"/>
                   <text x="${PADL-8}" y="${(yy+3.5).toFixed(1)}" text-anchor="end" font-size="10" fill="#6B7B8B">${Math.round(v).toLocaleString()}</text>`;
        }
        // X labels — every 5 years plus the endpoints
        let xlab = '';
        tr.forEach((t, i) => {
          const isEdge = i === 0 || i === tr.length - 1;
          if (isEdge || t.year % 5 === 0) {
            xlab += `<text x="${x(i).toFixed(1)}" y="${H-8}" text-anchor="middle" font-size="10" fill="#6B7B8B">${t.year}</text>`;
          }
        });
        // Per-point dots + value labels. Label every 5th year (and both
        // endpoints) so the numbers are readable instead of a wall of text;
        // every year still gets a dot and a native tooltip on hover.
        const lastI = tr.length - 1;
        let dots = '';
        tr.forEach((t, i) => {
          const cx = x(i), cy = y(t.enrollment);
          const isEdge = i === 0 || i === lastI;
          dots += `<circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="${isEdge ? 4 : 2.5}"
                           fill="${isEdge ? '#F25929' : '#13344E'}">
                     <title>${t.year}: ${t.enrollment.toLocaleString()} students</title>
                   </circle>`;
          if (isEdge || t.year % 5 === 0) {
            const anchor = i === lastI ? 'end' : (i === 0 ? 'start' : 'middle');
            const dx = i === lastI ? -6 : (i === 0 ? 4 : 0);
            dots += `<text x="${(cx+dx).toFixed(1)}" y="${(cy-9).toFixed(1)}" text-anchor="${anchor}"
                           font-size="${isEdge ? 11 : 9}" font-weight="${isEdge ? 700 : 500}"
                           fill="${isEdge ? '#F25929' : '#13344E'}">${t.enrollment.toLocaleString()}</text>`;
          }
        });

        // Year-by-year table so no value has to be guessed off the chart
        const rowsPerCol = Math.ceil(tr.length / 4);
        const cols = [];
        for (let c0 = 0; c0 < 4; c0++) {
          const slice = tr.slice(c0 * rowsPerCol, (c0 + 1) * rowsPerCol);
          if (!slice.length) continue;
          cols.push(`<div>${slice.map((t, k) => {
            const gi = c0 * rowsPerCol + k;
            const prev = gi > 0 ? tr[gi-1].enrollment : null;
            const chg = prev ? t.enrollment - prev : null;
            const chgTxt = chg == null ? ''
              : `<span style="color:${chg >= 0 ? '#2E7D32' : '#C62828'};font-size:10px">${chg >= 0 ? '+' : ''}${chg.toLocaleString()}</span>`;
            return `<div style="display:flex;justify-content:space-between;gap:8px;padding:1px 0">
                      <span style="color:#6B7B8B">${t.year}</span>
                      <span><b>${t.enrollment.toLocaleString()}</b> ${chgTxt}</span>
                    </div>`;
          }).join('')}</div>`);
        }

        chart = `
          <div style="margin:10px 0">
            <div style="font-size:11px;color:#6B7B8B;margin-bottom:2px;font-weight:600">
              ENROLLMENT HISTORY ${tr[0].year}–${tr[lastI].year} — migration indicator
            </div>
            <svg viewBox="0 0 ${W} ${H}" style="width:100%;max-width:${W}px;height:auto;overflow:visible">
              ${grid}
              <polygon points="${areaPts}" fill="#F25929" opacity="0.12"/>
              <polyline points="${linePts}" fill="none" stroke="#F25929" stroke-width="2.5"/>
              ${dots}${xlab}
            </svg>
            <details style="margin-top:6px">
              <summary style="font-size:11px;color:#6B7B8B;cursor:pointer;font-weight:600">
                Year-by-year enrollment (${tr.length} years, with annual change)
              </summary>
              <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;font-size:11px;margin-top:6px">
                ${cols.join('')}
              </div>
            </details>
          </div>`;
      }

      // Growth over multiple windows
      const w = (d.growth || {}).windows || {};
      function growthCell(label, win) {
        if (!win) return '';
        const up = win.total_pct >= 0;
        return `<div class="kpi">
                  <div class="label">${esc(label)} <span style="color:#A5ADB7">(${win.from_year}–${win.to_year})</span></div>
                  <div class="value" style="color:${up ? '#2E7D32' : '#C62828'};font-size:18px">${up ? '+' : ''}${win.total_pct.toFixed(1)}%</div>
                  <div class="label" style="margin-top:2px">${win.cagr_pct >= 0 ? '+' : ''}${win.cagr_pct.toFixed(2)}% CAGR · ${win.added_students >= 0 ? '+' : ''}${fmt(win.added_students, 0)} students</div>
                </div>`;
      }
      const growthGrid = (w.all_time || w['20_year'] || w['10_year'] || w['5_year']) ? `
        <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr);margin-top:6px">
          ${growthCell('All-time', w.all_time)}
          ${growthCell('20-year', w['20_year'])}
          ${growthCell('10-year', w['10_year'])}
          ${growthCell('5-year',  w['5_year'])}
        </div>` : '';

      html += `
        <div style="margin-bottom:18px">
          <h3 style="margin-bottom:6px">${esc(d.name)} ${trendPill}</h3>
          ${warningBlock}
          ${teaBlock}
          <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr)">
            <div class="kpi">
              <div class="label">Enrollment <span style="color:#A5ADB7">(${esc(d.enrollment_source || ('NCES ' + (d.nces_year || d.year || '—')))})</span></div>
              <div class="value">${numOrDash(d.enrollment)}</div>
              ${tea.students ? `<div class="label" style="margin-top:2px">TEA ${esc(d.tea_year||'')}: <b>${fmt(tea.students, 0)}</b></div>` : ''}
            </div>
            <div class="kpi"><div class="label">Schools</div><div class="value">${numOrDash(d.schools_count)}</div></div>
            <div class="kpi"><div class="label">Teachers (FTE)</div><div class="value">${numOrDash(d.teachers_fte)}</div></div>
            <div class="kpi"><div class="label">Student : teacher</div><div class="value">${d.student_teacher_ratio != null ? d.student_teacher_ratio + ':1' : '—'}</div></div>
          </div>
          ${chart}
          ${growthGrid}
          <div style="font-size:11px;color:#6B7B8B;margin-top:6px">
            ${esc(d.county || '')} · ${esc(d.city || '')}
          </div>

          <h3 style="margin-top:12px">Campuses <span style="font-weight:400;font-size:11px;color:#6B7B8B">(${(d.schools||[]).length} in ${esc(d.name)} · sorted by distance from the project)</span></h3>
          <table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead>
              <tr style="background:#13344E;color:#FFF">
                <th style="text-align:left;padding:6px 8px">Campus</th>
                <th style="text-align:center;padding:6px 8px">TEA rating</th>
                <th style="text-align:left;padding:6px 8px">Level</th>
                <th style="text-align:center;padding:6px 8px">Grades</th>
                <th style="text-align:right;padding:6px 8px">Enrollment</th>
                <th style="text-align:right;padding:6px 8px">Distance</th>
              </tr>
            </thead>
            <tbody>
              ${(d.schools || []).length ? (d.schools || []).map(s => `
                <tr style="border-bottom:1px solid #DDE3E8">
                  <td style="padding:5px 8px">
                    <b>${esc(s.name)}</b>
                  </td>
                  <td style="padding:5px 8px;text-align:center">${gradeBadge(s.tea_rating, s.tea_score)}</td>
                  <td style="padding:5px 8px;font-size:11px;color:#6B7B8B">${esc(s.level)}</td>
                  <td style="padding:5px 8px;text-align:center;font-size:11px">${esc(s.grades)}</td>
                  <td style="padding:5px 8px;text-align:right">${s.enrollment != null ? fmt(s.enrollment, 0) : '—'}</td>
                  <td style="padding:5px 8px;text-align:right">${distLink(s.distance_mi, s.direction, s.lat, s.lon, s.name)}</td>
                </tr>
              `).join('') : `
                <tr><td colspan="6" style="padding:10px 8px;color:#6B7B8B;text-align:center">
                  No NCES campus roster available for this district.
                </td></tr>`}
            </tbody>
          </table>
          <div style="font-size:11px;color:#6B7B8B;margin-top:6px;font-style:italic">
            Source: ${esc(d.source)}. TEA A–F ratings are the official Texas accountability grades; scores are 0–100.
          </div>
        </div>`;
    }
    body.innerHTML = html || '<div class="placeholder">No district data available.</div>';
  } catch (e) {
    body.innerHTML = `<div class="placeholder err">District profile failed: ${esc(e.message)}</div>`;
  }
}

window.searchNews = function () {
  const v = (document.getElementById('news-q') || {}).value || '';
  loadNews(v.trim());
};
window.clearNewsSearch = function () { loadNews(''); };

async function loadNews(customQuery) {
  const card = document.getElementById('news-card');
  const body = document.getElementById('news-body');
  card.style.display = 'block';
  const q = (customQuery || '').trim();
  body.innerHTML = '<div class="placeholder">Aggregating stories…</div>';
  const searchBar = (val, note) => `
    <div style="display:flex;gap:6px;align-items:center;margin-bottom:8px">
      <input autocomplete="off" data-lpignore="true" data-1p-ignore id="news-q" type="text" value="${esc(val)}" placeholder="Search news — an employer, corridor, competitor…"
             onkeydown="if(event.key==='Enter')searchNews()"
             style="flex:1;padding:6px 9px;border:1px solid #DDE3E8;border-radius:4px;font-size:12px">
      <button class="ghost" onclick="searchNews()" style="padding:6px 12px;font-size:12px">Search</button>
      ${val ? `<button class="ghost" onclick="clearNewsSearch()" style="padding:6px 10px;font-size:12px">Reset</button>` : ''}
    </div>
    ${note ? `<div style="font-size:11px;color:#6B7B8B;margin-bottom:8px">${note}</div>` : ''}`;
  try {
    const url = q
      ? `/api/acq/projects/${PROJECT_ID}/news?q=${encodeURIComponent(q)}`
      : `/api/acq/projects/${PROJECT_ID}/news`;
    const r = await fetch(url);
    const d = await r.json();
    if (!r.ok || d.error) {
      body.innerHTML = searchBar(q, '') +
        `<div class="placeholder err">News lookup failed: ${esc(d.error || ('HTTP ' + r.status))}</div>`;
      return;
    }
    if (!d.stories || !d.stories.length) {
      body.innerHTML = searchBar(q, '') +
        `<div class="placeholder">No stories found${q ? ' for "' + esc(q) + '"' : ''}. Google News also throttles — try again shortly.</div>`;
      return;
    }
    const searchTerms = (d.terms || []).slice(0, 6).map(esc).join(', ');
    if (!q) memoPut('news', d);      // the memo wants the submarket read, not an ad-hoc search
    body.innerHTML =
      searchBar(q, `${d.stories.length} stories · ${q
        ? 'searching: ' + esc(q)
        : `auto-derived from this site — nearest town and every school district the tract touches: <b>${searchTerms}</b>`}`) + `
      <div style="font-size:12px;max-height:500px;overflow-y:auto;border:1px solid #DDE3E8;border-radius:6px">
        ${d.stories.map(s => `
          <div style="padding:8px 10px;border-bottom:1px solid #F0F0F0">
            <a href="${esc(s.link)}" target="_blank" style="color:#1976D2;text-decoration:none;font-weight:600">${esc(s.title)}</a>
            <div style="font-size:10px;color:#6B7B8B;margin-top:2px">
              ${esc(s.source || 'Google News')} · ${esc(s.published || '')} <span style="color:#A5ADB7">· query: "${esc(s.query || '')}"</span>
            </div>
          </div>
        `).join('')}
      </div>
      <div style="font-size:11px;color:#6B7B8B;margin-top:8px;font-style:italic">
        Source: ${esc(d.source)}. Stories are relevance-searched by city/county/school-district terms, then de-duplicated by URL and sorted by recency.
      </div>
    `;
  } catch (e) {
    body.innerHTML = `<div class="placeholder err">News load failed: ${esc(e.message)}</div>`;
  }
}

async function loadFred() {
  const card = document.getElementById('fred-card');
  const body = document.getElementById('fred-body');
  card.style.display = 'block';
  body.innerHTML = '<div class="placeholder">Pulling FRED indicators (8 series, ~2 sec)…</div>';
  try {
    const r = await fetch(`/api/acq/projects/${PROJECT_ID}/fred`);
    const d = await r.json();
    if (!r.ok || d.error) {
      body.innerHTML = `<div class="placeholder">FRED unavailable: ${esc(d.error || ('HTTP ' + r.status))}<br><span style="font-size:11px">Add <code>FRED_API_KEY</code> in Railway Variables (free key at fred.stlouisfed.org) to enable.</span></div>`;
      return;
    }

    // Render each indicator as a KPI box with current value + YoY + 5yr CAGR
    function fmtValue(v, kind) {
      if (v == null) return '—';
      if (kind === 'percent')    return v.toFixed(2) + '%';
      if (kind === 'index')      return v.toFixed(1);
      if (kind === 'thousands')  return (v >= 1000 ? (v/1000).toFixed(1) + 'M' : v.toFixed(1) + 'K');
      if (kind === 'count')      return v.toLocaleString('en-US', {maximumFractionDigits: 0});
      return v.toFixed(2);
    }
    function trendBadge(pct) {
      if (pct == null) return '';
      const color = pct >= 0 ? '#2E7D32' : '#C62828';
      const arrow = pct >= 0 ? '▲' : '▼';
      return `<span style="color:${color};font-size:11px;font-weight:600">${arrow} ${Math.abs(pct).toFixed(1)}%</span>`;
    }
    function miniSparkline(data, width=80, height=20) {
      if (!data || data.length < 2) return '';
      const vals = data.map(d => d.value).filter(v => v != null);
      if (vals.length < 2) return '';
      const min = Math.min(...vals), max = Math.max(...vals);
      const range = max - min || 1;
      const step = width / (vals.length - 1);
      const points = vals.map((v, i) => `${(i * step).toFixed(1)},${(height - ((v - min) / range) * height).toFixed(1)}`).join(' ');
      return `<svg width="${width}" height="${height}" style="vertical-align:middle"><polyline points="${points}" fill="none" stroke="#F25929" stroke-width="1.5"/></svg>`;
    }

    let html = '<div class="kpi-grid" style="grid-template-columns:1fr 1fr">';
    for (const ind of d.indicators) {
      const x = ind.data || {};
      if (x.error) {
        html += `<div class="kpi"><div class="label">${esc(ind.label)}</div><div class="value err" style="font-size:11px">err: ${esc(x.error)}</div></div>`;
        continue;
      }
      html += `
        <div class="kpi">
          <div class="label">${esc(ind.label)} <span style="color:#A5ADB7;font-size:9px">(${ind.series_id})</span></div>
          <div class="value" style="display:flex;align-items:center;justify-content:space-between;gap:8px">
            <span>${fmtValue(x.current, ind.format)}</span>
            ${miniSparkline(x.sparkline)}
          </div>
          <div class="label" style="margin-top:2px;display:flex;justify-content:space-between">
            <span>YoY ${x.yoy_pct != null ? trendBadge(x.yoy_pct) : '—'}</span>
            <span>5yr ${x.five_year_total_pct != null ? trendBadge(x.five_year_total_pct) : '—'}${x.five_year_cagr_pct != null ? ' (' + x.five_year_cagr_pct.toFixed(1) + '% CAGR)' : ''}</span>
          </div>
        </div>`;
    }
    html += '</div>';
    html += `<div style="font-size:11px;color:#6B7B8B;margin-top:8px">${esc(d.source)} · ${esc(d.msa)}. Click any series ID at fred.stlouisfed.org for the full history + methodology.</div>`;
    body.innerHTML = html;
  } catch (e) {
    body.innerHTML = `<div class="placeholder err">FRED load failed: ${esc(e.message)}</div>`;
  }
}

loadProject();
