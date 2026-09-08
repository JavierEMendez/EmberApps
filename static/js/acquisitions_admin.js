/* Parcel-cache admin, ported from the standalone app's admin page.
 * Static file, so Jinja never parses it.
 */
// Declared at module scope on the page this came from; the extraction took
// the functions but not the state they close over.
let _cachePollTimer = null;
async function loadCacheStatus() {
  try {
    const r = await fetch('/api/acq/cache/status');
    const d = await r.json();
    const body = document.getElementById('cache-body');
    body.innerHTML = (d.counties || []).map(c => {
      const statusColors = {
        fresh:   'background:#DEF7E5;color:#1E6A3A',
        stale:   'background:#FFF3CD;color:#856404',
        partial: 'background:#FFE0B2;color:#8A4B08',
        loading: 'background:#CCE5FF;color:#004085',
        pending: 'background:#F1F2F3;color:#58595B',
        error:   'background:#F8D7DA;color:#721C24',
      };
      const style = statusColors[c.status] || statusColors.pending;
      const refreshed = c.last_refreshed_at
        ? new Date(c.last_refreshed_at * 1000).toLocaleString()
        : '—';
      const age = c.age_days != null ? ` <span style="color:#6B7B8B">(${c.age_days}d ago)</span>` : '';

      // Tile progress for a load in flight. This comes from the cache_tiles
      // table rather than the loader's in-memory status, because under gunicorn
      // the poll usually lands on a worker that is not the one bootstrapping --
      // so a running load showed LOADING and nothing else. Staleness is worth
      // showing too: a load whose last tile landed long ago has probably died,
      // and that reads identically to a slow one otherwise.
      let tileNote = '';
      if (c.status === 'loading' || (c.status === 'partial' && c.tiles_done)) {
        const since = c.last_tile_at ? Math.round(Date.now() / 1000 - c.last_tile_at) : null;
        const quiet = since == null ? '' :
          since < 90 ? '' :
          since < 900 ? ` · last tile ${Math.round(since / 60)}m ago` :
          ` · <span style="color:#C62828">no tile in ${Math.round(since / 60)}m</span>`;
        tileNote = `<div style="font-size:10px;color:#6B7B8B;margin-top:3px"
             title="Tiles committed so far. Each one is durable — an interrupted load resumes from here rather than starting over.">
             ${c.tiles_done || 0} tiles done${quiet}</div>`;
      }
      return `
        <tr>
          <td><b>${c.county_name}</b><br><span style="color:#9a9a9a;font-size:11px">FIPS ${c.county_fips}</span></td>
          <td><span style="${style};padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;text-transform:uppercase">${c.status}</span>${tileNote}</td>
          <td>${c.parcel_count.toLocaleString()}
            ${c.missing ? `<div style="font-size:10px;color:#C62828;font-weight:600;margin-top:2px"
                 title="This bootstrap reported inserting ${c.reported_inserted.toLocaleString()} parcels but only ${c.parcel_count.toLocaleString()} are in the table. Re-bootstrap this county.">
                 ${c.missing.toLocaleString()} missing</div>` : ''}
          </td>
          <td>${refreshed}${age}</td>
          <td><button class="ghost" data-cache-county="${c.county_fips}" data-cache-name="${c.county_name}">Bootstrap / refresh</button></td>
        </tr>`;
    }).join('');

    // Wire per-county buttons
    body.querySelectorAll('button[data-cache-county]').forEach(btn => {
      btn.addEventListener('click', () => triggerBootstrap([btn.dataset.cacheCounty], btn.dataset.cacheName));
    });

    // In-progress banner.
    //
    // `in_progress` comes from a dict held in one process's memory, so under
    // gunicorn this poll usually reaches a worker that is not the one running
    // the bootstrap and reports nothing. That used to hide the banner AND stop
    // the poll timer, so the page froze with the county stuck on LOADING and no
    // further updates. A county marked 'loading' in the table is the reliable
    // signal, and it works from any worker.
    const ip = d.in_progress || {};
    const loadingCounties = (d.counties || []).filter(c => c.status === 'loading');
    const active = ip.running || loadingCounties.length > 0;
    const prog = document.getElementById('cache-progress');
    if (active) {
      prog.style.display = 'block';
      if (ip.running) {
        prog.innerHTML = `<b>Bootstrapping ${ip.county || '…'}</b> — ${ip.pct}% — ${ip.msg || ''}`;
      } else {
        // Reached a worker that is not the one loading; report from the table.
        const names = loadingCounties.map(c => `${c.county_name} (${c.tiles_done || 0} tiles)`).join(', ');
        prog.innerHTML = `<b>Bootstrapping ${names}</b> — progress from the parcel table; ` +
                         `the worker handling this request is not the one loading, so there is no live percentage.`;
      }
      if (!_cachePollTimer) {
        _cachePollTimer = setInterval(loadCacheStatus, 5000);
      }
    } else {
      prog.style.display = 'none';
      if (_cachePollTimer) { clearInterval(_cachePollTimer); _cachePollTimer = null; }
    }

    // Orphaned spatial-index entries. Searches stay correct with these present
    // -- the query inner-joins parcels to the index -- but each one walks the
    // dead entries first, so a bloated index is slow, not wrong.
    const orph = document.getElementById('cache-orphans');
    const n = d.rtree_orphans;
    if (n && n > 50000) {
      const total = (d.counties || []).reduce((a, c) => a + (c.parcel_count || 0), 0);
      const mult = total ? ((total + n) / total).toFixed(1) : '?';
      orph.style.display = 'block';
      orph.innerHTML =
        '<b>Spatial index carries ' + n.toLocaleString() + ' dead entries</b> — it is ' +
        mult + '× the size of the data it indexes, which slows every search. ' +
        'Results stay correct either way. ' +
        '<button class="ghost" id="btn-cache-vacuum" style="margin-left:8px">Vacuum index</button>';
      document.getElementById('btn-cache-vacuum').addEventListener('click', vacuumIndex);
    } else {
      orph.style.display = 'none';
    }

    // Parcels the spatial index cannot SEE — the opposite problem, and the
    // damaging one. An orphan makes searches slower; an unindexed parcel makes
    // them wrong: the county count reads correct and every result is short.
    // Any number above zero is worth acting on, so this has no threshold.
    const un = document.getElementById('cache-unindexed');
    const u = d.rtree_unindexed || 0;
    if (un) {
      if (u > 0) {
        un.style.display = 'block';
        un.innerHTML =
          '<b>' + u.toLocaleString() + ' parcels are missing from the spatial index</b> — ' +
          'they are in the cache and count toward the totals below, but no search ' +
          'can return them. Rebuilt from stored geometry; nothing is re-downloaded. ' +
          '<button class="ghost" id="btn-cache-reindex" style="margin-left:8px">Rebuild index</button>';
        document.getElementById('btn-cache-reindex').addEventListener('click', reindexParcels);
      } else {
        un.style.display = 'none';
      }
    }
  } catch (e) {
    document.getElementById('cache-body').innerHTML = `<tr><td colspan="5" style="color:#c62828">Failed to load cache status: ${e.message}</td></tr>`;
  }
}

async function triggerBootstrap(counties, label) {
  const payload = counties === 'all' ? { counties: 'all' } : { counties: counties };
  const what = counties === 'all' ? 'all Houston-metro counties' : label;
  if (!confirm(`Bootstrap ${what}? This pulls fresh parcel data from StratMap into the local cache. Estimated 3–10 min per county.`)) return;
  try {
    const r = await fetch('/api/acq/cache/bootstrap', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const d = await r.json();
    if (!r.ok) {
      alert(d.error || 'Bootstrap failed');
      return;
    }
    loadCacheStatus();
  } catch (e) {
    alert('Failed to start bootstrap: ' + e.message);
  }
}

// Bind only what is present. The card this came from carried other
// controls (a backup restore, among them) that this page does not, and an
// addEventListener on a missing element threw before loadCacheStatus ever
// ran — a populated cache rendered as an empty table.
function on(id, ev, fn) {
  var el = document.getElementById(id);
  if (el) el.addEventListener(ev, fn);
}
async function reindexParcels() {
  const el = document.getElementById('cache-unindexed');
  if (el) el.innerHTML = 'Rebuilding index entries…';
  try {
    const r = await fetch('/api/acq/cache/reindex', { method: 'POST' });
    const d = await r.json();
    if (d.error) throw new Error(d.error);
    if (el) el.innerHTML = 'Restored <b>' + (d.reindexed || 0).toLocaleString() +
      '</b> index entries. Searches will return them now.';
    setTimeout(loadCacheStatus, 800);
  } catch (e) {
    if (el) el.innerHTML = 'Rebuild failed: ' + e.message;
  }
}

on('btn-cache-refresh-status', 'click', loadCacheStatus);
on('btn-cache-bootstrap-all', 'click', function () { triggerBootstrap('all'); });
loadCacheStatus();

// Pull the appraisal district's own acreage. StratMap draws some accounts as
// the whole abstract rather than the tract -- 3.5% of sampled Harris parcels
// over 90 acres, always oversized, up to eight times -- so until this has run
// the acreage shown for those is the polygon's, not the parcel's.
on('btn-reconcile-cad', 'click', async (ev) => {
  const b = ev.currentTarget;
  const label = b.textContent;
  if (!confirm('Pull HCAD acreage for every Harris parcel over 5 acres?'
             + '\nAbout 56,000 parcels; takes a few minutes.')) return;
  b.disabled = true;
  b.textContent = 'Reconciling...';
  try {
    const r = await fetch('/api/acq/cache/reconcile-cad?county_fips=48201&min_acres=5',
                          { method: 'POST' });
    const d = await r.json();
    if (!r.ok || d.error) throw new Error(d.error || ('HTTP ' + r.status));
    alert('Reconciled ' + d.matched.toLocaleString() + ' of '
        + d.considered.toLocaleString() + ' parcels.\n'
        + d.disagreeing.toLocaleString() + ' disagree with StratMap by more than '
        + '20% and now report the district acreage.');
    loadCacheStatus();
  } catch (e) {
    alert('Reconcile failed: ' + e.message);
  } finally {
    b.disabled = false;
    b.textContent = label;
  }
});
