/* Market evidence for pricing decisions in MPC Underwriting.
 *
 * A model created from an acquisition carries the submarket read with it, in
 * inputs._acq_link.market. Pace, yield and acreage were measured and came
 * across as values; price was not. A price is a judgement about a market and
 * this model prices thousands of lots off it, so the number stays the
 * underwriter's. These charts put the evidence in front of that decision and
 * nothing more: what each width is achieving, which builders make up the
 * average, and how wide the range behind it is.
 *
 * Chart colours are the two that pass the palette checks against a light
 * surface — #2a78d6 for a measured value, #F25929 for a reference. The brand
 * navy fails as a categorical hue (too dark, reads grey), so it is used for
 * ink here and never to carry identity.
 *
 * Builders are identified by their own axis label rather than by colour, so
 * ten builders need ten labels instead of ten hues — identity never rests on
 * colour alone, and the chart does not degrade as a market gets busier.
 */
(function () {
  'use strict';

  const BLUE = '#2a78d6';       // a measured value
  const BLUE_SOFT = 'rgba(42,120,214,0.22)';
  const ACCENT = '#F25929';     // a reference: the average, the suggestion
  const INK = '#13344E';
  const MUTED = '#93A0AC';
  const GRID = 'rgba(19,52,78,0.08)';

  let charts = [];

  const money = (n) => (n === null || n === undefined || !isFinite(n))
    ? '—' : '$' + Math.round(n).toLocaleString();
  const money1 = (n) => (n === null || n === undefined || !isFinite(n))
    ? '—' : '$' + Number(n).toLocaleString(undefined, { maximumFractionDigits: 0 });

  function market() {
    // The page keeps its inputs in a module-scoped `let`; _lastInputs is the
    // handle it puts on window, refreshed every time a project loads.
    try {
      const inp = window._lastInputs || window.inputs;
      return (inp && inp._acq_link && inp._acq_link.market) || null;
    } catch (e) { return null; }
  }

  function destroyCharts() {
    charts.forEach(c => { try { c.destroy(); } catch (e) {} });
    charts = [];
  }

  /* Values are printed on the marks rather than left to be read off a scale.
   * Carlos's standing note on charts: otherwise we have to guess. */
  const printValues = {
    id: 'printValues',
    afterDatasetsDraw(chart, _args, opts) {
      const { ctx } = chart;
      ctx.save();
      ctx.font = '600 10px Inter, system-ui, sans-serif';
      ctx.fillStyle = INK;
      ctx.textBaseline = 'middle';
      chart.data.datasets.forEach((ds, di) => {
        if (ds._noLabels) return;
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        meta.data.forEach((el, i) => {
          const raw = ds.data[i];
          if (raw === null || raw === undefined) return;
          const v = Array.isArray(raw) ? raw[1] : raw;
          if (!isFinite(v)) return;
          ctx.textAlign = 'left';
          ctx.fillText((opts && opts.fmt ? opts.fmt(v, i, ds) : money(v)), el.x + 7, el.y);
        });
      });
      ctx.restore();
    }
  };

  function lotPriceChart(canvas, mk) {
    const rows = (mk.bands || []).filter(w => w.implied_lot_ff);
    if (!rows.length) return null;
    const ratioPct = Math.round((mk.lot_ratio || 0.22) * 100);
    const c = new Chart(canvas, {
      type: 'bar',
      plugins: [printValues],
      data: {
        labels: rows.map(w => w.label + (w.in_mix ? '' : '  (not in mix)')),
        datasets: [{
          label: 'Implied lot $/FF',
          data: rows.map(w => w.implied_lot_ff),
          backgroundColor: rows.map(w => w.in_mix ? BLUE : BLUE_SOFT),
          borderRadius: 4,
          borderSkipped: 'start',
          barThickness: 16,
        }],
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        layout: { padding: { right: 64 } },
        plugins: {
          legend: { display: false },     // one series; the title names it
          printValues: { fmt: v => '$' + Number(v).toFixed(0) + '/FF' },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const w = rows[ctx.dataIndex];
                return [
                  'Implied lot $/FF: $' + w.implied_lot_ff.toFixed(0),
                  'Lot value: ' + money(w.implied_lot_value)
                    + ' = ' + money(w.avg_price) + ' x ' + ratioPct + '%',
                  'Priced off ' + w.mid_ff + ' FF, the average frontage built here'
                    + ((w.widths || []).length > 1
                       ? ' (' + w.widths.join(', ') + ' FF)' : ''),
                  w.lots ? w.lots.toLocaleString() + ' lots across '
                    + w.communities + ' communities' : '',
                ].filter(Boolean);
              }
            }
          },
        },
        scales: {
          x: { title: { display: true, text: 'Implied finished lot $ per front foot',
                        color: MUTED, font: { size: 10 } },
               grid: { color: GRID }, border: { display: false },
               ticks: { color: MUTED, font: { size: 10 },
                        callback: v => '$' + v } },
          y: { grid: { display: false }, border: { display: false },
               ticks: { color: INK, font: { size: 11, weight: '600' } } },
        },
      },
    });
    return c;
  }

  function homePriceChart(canvas, mk) {
    /* One row per builder per width: a bar spanning the builder's min-max and
     * a dot at their average. The spread is the point — a width whose range
     * is $180k wide is not really one price, and an average over two builders
     * is not the same claim as an average over nine. */
    const labels = [], ranges = [], avgs = [], meta = [];
    (mk.bands || []).forEach(w => {
      if (!w.avg_price) return;
      labels.push(w.label + ' — market');
      ranges.push((w.min_price && w.max_price) ? [w.min_price, w.max_price] : null);
      avgs.push(w.avg_price);
      meta.push({ kind: 'market', w, name: 'All builders', lots: w.lots });
      (w.builders || []).forEach(b => {
        labels.push('      ' + b.name);
        ranges.push((b.min_price && b.max_price) ? [b.min_price, b.max_price] : null);
        avgs.push(b.avg_price);
        meta.push({ kind: 'builder', w, name: b.name, lots: b.lots,
                    sqft: b.avg_sqft, plans: b.plans });
      });
    });
    if (!labels.length) return null;

    return new Chart(canvas, {
      type: 'bar',
      plugins: [printValues],
      data: {
        labels,
        datasets: [
          { label: 'Price range', data: ranges, _noLabels: true,
            backgroundColor: meta.map(m => m.kind === 'market'
              ? 'rgba(242,89,41,0.18)' : BLUE_SOFT),
            borderRadius: 3, barThickness: 9, order: 2 },
          { label: 'Average', data: avgs, type: 'scatter',
            backgroundColor: meta.map(m => m.kind === 'market' ? ACCENT : BLUE),
            borderColor: '#FFFFFF', borderWidth: 1.5,
            pointRadius: meta.map(m => m.kind === 'market' ? 6 : 5),
            pointHoverRadius: 8, order: 1 },
        ],
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        layout: { padding: { right: 86 } },
        plugins: {
          legend: { display: true, position: 'top', align: 'end',
                    labels: { color: MUTED, boxWidth: 10, font: { size: 10 },
                              usePointStyle: true } },
          printValues: { fmt: (v) => money1(v) },
          tooltip: {
            callbacks: {
              title: (items) => {
                const m = meta[items[0].dataIndex];
                return m.w.label + ' — ' + m.name;
              },
              label: (ctx) => {
                const m = meta[ctx.dataIndex];
                const r = ranges[ctx.dataIndex];
                return [
                  'Average: ' + money(avgs[ctx.dataIndex]),
                  r ? ('Range: ' + money(r[0]) + ' – ' + money(r[1])) : '',
                  m.lots ? (m.lots.toLocaleString() + ' lots') : '',
                  m.sqft ? (m.sqft.toLocaleString() + ' sf avg') : '',
                  m.plans ? (m.plans + ' plans') : '',
                ].filter(Boolean);
              }
            }
          },
        },
        scales: {
          x: { title: { display: true, text: 'New-home price in the submarket',
                        color: MUTED, font: { size: 10 } },
               grid: { color: GRID }, border: { display: false },
               ticks: { color: MUTED, font: { size: 10 },
                        callback: v => '$' + (v / 1000) + 'k' } },
          y: { grid: { display: false }, border: { display: false },
               ticks: { color: INK, font: { size: 10 }, autoSkip: false } },
        },
      },
    });
  }

  function open(focus) {
    const mk = market();
    const overlay = document.getElementById('uw-market-modal');
    const body = document.getElementById('uw-market-body');
    if (!overlay || !body) return;
    destroyCharts();

    if (!mk || !(mk.bands || []).length) {
      body.innerHTML =
        '<div class="notice info" style="margin:0">No submarket read is attached to this '
        + 'model. It is carried over when a model is created from an acquisition project '
        + 'with a completed analysis — open the acquisition and use <b>Underwrite this</b>.'
        + '</div>';
      overlay.classList.add('open');
      return;
    }

    const ratioPct = Math.round((mk.lot_ratio || 0.22) * 100);
    const sug = mk.suggested_price_per_ff;
    const ring = mk.ring || {};
    const comms = mk.communities || [];
    const esc = (t) => String(t == null ? '' : t)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const num = (n) => (n == null || !isFinite(n)) ? '—'
      : Number(n).toLocaleString(undefined, { maximumFractionDigits: 1 });
    const range$ = (c) => (c.price_min && c.price_max)
      ? (money1(c.price_min) + '–' + money1(c.price_max))
      : (c.price_min ? money1(c.price_min) : '—');
    // Carlos's standing rule: a distance on screen opens driving directions.
    const dist = (c) => {
      const label = (c.distance_mi != null ? c.distance_mi.toFixed(1) + ' mi' : '—')
        + (c.direction ? ' ' + esc(c.direction) : '');
      if (c.lat == null || c.lon == null) return label;
      const dest = c.lat + ',' + c.lon;
      return `<a href="https://www.google.com/maps/dir/?api=1&destination=${dest}"
                 target="_blank" rel="noopener" style="color:${BLUE};text-decoration:none"
                 title="Driving directions">${label}</a>`;
    };
    const nH = Math.max(220, (mk.bands || []).reduce(
      (n, w) => n + 1 + (w.builders || []).length, 0) * 20 + 70);
    const cap = mk.capture || {};

    body.innerHTML = `
      <div class="notice info" style="margin:0 0 14px">
        The submarket as CBAS reported it when this model was created. Pace, yield and
        acreage came across as values because they were measured; <b>price is a judgement
        and stays yours</b> — nothing here writes a number into the model unless you press
        the button below.
      </div>

      ${ring.radius_mi ? `
      <div class="section-header" style="margin-top:0">How this submarket is measured</div>
      <div style="font-size:11.5px;color:${MUTED};margin:-4px 0 10px;line-height:1.55">
        Every figure below comes from CBAS communities within
        <b style="color:${INK}">${ring.radius_mi} miles</b> of the project centroid, by true
        great-circle distance from each community's own coordinates — not a ZIP or county
        approximation. ${ring.district_name
          ? `Communities in <b style="color:${INK}">${esc(ring.district_name)}</b> are listed
             first and flagged, because school district is the submarket boundary that
             matters. ` : ''}${ring.quarter_label ? `CBAS quarter ${esc(ring.quarter_label)}. ` : ''}
        Only tracked communities appear — large communities and MPCs, not every subdivision.
      </div>
      <div style="display:flex;gap:16px;flex-wrap:wrap;margin:0 0 16px;padding:10px 12px;
                  background:#F7F9FA;border:1px solid #E5E8EC;border-radius:8px;font-size:11.5px">
        <div><b>${ring.community_count ?? '—'}</b> communities
          <span style="color:${MUTED}">· ${ring.active_count ?? '—'} active</span></div>
        <div><b>${ring.builder_count ?? '—'}</b> builders</div>
        ${ring.months_lot_supply != null ? `<div><b>${ring.months_lot_supply}</b> months of lot supply
          <span style="color:${MUTED}">${ring.lot_market ? '· ' + ring.lot_market : ''}</span></div>` : ''}
        ${ring.years_of_pipeline != null ? `<div><b>${ring.years_of_pipeline}</b> yrs of pipeline</div>` : ''}
        ${ring.dominant_product ? `<div>Dominant product <b>${esc(ring.dominant_product)}</b></div>` : ''}
        ${ring.top3_start_share_pct != null ? `<div>Top 3 builders <b>${ring.top3_start_share_pct}%</b>
          <span style="color:${MUTED}">of starts${ring.builder_concentration ? ' · ' + esc(ring.builder_concentration) : ''}</span></div>` : ''}
      </div>` : ''}

      ${comms.length ? `
      <div class="section-header">Communities nearby <span style="font-weight:400;color:${MUTED};
        text-transform:none;letter-spacing:0">${comms.length} within ${ring.radius_mi || '—'} mi</span></div>
      <div style="overflow-x:auto;margin-bottom:6px">
        <table style="width:100%;border-collapse:collapse;font-size:11px">
          <thead><tr style="text-align:left;color:${MUTED};border-bottom:1px solid #E5E8EC">
            <th style="padding:5px 7px;font-weight:600">Community</th>
            <th style="padding:5px 7px;font-weight:600">Dist</th>
            <th style="padding:5px 7px;font-weight:600">Status</th>
            <th style="padding:5px 7px;font-weight:600">Lot widths</th>
            <th style="padding:5px 7px;font-weight:600;text-align:right">Price range</th>
            <th style="padding:5px 7px;font-weight:600;text-align:right">Starts/yr</th>
            <th style="padding:5px 7px;font-weight:600;text-align:right">Closings/yr</th>
            <th style="padding:5px 7px;font-weight:600;text-align:right">VDL</th>
            <th style="padding:5px 7px;font-weight:600;text-align:right">Future</th>
            <th style="padding:5px 7px;font-weight:600;text-align:right">MoS</th>
            <th style="padding:5px 7px;font-weight:600;text-align:right">Built</th>
          </tr></thead>
          <tbody>${comms.map(c => `
            <tr style="border-bottom:1px solid #F1F4F6">
              <td style="padding:5px 7px">
                <span style="font-weight:600;color:${INK}">${esc(c.name || '—')}</span>
                ${c.in_district ? `<span style="margin-left:5px;font-size:9px;font-weight:700;
                   background:#E8F1EA;color:#1C6B47;padding:1px 5px;border-radius:3px">IN DISTRICT</span>` : ''}
                ${c.developer ? `<div style="color:${MUTED};font-size:10px">${esc(c.developer)}</div>` : ''}
              </td>
              <td style="padding:5px 7px;white-space:nowrap">${dist(c)}</td>
              <td style="padding:5px 7px;color:${MUTED}">${esc(c.status || '—')}</td>
              <td style="padding:5px 7px;white-space:nowrap">${esc(c.lot_type_range || '—')}
                ${c.builder_count ? `<span style="color:${MUTED}">· ${c.builder_count} blt</span>` : ''}</td>
              <td style="padding:5px 7px;text-align:right;white-space:nowrap">${range$(c)}</td>
              <td style="padding:5px 7px;text-align:right">${num(c.annual_starts)}</td>
              <td style="padding:5px 7px;text-align:right">${num(c.annual_closings)}</td>
              <td style="padding:5px 7px;text-align:right">${num(c.vdls)}</td>
              <td style="padding:5px 7px;text-align:right">${num(c.futures)}</td>
              <td style="padding:5px 7px;text-align:right">${num(c.months_lot_supply)}</td>
              <td style="padding:5px 7px;text-align:right">${c.pct_built_out != null
                 ? Math.round(c.pct_built_out) + '%' : '—'}</td>
            </tr>`).join('')}</tbody>
        </table>
      </div>
      <div style="font-size:10.5px;color:${MUTED};margin:0 0 20px">
        VDL = vacant developed lots. MoS = months of lot supply at the community's own
        closing rate. Distances open driving directions.
      </div>` : ''}

      ${cap.addressable_starts ? `
      <div style="display:flex;gap:18px;flex-wrap:wrap;margin:0 0 14px;padding:10px 12px;
                  background:#F7F9FA;border:1px solid #E5E8EC;border-radius:8px;font-size:11.5px">
        <div><b>${Math.round(cap.ring_annual_starts || 0)}</b> starts/yr in the ring
          <span style="color:${MUTED}">across ${cap.active_communities || '—'} active communities</span></div>
        <div><b>${Math.round(cap.addressable_starts)}</b> addressable
          <span style="color:${MUTED}">in the widths you target</span></div>
        <div>Median community share <b>${cap.share_median_pct ?? '—'}%</b>
          <span style="color:${MUTED}">· 75th ${cap.share_p75_pct ?? '—'}%</span></div>
      </div>` : ''}

      <div class="section-header" style="margin-top:0">Implied finished lot $/FF by lot width</div>
      <div style="font-size:11px;color:${MUTED};margin:-4px 0 8px">
        Finished lot taken at <b>${ratioPct}%</b> of the average new-home price at that
        width, divided by the average frontage actually built there. Widths are grouped
        the way builders talk about them — the 40s, the 50s — with everything under 40
        and over 90 in one bucket each. Widths your product mix does not use are faded.
      </div>
      <div style="height:${Math.max(170, (mk.bands || []).length * 32 + 60)}px">
        <canvas id="uw-mk-ff"></canvas>
      </div>

      ${sug ? `
      <div style="display:flex;align-items:center;gap:12px;margin-top:12px;padding:10px 12px;
                  background:#FEF4EF;border:1px solid #F5B79E;border-radius:8px">
        <div style="flex:1;font-size:12px;color:${INK}">
          <b>Blended across your mix: $${Number(sug).toFixed(0)}/FF.</b>
          <span style="color:#6B7B8B">${mk.suggested_basis || ''}</span>
        </div>
        <button id="uw-mk-apply" class="btn accent"
                title="Writes this into Year 0 of the $/FF table. Later years and escalation stay yours.">
          Use $${Number(sug).toFixed(0)} for Year 0
        </button>
      </div>` : ''}

      <div class="section-header" style="margin-top:22px">New-home price by lot width and builder</div>
      <div style="font-size:11px;color:${MUTED};margin:-4px 0 8px">
        The bar is each builder's min-to-max; the dot is their average. The orange row
        per width is the market as a whole. A wide bar means that width is not really one
        price. Builders active in several communities are merged, weighted by lot count.
      </div>
      <div style="height:${nH}px"><canvas id="uw-mk-home"></canvas></div>
      <div style="font-size:11px;color:${MUTED};margin-top:10px">
        Home price per lot width drives assessed value through the <b>AV %</b> already set
        on each row of the lot table, which is what carries into MUD capacity. Type the
        figure you want to underwrite into <b>Home Price</b> on that table.
      </div>`;

    overlay.classList.add('open');

    // Charts size off their container, so they are built after it is visible.
    requestAnimationFrame(() => {
      const a = lotPriceChart(document.getElementById('uw-mk-ff'), mk);
      const b = homePriceChart(document.getElementById('uw-mk-home'), mk);
      charts = [a, b].filter(Boolean);
      if (focus === 'home') {
        const el = document.getElementById('uw-mk-home');
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    });

    const applyBtn = document.getElementById('uw-mk-apply');
    if (applyBtn) {
      applyBtn.addEventListener('click', () => {
        // One deliberate press, Year 0 only. Escalation across later years is a
        // separate judgement and is left alone.
        const el = document.getElementById('ri-ppff-0');
        if (!el) return;
        el.value = Math.round(sug);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        applyBtn.textContent = 'Applied to Year 0';
        applyBtn.disabled = true;
      });
    }
  }

  function close() {
    destroyCharts();
    const overlay = document.getElementById('uw-market-modal');
    if (overlay) overlay.classList.remove('open');
  }

  window.openUwMarket = open;
  window.closeUwMarket = close;

  /* The button is always there. Hiding it on models without a market read
   * made the feature impossible to find -- you had to already know it existed
   * to go looking for it. Dimmed it still reads as available, and the popup's
   * empty state explains how to get the data. */
  window.syncUwMarketButtons = function () {
    const has = !!(market() && (market().bands || []).length);
    document.querySelectorAll('.uw-market-btn').forEach(b => {
      b.style.opacity = has ? '' : '0.55';
      b.title = has
        ? 'What the submarket shows about pricing — evidence, not a value'
        : 'No submarket read on this model yet — open it to see how to attach one';
    });
  };
})();
