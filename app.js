(function () {
  "use strict";

  const M = window.MATRIZ;
  const USUARIOS = window.USUARIOS || [];
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  // Resolución de las imágenes que se copian (fichas, cuadro de resultado y reporte de producto):
  // 4 píxeles por cada punto de pantalla, para que sigan nítidas al agrandarlas o hacer zoom en Word u Outlook.
  const ESCALA_COPIA = 4;
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* sin almacenamiento */ } },
  };
  const sess = {
    get(k) { try { return sessionStorage.getItem(k); } catch { return null; } },
    set(k, v) { try { sessionStorage.setItem(k, v); } catch { /* sin almacenamiento */ } },
    del(k) { try { sessionStorage.removeItem(k); } catch { /* sin almacenamiento */ } },
  };

  // ---------- SHA-256 (crypto.subtle con respaldo en JS puro) ----------
  function sha256Fallback(str) {
    const bytes = new TextEncoder().encode(str);
    const K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
    let H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
    const l = bytes.length, n = ((l + 9 + 63) >> 6) << 6, buf = new Uint8Array(n);
    buf.set(bytes); buf[l] = 0x80;
    const dv = new DataView(buf.buffer); dv.setUint32(n - 4, l * 8); dv.setUint32(n - 8, Math.floor(l / 0x20000000));
    const w = new Uint32Array(64), r = (x, k) => (x >>> k) | (x << (32 - k));
    for (let o = 0; o < n; o += 64) {
      for (let i = 0; i < 16; i++) w[i] = dv.getUint32(o + i * 4);
      for (let i = 16; i < 64; i++) {
        const s0 = r(w[i-15], 7) ^ r(w[i-15], 18) ^ (w[i-15] >>> 3), s1 = r(w[i-2], 17) ^ r(w[i-2], 19) ^ (w[i-2] >>> 10);
        w[i] = (w[i-16] + s0 + w[i-7] + s1) | 0;
      }
      let [a, b, c, d, e, f, g, h] = H;
      for (let i = 0; i < 64; i++) {
        const t1 = (h + (r(e, 6) ^ r(e, 11) ^ r(e, 25)) + ((e & f) ^ (~e & g)) + K[i] + w[i]) | 0;
        const t2 = ((r(a, 2) ^ r(a, 13) ^ r(a, 22)) + ((a & b) ^ (a & c) ^ (b & c))) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
      }
      H = H.map((v, i) => (v + [a, b, c, d, e, f, g, h][i]) | 0);
    }
    return H.map((v) => (v >>> 0).toString(16).padStart(8, "0")).join("");
  }
  async function hashClave(clave) {
    try {
      if (window.crypto && crypto.subtle) {
        const d = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(clave));
        return Array.from(new Uint8Array(d)).map((b) => b.toString(16).padStart(2, "0")).join("");
      }
    } catch { /* usa respaldo */ }
    return sha256Fallback(clave);
  }
  window.hashClave = hashClave;

  // ---------- Estado ----------
  let usuario = null;
  let localActual = null;
  let orden = { k: "local", dir: 1 };
  const locales = M.locales;
  const checklist = M.checklist.filter((c) => c.puntos !== "%");
  // Efectividad inventarios como en la matriz: 100% − Error/Ítems, y resta 1 punto por cada error
  const productoVacio = () => ({ tipo: "Inventario Total Producto", venta: "", fecha: "", items: "", error: "", sobrantes: [], faltantes: [], costos: {} });
  // Costos ideal / real como % de la venta. Diferencia = Ideal − Real: positiva es sobrante, negativa faltante.
  function difCosto(ideal, real) {
    const i = parseFloat(ideal), r = parseFloat(real);
    if (isNaN(i) || isNaN(r)) return null;
    const v = Math.round((i - r) * 100) / 100;
    return { ideal: i, real: r, v, tipo: v > 0 ? "sobrante" : v < 0 ? "faltante" : "sin diferencia" };
  }
  // Línea de costos para el pie del reporte: «Costo IDEAL 32,49 % · REAL 32,44 % · sobrante 0,05 % · merma 0,02 %»
  function textoCostos(co) {
    const d = difCosto(co.audIdeal, co.audReal) || difCosto(co.perIdeal, co.perReal);
    const partes = [];
    if (d) partes.push(`Costo IDEAL ${fmtP(d.ideal)}`, `REAL ${fmtP(d.real)}`, d.v ? `${d.tipo} ${fmtP(Math.abs(d.v))}` : "sin diferencia");
    if (co.merma !== "" && co.merma != null) partes.push(`merma ${fmtP(co.merma)}`);
    if (co.lunch !== "" && co.lunch != null) partes.push(`lunch ${fmtP(co.lunch)}`);
    return partes.join("  ·  ");
  }
  const fmtP = (v) => (v == null || v === "" || isNaN(v) ? "—" : (+v).toLocaleString("es-EC", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " %");
  function efectividadDe(p) {
    const items = parseFloat(p.items), error = parseFloat(p.error) || 0;
    if (!(items > 0)) return { ef: NaN, penal: 0, items: 0, error };
    return { ef: Math.round((1 - error / items) * 10000) / 100, penal: error, items, error };
  }

  const claveLocal = (l) => `matriz:${usuario.usuario}:${l}`;
  const estadoLocal = (l) => store.get(claveLocal(l), { marcados: [], efectividad: "", observaciones: "", informe: "", fecha: "", activos: [], sir: 0 });
  const guardar = (patch) => {
    if (!localActual) return;
    store.set(claveLocal(localActual.local), Object.assign(estadoLocal(localActual.local), patch));
  };

  // Aviso breve; con acción (p. ej. «Deshacer») se queda más tiempo y muestra un botón
  function toast(msg, accion, alHacer) {
    const t = $("#toast"); t.textContent = msg; t.hidden = false;
    if (accion) {
      const b = document.createElement("button"); b.type = "button"; b.className = "toast-btn"; b.textContent = accion;
      b.addEventListener("click", () => { t.hidden = true; alHacer(); });
      t.append(" ", b);
    }
    clearTimeout(toast._t); toast._t = setTimeout(() => (t.hidden = true), accion ? 6000 : 2200);
  }

  // ---------- Login ----------
  let intentos = 0;
  $("#login-version").textContent = `Grupo KFC · Auditoría Ecuador · ${M.version}`;
  $("#btn-ver-clave").addEventListener("click", (e) => {
    const i = $("#login-clave"); i.type = i.type === "password" ? "text" : "password";
    e.target.textContent = i.type === "password" ? "Ver" : "Ocultar";
  });
  $("#form-login").addEventListener("submit", async (e) => {
    e.preventDefault();
    const err = $("#login-error");
    if (intentos >= 5) { err.textContent = "Demasiados intentos. Espera 30 segundos."; err.hidden = false; return; }
    const u = $("#login-usuario").value.trim().toLowerCase();
    const h = await hashClave($("#login-clave").value);
    const encontrado = USUARIOS.find((x) => x.usuario.toLowerCase() === u && x.hash === h);
    if (!encontrado) {
      intentos++;
      if (intentos >= 5) setTimeout(() => (intentos = 0), 30000);
      err.textContent = "Usuario o contraseña incorrectos."; err.hidden = false;
      $("#login-clave").select();
      return;
    }
    err.hidden = true; intentos = 0;
    sess.set("matriz:usuario", encontrado.usuario);
    entrar(encontrado);
  });
  $("#btn-salir").addEventListener("click", () => {
    sess.del("matriz:usuario"); usuario = null;
    $("#login-clave").value = "";
    $("#vista-app").hidden = true; $("#vista-login").hidden = false;
  });

  function entrar(u) {
    usuario = u;
    $("#vista-login").hidden = true; $("#vista-app").hidden = false;
    $("#user-nombre").textContent = u.nombre;
    $("#app-version").textContent = `${M.version} · ${M.elaborado}`;
    $("#f-mios").checked = !!u.auditor;
    $("#f-mios").parentElement.hidden = !u.auditor;
    renderLocales();
    const ultimo = store.get(`matriz:${u.usuario}:ultimo`, null);
    const l = locales.find((x) => x.local === ultimo) || (u.auditor && locales.find((x) => x.auditor === u.auditor));
    if (l) seleccionarLocal(l.local, false);
    // Al cambiar de usuario se empieza siempre en «Locales»: ninguna pestaña conserva lo que se dibujó para el usuario anterior
    $("#mn-secciones").innerHTML = ""; $("#mn-admin").hidden = true; $("#mn-aviso").hidden = true;
    $("#cu-zona").disabled = false; delete $("#cu-zona").dataset.elegida;
    // Usuarios de una sola pestaña: Activos (Novedades de activos) y Conciliación (Conciliación bancaria). Los demás no ven «Novedades de activos».
    const unica = { activos: "novact", conciliacion: "conciliacion" }[u.rol];
    $$(".tabs button").forEach((b) => (b.hidden = unica ? b.dataset.tab !== unica : b.dataset.tab === "novact"));
    $$("[data-enviar-activos]").forEach((b) => (b.hidden = u.rol !== "auditor"));
    $("#cb-mapeo").hidden = true;
    irA(unica || "locales");
  }

  // ---------- Tabs ----------
  function irA(tab) {
    $$(".tabs button").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
    $$("section[data-panel]").forEach((s) => (s.hidden = s.dataset.panel !== tab));
    if (tab === "correo") prepararCorreo();
    if (tab === "tendencias") renderTendencias();
    if (tab === "producto") renderDashboard();
    if (tab === "bajas") renderBajas();
    if (tab === "cumplimiento") prepararCumplimiento();
    if (tab === "manuales") renderManuales();
    if (tab === "novact") renderNovedadesActivos();
    if (tab === "conciliacion") renderConciliacion();
    if (tab === "activos") renderBarraLibros();
    window.scrollTo(0, 0);
  }
  $$(".tabs button").forEach((b) => b.addEventListener("click", () => irA(b.dataset.tab)));
  $$("[data-ir]").forEach((b) => b.addEventListener("click", () => {
    if (b.dataset.enfoque) elegirCorreo(b.dataset.enfoque); // cada pestaña abre su propio tipo de correo
    irA(b.dataset.ir);
  }));

  // ---------- Locales ----------
  const unicos = (k) => [...new Set(locales.map((l) => l[k]).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
  const llenar = (sel, vals) => vals.forEach((v) => sel.insertAdjacentHTML("beforeend", `<option>${esc(v)}</option>`));
  llenar($("#f-zona"), unicos("zona"));
  llenar($("#f-cadena"), unicos("cadena"));
  llenar($("#f-auditor"), unicos("auditor"));
  llenar($("#f-banco"), unicos("banco"));
  $("#lista-locales").innerHTML = locales.map((l) => `<option value="${esc(l.local)}">${esc(l.cadena)} · ${esc(l.auditor)}</option>`).join("");

  const sinGrabacion = (l) => isNaN(parseFloat(l.grabacion)) || /DAÑADO/i.test(l.status);
  function filtrados() {
    const q = $("#f-buscar").value.trim().toLowerCase();
    const z = $("#f-zona").value, c = $("#f-cadena").value, a = $("#f-auditor").value, b = $("#f-banco").value;
    const mios = $("#f-mios").checked && usuario && usuario.auditor;
    return locales.filter((l) =>
      (!z || l.zona === z) && (!c || l.cadena === c) && (!a || l.auditor === a) && (!b || l.banco === b) &&
      (!mios || l.auditor === usuario.auditor) &&
      (!q || [l.local, l.jefe, l.auditor, l.cadena, l.razon].join(" ").toLowerCase().includes(q)));
  }
  function renderLocales() {
    const lista = filtrados().sort((x, y) => {
      const a = x[orden.k], b = y[orden.k], na = parseFloat(a), nb = parseFloat(b);
      const r = !isNaN(na) && !isNaN(nb) ? na - nb : String(a).localeCompare(String(b), "es");
      return r * orden.dir;
    });
    $$("#tabla-locales th").forEach((th) => th.className = (th.classList.contains("num") ? "num " : "") + (th.dataset.k === orden.k ? (orden.dir > 0 ? "asc" : "desc") : ""));
    $("#tabla-locales tbody").innerHTML = lista.map((l) => {
      const grab = sinGrabacion(l) ? `<span class="tag bad">${esc(l.grabacion || "—")}</span>` : esc(l.grabacion);
      const st = l.status ? `<span class="tag ${/DAÑADO/i.test(l.status) ? "bad" : "warn"}">${esc(l.status)}</span>` : "";
      return `<tr data-local="${esc(l.local)}"><td><b>${esc(l.local)}</b></td><td>${esc(l.cadena)}</td><td>${esc(l.zona)}</td>
        <td>${esc(l.auditor)}</td><td>${esc(l.jefe)}</td><td>${esc(l.banco)}</td><td>${esc(l.deposito)}</td>
        <td class="num">$${esc(l.caja)}</td><td>${esc(l.dvr)}</td><td class="num">${grab}</td><td>${st}</td></tr>`;
    }).join("");
    const caja = lista.reduce((s, l) => s + (parseFloat(l.caja) || 0), 0);
    const kpis = [
      [lista.length, "Locales en vista"],
      [new Set(lista.map((l) => l.cadena)).size, "Cadenas"],
      [new Set(lista.map((l) => l.auditor)).size, "Auditores"],
      ["$" + caja.toLocaleString("es-EC"), "Caja chica total"],
      [lista.filter(sinGrabacion).length, "DVR sin grabación / dañado"],
    ];
    $("#kpis").innerHTML = kpis.map(([v, t]) => `<div><b>${esc(v)}</b><span>${esc(t)}</span></div>`).join("");
    $("#conteo").textContent = `${lista.length} de ${locales.length} locales · clic en una fila para abrir su informe`;
  }
  ["#f-buscar", "#f-zona", "#f-cadena", "#f-auditor", "#f-banco", "#f-mios"].forEach((s) => $(s).addEventListener("input", renderLocales));
  $$("#tabla-locales th").forEach((th) => th.addEventListener("click", () => {
    orden = { k: th.dataset.k, dir: orden.k === th.dataset.k ? -orden.dir : 1 }; renderLocales();
  }));
  $("#tabla-locales tbody").addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-local]"); if (!tr) return;
    seleccionarLocal(tr.dataset.local); irA("informe");
  });

  // ---------- Informe / Checklist ----------
  // Ítems 1–16 (procesos sin descripción en la matriz) se evalúan Cumple / No cumple con observación;
  // los demás se marcan como novedad.
  const TXT_CUMPLE = "Se verifica proceso y existe cumplimiento";
  const esProceso = (c) => !c.desc;
  const grupos = checklist.reduce((g, c) => {
    const area = c.desc ? c.area : "Procesos de caja y administración";
    (g[area] = g[area] || []).push(c); return g;
  }, {});
  $("#checklist").innerHTML = Object.entries(grupos).map(([area, items]) => `
    <div class="grupo"><h3>${esc(area)}${items.some(esProceso) ? ` <button type="button" class="ghost sm" id="btn-todos-cumplen">Todos cumplen</button>` : ""}</h3>${items.map((c) => esProceso(c) ? `
      <div class="proc" data-n="${c.n}">
        <div class="proc-head"><span class="proc-n">${c.n}</span><b>${esc(c.area)}</b><span class="pts">-${esc(c.puntos)}</span></div>
        <div class="seg" role="radiogroup" aria-label="${esc(c.area)}">
          <label><input type="radio" name="ev-${c.n}" value="cumple"><span>Cumple</span></label>
          <label><input type="radio" name="ev-${c.n}" value="no"><span>No cumple</span></label>
        </div>
        <p class="proc-ok" hidden>✓ ${TXT_CUMPLE}</p>
        <textarea class="proc-obs" data-n="${c.n}" rows="2" placeholder="Observación de la novedad (obligatoria)" hidden></textarea>
      </div>` : `
      <label class="item" data-n="${c.n}"><input type="checkbox" value="${c.n}">
        <span>${esc(c.desc ? c.desc : c.area)}${c.desc && c.area !== area ? `<small>${esc(c.area)}</small>` : ""}</span>
        <span class="pts">-${esc(c.puntos)}</span></label>`).join("")}</div>`).join("");

  function leerEvaluacion() {
    const ev = {};
    $$("#checklist .proc").forEach((el) => {
      const sel = el.querySelector("input[type=radio]:checked");
      const obs = el.querySelector(".proc-obs").value;
      if (sel || obs) ev[el.dataset.n] = { estado: sel ? sel.value : "", obs };
    });
    return ev;
  }
  // Puntos por novedad de activo (fórmula de la matriz) y el ítem del check list que cada una alimenta.
  // En la matriz: K = cantidad de activos con esa novedad, M = −K × puntos.
  const PUNTOS_NOV = { "activo perdido": 10, "se mueve sin transferencia": 5, "equipo no creado": 5, "no realizan inv. mensual": 5 };
  const puntosNovedad = (n) => PUNTOS_NOV[String(n || "").toLowerCase()] || 0;
  const ITEM_ACTIVO = Object.fromEntries(checklist.filter((c) => c.area === "Activos").map((c) => [c.n,
    /PERDIDO/i.test(c.desc) ? "Activo perdido" : /INVENTARIO TOTAL/i.test(c.desc) ? "No realizan INV. Mensual"
      : /Nuevos/i.test(c.desc) ? "Equipo No creado" : /Transferencia/i.test(c.desc) ? "Se mueve sin transferencia" : null]).filter(([, v]) => v));
  // Si el local tiene listado de activos, los ítems de Activos del check list se calculan solos desde ese listado
  function novedadesActivos() {
    const st = localActual ? estadoLocal(localActual.local) : { activos: [] };
    const lista = (st.activos || []).filter((a) => a.estado !== "Transferido");
    const porTipo = {};
    lista.forEach((a) => { if (puntosNovedad(a.novedad)) (porTipo[a.novedad.toLowerCase()] = porTipo[a.novedad.toLowerCase()] || []).push(a); });
    return { auto: lista.length > 0, porTipo };
  }
  function calcular() {
    const na = novedadesActivos();
    const esAuto = (c) => na.auto && ITEM_ACTIVO[c.n];
    const marcados = $$("#checklist .item input:checked").map((i) => +i.value).filter((n) => !(na.auto && ITEM_ACTIVO[n]));
    const evaluacion = leerEvaluacion();
    const estado = (c) => (evaluacion[c.n] || {}).estado;
    const items = checklist.filter((c) => (esProceso(c) ? estado(c) === "no" : esAuto(c) ? (na.porTipo[ITEM_ACTIVO[c.n].toLowerCase()] || []).length : marcados.includes(c.n)))
      .map((c) => {
        const activos = esAuto(c) ? na.porTipo[ITEM_ACTIVO[c.n].toLowerCase()] : null;
        return Object.assign({}, c, { obs: esProceso(c) ? evaluacion[c.n].obs.trim() : "", activos, cant: activos ? activos.length : 1,
          puntosTot: (parseFloat(c.puntos) || 0) * (activos ? activos.length : 1) });
      });
    const cumplen = checklist.filter((c) => esProceso(c) && estado(c) === "cumple");
    const pendientes = checklist.filter((c) => esProceso(c) && !estado(c));
    const sinObs = items.filter((c) => esProceso(c) && !c.obs);
    let resta = items.reduce((s, c) => s + c.puntosTot, 0);
    const efx = efectividadDe(localActual ? estadoLocal(localActual.local).producto || {} : {});
    const ef = efx.ef, penalEf = efx.penal, efItems = efx.items, efError = efx.error;
    // La efectividad de inventarios se muestra en Producto; ya no descuenta en el check list
    const score = Math.max(0, 100 - resta);
    return { marcados, evaluacion, items, cumplen, pendientes, sinObs, score, ef, penalEf, efItems, efError, autoActivos: na.auto };
  }
  function renderScore() {
    const r = calcular();
    $$("#checklist .item").forEach((el) => {
      const n = +el.dataset.n, it = r.items.find((c) => c.n === n), auto = r.autoActivos && ITEM_ACTIVO[n];
      const input = el.querySelector("input"), base = checklist.find((c) => c.n === n);
      if (auto) input.checked = !!it;
      input.disabled = !!auto;
      el.classList.toggle("on", !!it);
      el.querySelector(".pts").textContent = `-${it ? it.puntosTot : base.puntos}`;
      let nota = el.querySelector(".auto"); if (!nota) { nota = document.createElement("small"); nota.className = "auto"; el.children[1].appendChild(nota); }
      nota.textContent = auto ? (it ? `Desde Activos: ${it.cant} activo(s) × ${base.puntos} pts — ${it.activos.map((a) => a.nombre).join(", ")}` : "Desde Activos: sin activos con esta novedad") : "";
    });
    $$("#checklist .proc").forEach((el) => {
      const e = r.evaluacion[el.dataset.n] || {};
      el.classList.toggle("ok", e.estado === "cumple");
      el.classList.toggle("on", e.estado === "no");
      el.querySelector(".proc-ok").hidden = e.estado !== "cumple";
      const ta = el.querySelector(".proc-obs");
      ta.hidden = e.estado !== "no";
      ta.classList.toggle("falta", e.estado === "no" && !ta.value.trim());
    });
    $("#res-cuadro").innerHTML = htmlResultado(false);
    const partes = [];
    partes.push(r.items.length ? `${r.items.length} novedad(es)` : "Sin novedades");
    partes.push(`${r.cumplen.length} procesos cumplen`);
    if (r.pendientes.length) partes.push(`${r.pendientes.length} sin evaluar`);
    if (r.sinObs.length) partes.push(`${r.sinObs.length} sin observación`);
    $("#score-txt").textContent = partes.join(" · ");
    return r;
  }
  // Cuadro de resultado como en la matriz: Informe (chequeo total / visita parcial), local, fecha, calificación
  // (solo chequeo total), puntos perdidos y una fila por novedad con su nota. correo = true: estilos en línea para Outlook.
  function htmlResultado(correo) {
    if (!localActual) return `<p class="muted small">Selecciona un local.</p>`;
    const r = calcular(), parcial = /visita/i.test($("#sel-informe").value), perdidos = 100 - r.score;
    const s = (css) => (correo ? ` style="${css}"` : "");
    const GRIS = "#646d78", BORDE = "#dfe3e8", ENC = "#ffc629", ROJO = "#b42318";
    const iso = $("#sel-fecha").value || hoy(), d = new Date(iso + "T12:00:00");
    const fecha = isNaN(d) ? iso : d.toLocaleDateString("es-EC", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
    const filas = r.items.map((c) => {
      const nota = c.activos ? c.activos.map((a) => a.nombre).join(", ") : esProceso(c) ? "" : c.obs;
      return { area: esProceso(c) ? c.area : c.area, plan: esProceso(c) ? c.obs || c.area : c.desc, nota, cant: c.cant, valor: c.puntos, pts: c.puntosTot };
    });
    const tdc = (v, css = "", cls = "") => `<td${cls ? ` class="${cls}"` : ""}${s(`padding:5px 8px;border-bottom:1px dotted ${BORDE};font-size:10pt;vertical-align:top;${css}`)}>${v}</td>`;
    const informe = correo ? `<b${s("font-size:15pt")}>${esc($("#sel-informe").value)}</b>`
      : `<select id="res-informe" aria-label="Tipo de check list">${["Chequeo Total - Check List Procesos", "Visita Parcial - Check List Procesos"].map((o) => `<option${o === $("#sel-informe").value ? " selected" : ""}>${o}</option>`).join("")}</select>`;
    return `<table class="rc"${s("width:100%;border-collapse:collapse;font-family:'Segoe UI',Arial,sans-serif;color:#1b1f24")}>
      <tr><td${s("vertical-align:top")}>
        <table class="rc-cab"${s("border-collapse:collapse")}>
          <tr><td${s(`color:${GRIS};font-size:9pt;font-weight:bold;text-align:right;padding:2px 8px`)}>Informe:</td><td${s("padding:2px 8px")}>${informe}</td></tr>
          <tr><td${s(`color:${GRIS};font-size:9pt;font-weight:bold;text-align:right;padding:2px 8px`)}>Local:</td><td${s(`color:${GRIS};font-size:10pt;padding:2px 8px`)}>${esc(localActual.local)}</td></tr>
          <tr><td${s(`color:${GRIS};font-size:9pt;font-weight:bold;text-align:right;padding:2px 8px`)}>Fecha:</td><td${s(`color:${GRIS};font-size:10pt;padding:2px 8px`)}>${esc(fecha)}</td></tr>
        </table></td>
        <td class="rc-nota"${s("vertical-align:top;text-align:right;width:160px")}>
          ${parcial ? "" : `<div class="rc-score"${s(`background:${ENC};font-size:22pt;font-weight:bold;padding:6px 12px;text-align:right`)}>${r.score} <small${s("font-size:14pt")}>/100</small></div>`}
          <div class="rc-lbl"${s(`color:${GRIS};font-size:8.5pt;font-weight:bold;margin-top:4px`)}>Puntos perdidos</div>
        </td></tr>
      <tr><td colspan="2"${s("padding-top:4px")}>
        <table class="rc-tabla"${s("width:100%;border-collapse:collapse")}>
          <tr><th${s(`background:${ENC};padding:5px 8px;font-size:9pt;text-align:center;width:150px`)}>Área</th>
            <th${s("background:#e9ebef;padding:5px 8px;font-size:9pt;text-align:center")}>Plan de Trabajo</th>
            <th${s("background:#e9ebef;padding:5px 8px;font-size:9pt;width:50px")}>Cant.</th>
            <th${s("background:#e9ebef;padding:5px 8px;font-size:9pt;width:50px")}>Valor</th>
            <th class="rc-total"${s(`background:${ENC};padding:5px 8px;font-size:15pt;text-align:right;width:70px`)}>${perdidos ? `-${perdidos}` : "0"}</th></tr>
          ${filas.length ? filas.map((f) => `<tr>${tdc(esc(f.area), "font-weight:bold;text-align:center;background:#fff6d6", "rc-area")}
            ${tdc(esc(f.plan))}
            ${tdc(f.cant, `text-align:center;color:${GRIS}`)}${tdc(esc(f.valor), `text-align:center;color:${GRIS}`)}
            ${tdc(`-${f.pts}`, `text-align:right;font-weight:bold;font-size:12pt;color:${ROJO}`, "rc-pts")}</tr>`).join("")
            : `<tr>${tdc("Sin novedades en los procesos revisados.", `color:${GRIS}`)}</tr>`}
        </table></td></tr></table>`;
  }
  // El selector del cuadro cambia el tipo de check list (el mismo de la barra superior)
  $("#res-cuadro").addEventListener("change", (e) => {
    if (e.target.id !== "res-informe") return;
    $("#sel-informe").value = e.target.value; $("#sel-informe").dispatchEvent(new Event("change")); renderScore();
  });
  $("#sel-informe").addEventListener("change", () => renderScore());
  // El cuadro dibujado como imagen (PNG nítido) para pegar en el informe o el correo tal cual se ve
  function dibujarResultado() {
    const r = calcular(), parcial = /visita/i.test($("#sel-informe").value), perdidos = 100 - r.score;
    const W = 1000, pad = 20, S = ESCALA_COPIA, FUENTE_R = '"Segoe UI", Arial, sans-serif';
    const cols = [["Área", 200, "l"], ["Plan de Trabajo", 540, "l"], ["Cant.", 70, "c"], ["Valor", 70, "c"], [perdidos ? `-${perdidos}` : "0", 80, "r"]];
    const med = document.createElement("canvas").getContext("2d");
    const partir = (t, ancho, px, peso = 400) => { // corta el texto en renglones que quepan en la columna
      med.font = `${peso} ${px}px ${FUENTE_R}`;
      const lineas = []; let actual = "";
      String(t).split(/\s+/).forEach((p) => { const prueba = actual ? `${actual} ${p}` : p; if (med.measureText(prueba).width > ancho && actual) { lineas.push(actual); actual = p; } else actual = prueba; });
      if (actual) lineas.push(actual); return lineas.length ? lineas : [""];
    };
    const filas = r.items.map((c) => {
      const area = esProceso(c) ? c.area : c.area, plan = esProceso(c) ? c.obs || c.area : c.desc;
      const la = partir(area, cols[0][1] - 20, 15, 700), lp = partir(plan, cols[1][1] - 20, 15);
      return { la, lp, cant: c.cant, valor: c.puntos, pts: c.puntosTot, alto: Math.max(la.length, lp.length) * 21 + 16 };
    });
    const altoTabla = 38 + (filas.length ? filas.reduce((s, f) => s + f.alto, 0) : 40);
    const H = pad + 100 + altoTabla + pad;
    const cv = document.createElement("canvas"); cv.width = W * S; cv.height = H * S;
    const g = cv.getContext("2d"); g.scale(S, S); g.imageSmoothingQuality = "high";
    const txt = (t, x, y, { px = 15, peso = 400, color = "#1b1f24", al = "left" } = {}) => { g.font = `${peso} ${px}px ${FUENTE_R}`; g.fillStyle = color; g.textAlign = al; g.fillText(t, x, y); };
    g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
    // Encabezado
    [["Informe:", $("#sel-informe").value, 20, 700], ["Local:", localActual.local, 15, 400], ["Fecha:", (() => { const iso = $("#sel-fecha").value || hoy(), d = new Date(iso + "T12:00:00");
      return isNaN(d) ? iso : d.toLocaleDateString("es-EC", { weekday: "long", day: "numeric", month: "long", year: "numeric" }); })(), 15, 400]]
      .forEach(([k, v, px, peso], i) => { const y = pad + 26 + i * 30; txt(k, pad + 70, y, { px: 13, peso: 700, color: "#646d78", al: "right" }); txt(v, pad + 84, y, { px, peso }); });
    if (!parcial) {
      const gr = g.createLinearGradient(W - pad - 170, 0, W - pad, 0); gr.addColorStop(0, "#ffe08a"); gr.addColorStop(1, "#ffc629");
      g.fillStyle = gr; g.fillRect(W - pad - 170, pad + 4, 170, 52);
      txt(`${r.score}`, W - pad - 64, pad + 43, { px: 32, peso: 800, al: "right" }); txt("/100", W - pad - 12, pad + 43, { px: 20, peso: 700, al: "right" });
    }
    txt("Puntos perdidos", W - pad, pad + 84, { px: 12, peso: 700, color: "#646d78", al: "right" });
    // Tabla
    let y = pad + 100, x = pad;
    cols.forEach(([n, w, al], i) => {
      g.fillStyle = i === 0 || i === 4 ? "#ffc629" : "#e9ebef"; g.fillRect(x, y, w, 38);
      txt(n, al === "r" ? x + w - 10 : al === "c" ? x + w / 2 : x + 10, y + 25, { px: i === 4 ? 20 : 14, peso: i === 4 ? 800 : 700, al: al === "r" ? "right" : al === "c" ? "center" : "left" });
      x += w;
    });
    y += 38;
    if (!filas.length) { txt("Sin novedades en los procesos revisados.", pad + 10, y + 26, { color: "#646d78" }); }
    filas.forEach((f) => {
      x = pad;
      g.fillStyle = "#fff6d6"; g.fillRect(x, y, cols[0][1], f.alto);
      f.la.forEach((l, i) => txt(l, x + 10, y + 24 + i * 21, { peso: 700 }));
      x += cols[0][1];
      f.lp.forEach((l, i) => txt(l, x + 10, y + 24 + i * 21));
      x += cols[1][1];
      txt(String(f.cant), x + cols[2][1] / 2, y + 24, { color: "#646d78", al: "center" }); x += cols[2][1];
      txt(String(f.valor), x + cols[3][1] / 2, y + 24, { color: "#646d78", al: "center" }); x += cols[3][1];
      txt(`-${f.pts}`, x + cols[4][1] - 10, y + 25, { px: 17, peso: 800, color: "#b42318", al: "right" });
      y += f.alto; g.fillStyle = "#dfe3e8"; g.fillRect(pad, y - 1, W - 2 * pad, 1);
    });
    return cv;
  }
  $("#res-copiar").addEventListener("click", async () => {
    if (!localActual) return toast("Selecciona un local primero");
    const blob = await new Promise((ok) => dibujarResultado().toBlob(ok, "image/png"));
    try { await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]); toast("Cuadro copiado como imagen: pégalo en el informe o el correo con Ctrl+V"); }
    catch { descargar(blob, `Resultado check list ${localActual.local}.png`.replace(/[\\/:*?"<>|]/g, "_")); toast("Tu navegador no permitió copiar: se descargó la imagen"); }
  });
  const guardarChecklist = () => { const r = renderScore(); guardar({ marcados: r.marcados, evaluacion: r.evaluacion }); };
  $("#checklist").addEventListener("change", guardarChecklist);
  $("#checklist").addEventListener("input", (e) => { if (e.target.classList.contains("proc-obs")) guardarChecklist(); });
  $("#checklist").addEventListener("click", (e) => {
    if (e.target.id !== "btn-todos-cumplen") return;
    $$("#checklist .proc").forEach((el) => { if (!el.querySelector("input:checked")) el.querySelector("input[value=cumple]").checked = true; });
    guardarChecklist();
  });
  $("#observaciones").addEventListener("input", () => guardar({ observaciones: $("#observaciones").value }));
  $("#sel-informe").addEventListener("change", () => guardar({ informe: $("#sel-informe").value }));
  $("#sel-fecha").addEventListener("change", () => guardar({ fecha: $("#sel-fecha").value }));
  $("#btn-limpiar").addEventListener("click", () => {
    if (!localActual || !confirm("¿Limpiar el checklist de este local?")) return;
    guardar({ marcados: [], evaluacion: {}, efectividad: "", observaciones: "" }); seleccionarLocal(localActual.local, false);
  });
  $("#sel-local").addEventListener("change", () => {
    const l = locales.find((x) => x.local.toLowerCase() === $("#sel-local").value.trim().toLowerCase());
    if (l) seleccionarLocal(l.local, false); else toast("Local no encontrado");
  });

  const hoy = () => new Date().toISOString().slice(0, 10);
  function seleccionarLocal(nombre, avisar = true) {
    localActual = locales.find((l) => l.local === nombre); if (!localActual) return;
    store.set(`matriz:${usuario.usuario}:ultimo`, nombre);
    const st = estadoLocal(nombre), l = localActual;
    $("#sel-local").value = l.local;
    $("#sel-informe").value = st.informe || "Chequeo Total - Check List Procesos";
    $("#sel-fecha").value = st.fecha || hoy();
    $("#observaciones").value = st.observaciones;
    $$("#checklist .item input").forEach((i) => (i.checked = st.marcados.includes(+i.value)));
    const ev = st.evaluacion || {};
    $$("#checklist .proc").forEach((el) => {
      const e = ev[el.dataset.n] || {};
      el.querySelectorAll("input[type=radio]").forEach((i) => (i.checked = i.value === e.estado));
      el.querySelector(".proc-obs").value = e.obs || "";
    });
    const campos = [["Cadena", l.cadena], ["Zona", l.zona], ["Razón social", l.razon], ["Jefe de área", l.jefe], ["Auditor", l.auditor],
      ["Banco", `${l.banco} · ${l.deposito}`], ["Caja chica", "$" + l.caja], ["DVR", `${l.dvr} · ${l.grabacion} días`]];
    if (l.status) campos.push(["Status", l.status]);
    $("#ficha-local").innerHTML = campos.map(([k, v]) => `<div><span>${esc(k)}</span>${esc(v)}</div>`).join("");
    renderProducto();
    renderScore();
    renderActivos();
    $("#m-para").value = ""; delete $("#m-cuerpo").dataset.editado;
    if (avisar) toast(`Local ${l.local}`);
  }

  // ---------- Activos ----------
  // Cada activo: { id, codigo, erp, nombre, marca, modelo, serie, fecha, estado, origen, novedad, fotos: [ids] }
  // novedad vacía = pendiente de revisar. Los «Transferido» del SIR no cuentan para el inventario del local.
  const OK_ACT = "Procesos y Controles OK";
  const enUso = (a) => a.estado !== "Transferido";
  const conNovedadActivo = (a) => enUso(a) && !!a.novedad && a.novedad !== OK_ACT;
  const nuevoId = () => "a" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  function activosLocal() {
    const st = estadoLocal(localActual.local);
    let cambio = false;
    st.activos.forEach((a) => { if (!a.id) { a.id = nuevoId(); a.origen = a.origen || "manual"; a.estado = a.estado || "Manual"; cambio = true; } a.fotos = a.fotos || []; });
    if (cambio) guardar({ activos: st.activos });
    return st;
  }
  let ultimaVistaActivos = [];
  // Línea con lo que trae la app de inventario: estado físico, ubicación, novedad y valor en libros
  function infoApp(a) {
    if (!a.app) return "";
    const x = a.app, partes = [x.estado, x.ubicacion, x.detalle, x.libros != null ? `libros ${fmtUSD(x.libros)}` : ""].filter(Boolean);
    const titulo = [x.notas, ...(x.hist || [])].filter(Boolean).join("\n");
    return `<small class="app-info${x.estado === "Baja" ? " baja" : ""}"${titulo ? ` title="${esc(titulo)}"` : ""}>📱 App: ${esc(partes.join(" · ") || "registrado")}</small>`;
  }
  function renderActivos() {
    // valor en libros: la base se carga una vez (de este equipo o de la carpeta publicada) y se vuelve a dibujar
    if (!libros && !librosIntentado) { librosIntentado = true; cargarLibros().then((b) => (b ? renderActivos() : renderBarraLibros())); }
    renderBarraLibros();
    if (!localActual) {
      $("#activos-local").textContent = "Selecciona un local en la pestaña Informe o en la tabla.";
      $("#tabla-activos tbody").innerHTML = `<tr><td colspan="9" class="muted">Selecciona un local.</td></tr>`; return;
    }
    const st = activosLocal(), lista = st.activos, base = lista.filter(enUso);
    const bMap = new Map((st.bajas || []).map((b) => [b.activoId, b]));
    if (gActivo && !lista.some((x) => x.id === gActivo.id)) { gActivo = null; renderGestion(); } // se cambió de local
    $("#activos-local").textContent = `${localActual.local} · Jefe de área: ${localActual.jefe}`;
    $("#a-sir").value = st.sir || 0;
    $("#a-registrados").textContent = st.sir || 0;
    $("#a-verificados").textContent = base.filter((a) => a.novedad).length;
    $("#a-pendientes").textContent = base.filter((a) => !a.novedad).length;
    $("#a-novedades").textContent = base.filter(conNovedadActivo).length;
    const ar = st.archivoActivos;
    $("#a-archivo-info").textContent = ar
      ? `Cargado «${ar.nombre}» · ${ar.local} · ${ar.fecha} · ${ar.activos} activos y ${ar.transferidos} transferidos (se quitaron ${ar.duplicados} filas repetidas).`
      : "Sube el archivo .xls de activos fijos exportado del SIR (o arrástralo aquí) para ver todos los activos del local.";
    const rj = st.respaldoApp, enApp = lista.filter((a) => a.app && a.origen !== "app").length;
    $("#j-info").textContent = rj
      ? `Importado «${rj.nombre}» el ${rj.fecha}: ${rj.items} activos de la app, ${enApp} enlazados con el listado y ${lista.filter((a) => a.origen === "app").length} solo en la app. Usa el filtro para ver los que no están en la app.`
      : "Importa el respaldo .json (con fotos, estado, ubicación y novedades) y se cruza con el listado del SIR por código y serie.";
    const nSueltas = (st.fotosSueltas || []).length;
    $("#j-sueltas").hidden = !nSueltas;
    $("#j-sueltas").textContent = `📷 Asignar fotos sueltas (${nSueltas})`;
    const nTransf = lista.length - base.length;
    $("#a-n-transf").textContent = nTransf ? `(${nTransf})` : "";
    const q = $("#a-buscar").value.trim().toLowerCase(), f = $("#a-filtro").value, verT = $("#a-transferidos").checked;
    const vista = lista.filter((a) => (verT || enUso(a))
      && (!q || [a.codigo, a.erp, a.nombre, a.marca, a.modelo, a.serie].join(" ").toLowerCase().includes(q))
      && (!f || ({ pend: () => enUso(a) && !a.novedad, nov: () => conNovedadActivo(a), ok: () => a.novedad === OK_ACT, fotos: () => a.fotos.length,
        enapp: () => !!a.app, noapp: () => !a.app && a.origen === "SIR" && enUso(a), soloapp: () => a.origen === "app" })[f]()))
      .sort((x, y) => (enUso(y) - enUso(x)) || x.nombre.localeCompare(y.nombre, "es"));
    ultimaVistaActivos = vista;
    const opciones = (sel) => `<option value="">— Sin revisar —</option>` + M.novedadesActivo.map((n) => `<option${n === sel ? " selected" : ""}>${esc(n)}</option>`).join("");
    $("#tabla-activos tbody").innerHTML = vista.length ? vista.map((a) => {
      const clase = (!enUso(a) ? "transf" : !a.novedad ? "pend" : a.novedad === OK_ACT ? "okk" : "nov") + (a.creadoAuditor ? " nuevo" : "");
      return `<tr data-id="${a.id}" class="${clase}">
        <td class="col-quitar"><button type="button" class="quitar" data-quitar-activo title="Eliminar activo" aria-label="Eliminar ${esc(a.nombre)}">✕</button></td>
        <td>${esc(a.codigo)}${a.erp ? `<br><small class="muted">ERP ${esc(a.erp)}</small>` : ""}</td>
        <td><b>${valorMod(a, "nombre")}</b>${a.fecha ? `<br><small class="muted">Creado ${esc(a.fecha)}</small>` : ""}${a.creadoAuditor ? `<br><small>Nuevo · creado por auditoría</small>` : ""}${bMap.has(a.id) ? `<br><span class="tag-baja">${bajaPendiente(bMap.get(a.id)) ? "🗑️ Baja pendiente" : "🗑️ Dado de baja"}</span>` : ""}${infoApp(a)}</td>
        <td>${valorMod(a, "marca")}${a.modelo || (a.cambios && a.cambios.modelo) ? `<br><small>${valorMod(a, "modelo")}</small>` : ""}</td><td class="serie">${valorMod(a, "serie")}</td>
        <td class="num vl-celda">${celdaLibros(a)}</td>
        <td><span class="tag ${a.estado === "Activo" ? "ok" : a.estado === "Transferido" ? "" : "warn"}">${esc(a.estado || "")}</span></td>
        <td><select data-novedad aria-label="Novedad de ${esc(a.nombre)}">${opciones(a.novedad)}</select></td>
        <td><button type="button" class="ghost btn-fotos${a.fotos.length ? " con" : ""}" data-fotos>📷 ${a.fotos.length || "+"}</button><button type="button" class="ghost btn-editar" data-editar title="Modificar o reportar perdido" aria-label="Modificar ${esc(a.nombre)}">✏️</button></td></tr>`;
    }).join("") : `<tr><td colspan="9" class="muted">${lista.length ? "Ningún activo coincide con el filtro." : "Sin activos todavía: sube el archivo del SIR o agrega uno abajo."}</td></tr>`;
    $("#cuadro-act").innerHTML = htmlCuadroActivos(false);
    renderScore(); // los ítems de Activos del check list dependen de estas novedades
  }

  // Cuadro «Inventario Total de Activos» como en la matriz: equipos con novedad, modificados o nuevos,
  // y a la derecha cada tipo de novedad con sus puntos perdidos (−10 perdido, −5 sin transferencia / no creado / sin inv. mensual).
  function datosCuadroActivos() {
    const st = estadoLocal(localActual.local), lista = (st.activos || []).filter((a) => a.estado !== "Transferido");
    const filas = lista.filter((a) => (a.novedad && a.novedad !== OK_ACT) || (a.cambios && Object.keys(a.cambios).length) || a.creadoAuditor);
    const tipos = M.novedadesActivo.map((n) => { const xs = lista.filter((a) => a.novedad === n); return { n, cant: xs.length, pts: -puntosNovedad(n) * xs.length, base: puntosNovedad(n) }; });
    return { st, lista, filas, tipos, total: tipos.reduce((s, t) => s + t.pts, 0), verificados: lista.filter((a) => a.novedad).length };
  }
  // correo = true: tablas con estilos en línea para Outlook; false: para la página
  function htmlCuadroActivos(correo) {
    if (!localActual) return "";
    const d = datosCuadroActivos();
    if (!d.lista.length) return correo ? "" : `<p class="muted small">El cuadro aparece cuando subes el listado de activos del SIR.</p>`;
    const fecha = new Date().toLocaleDateString("es-EC", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
    const AM = "#fff3b0", AMB = "#3d3000", ROJO = "#b42318", VERDE = "#1f7a4d", GRIS = "#646d78", BORDE = "#dfe3e8", ENC = "#ffc629";
    const s = (css) => (correo ? ` style="${css}"` : "");
    const val = (a, k) => { const c = a.cambios && a.cambios[k], v = esc(a[k] || ""); return (c || a.creadoAuditor) && v ? `<span${s(`background:${AM};color:${AMB};padding:0 3px`)} class="mod"${c ? ` title="Antes: ${esc(c.antes || "(vacío)")}"` : ""}>${v}</span>` : v; };
    const td = (v, extra = "") => `<td${s(`padding:5px 8px;border-bottom:1px solid ${BORDE};font-size:10pt;${extra}`)}>${v}</td>`;
    const th = (v, extra = "") => `<th${s(`padding:6px 8px;background:${ENC};color:#1b1f24;font-size:9pt;text-align:left;${extra}`)}>${v}</th>`;
    const pendientes = d.lista.length - d.verificados;
    const titulo = d.total < 0 ? `Puntos perdidos: ${d.total}` : !d.verificados ? "Pendiente de verificación"
      : pendientes ? `Sin puntos perdidos · ${pendientes} pendientes` : "Felicitaciones";
    return `<table class="ca"${s(`width:100%;border-collapse:collapse;font-family:'Segoe UI',Arial,sans-serif;color:#1b1f24`)}><tr>
      <td class="ca-izq"${s("vertical-align:top;padding-right:16px")}>
        <table class="ca-cab"${s("border-collapse:collapse;margin-bottom:8px")}>
          <tr><td${s(`color:${GRIS};font-size:9pt;font-weight:bold;text-align:right;padding:2px 8px`)}>Informe:</td><td${s("font-size:15pt;font-weight:bold;padding:2px 8px")} class="ca-informe">Inventario Total de Activos</td></tr>
          <tr><td${s(`color:${GRIS};font-size:9pt;font-weight:bold;text-align:right;padding:2px 8px`)}>Local:</td><td${s(`color:${GRIS};font-size:10pt;padding:2px 8px`)}>${esc(localActual.local)}</td></tr>
          <tr><td${s(`color:${GRIS};font-size:9pt;font-weight:bold;text-align:right;padding:2px 8px`)}>Fecha:</td><td${s(`color:${GRIS};font-size:10pt;padding:2px 8px`)}>${esc(fecha)}</td></tr>
        </table>
        <div class="ca-titulo ${d.total < 0 ? "mal" : titulo === "Felicitaciones" ? "bien" : "pend"}"${s(`font-size:20pt;font-weight:bold;margin:4px 0 6px;color:${d.total < 0 ? ROJO : titulo === "Felicitaciones" ? VERDE : GRIS}`)}>${esc(titulo)}</div>
        <div class="ca-kpis"${s("margin:0 0 10px")}>
          <span${s(`display:inline-block;background:${ENC};font-weight:bold;font-size:10pt;padding:4px 10px;margin-right:6px;border-radius:4px`)}>Activos Registrados en el SIR: ${d.st.sir || "-"}</span>
          <span${s(`display:inline-block;background:${ENC};font-weight:bold;font-size:10pt;padding:4px 10px;border-radius:4px`)}>Activos Verificados Auditor: ${d.verificados || "-"}</span>
        </div>
        <table class="ca-tabla"${s("width:100%;border-collapse:collapse")}>
          <tr>${th("Código")}${th("Nombre")}${th("Marca")}${th("Modelo")}${th("Serie")}${th("Novedad")}${th("Pts", "text-align:right")}</tr>
          ${d.filas.length ? d.filas.map((a) => `<tr>${td(esc(a.codigo))}${td(`<b>${val(a, "nombre")}</b>`)}${td(val(a, "marca"))}${td(val(a, "modelo"))}${td(val(a, "serie"))}
            ${td(esc(a.novedad || "—"), `color:${puntosNovedad(a.novedad) ? ROJO : GRIS}`)}${td(puntosNovedad(a.novedad) ? `-${puntosNovedad(a.novedad)}` : "0", "text-align:right;font-weight:bold")}</tr>`).join("")
            : `<tr>${td(`<span${s(`color:${VERDE}`)}>Todos los activos verificados están sin novedad.</span>`, "")}</tr>`}
        </table>
        <p class="muted small"${s(`font-size:8.5pt;color:${GRIS};margin:6px 0 0`)}>Los puntos se descuentan en la evaluación del check list (pestaña Informe, ítems de Activos).</p>
      </td></tr></table>`;
  }
  $("#ca-copiar").addEventListener("click", async () => {
    const html = htmlCuadroActivos(true); if (!html) return toast("Primero sube el listado de activos");
    try { await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([$("#cuadro-act").innerText], { type: "text/plain" }) })]); }
    catch { const div = document.createElement("div"); div.innerHTML = html; div.style.cssText = "position:fixed;left:-9999px;background:#fff"; document.body.appendChild(div);
      const rg = document.createRange(); rg.selectNodeContents(div); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(rg); document.execCommand("copy"); sel.removeAllRanges(); div.remove(); }
    toast("Cuadro copiado: pégalo en el correo con Ctrl+V");
  });
  ["#a-buscar", "#a-filtro", "#a-transferidos"].forEach((s) => $(s).addEventListener("input", renderActivos));
  const editarActivo = (id, cambios) => {
    const st = activosLocal(), a = st.activos.find((x) => x.id === id); if (!a) return;
    Object.assign(a, cambios); guardar({ activos: st.activos }); renderActivos();
  };
  $("#tabla-activos").addEventListener("change", (e) => {
    const s = e.target.closest("[data-novedad]"); if (!s) return;
    editarActivo(s.closest("tr").dataset.id, { novedad: s.value });
  });
  $("#tabla-activos").addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-id]"); if (!tr || !localActual) return;
    if (e.target.closest("[data-fotos]")) return abrirFotos(tr.dataset.id);
    if (e.target.closest("[data-editar]")) return abrirGestion("modificar", tr.dataset.id);
    if (!e.target.closest("[data-quitar-activo]")) return;
    const st = activosLocal(), i = st.activos.findIndex((x) => x.id === tr.dataset.id), local = localActual.local;
    const [quitado] = st.activos.splice(i, 1); guardar({ activos: st.activos }); renderActivos();
    toast(`Activo eliminado: ${quitado.nombre}`, "Deshacer", () => {
      if (!localActual || localActual.local !== local) return;
      const s2 = activosLocal(); s2.activos.splice(Math.min(i, s2.activos.length), 0, quitado); guardar({ activos: s2.activos }); renderActivos();
    });
  });
  $("#a-todos-ok").addEventListener("click", () => {
    if (!localActual) return toast("Selecciona un local primero");
    const st = activosLocal(), pend = st.activos.filter((a) => enUso(a) && !a.novedad);
    if (!pend.length) return toast("No hay activos pendientes");
    if (!confirm(`¿Marcar ${pend.length} activos pendientes como «${OK_ACT}»?`)) return;
    pend.forEach((a) => (a.novedad = OK_ACT)); guardar({ activos: st.activos }); renderActivos();
  });
  // ---------- Activo nuevo · Modificar equipo · Equipo perdido ----------
  // Los campos modificados guardan su valor anterior en a.cambios[campo] = { antes, nuevo, fecha } y se resaltan en amarillo pálido.
  const CAMPOS_MOD = { nombre: "Nombre", marca: "Marca", modelo: "Modelo", serie: "Serie" };
  function valorMod(a, campo) {
    const c = a.cambios && a.cambios[campo], v = esc(a[campo] || "");
    return c ? `<span class="mod" title="Antes: ${esc(c.antes || "(vacío)")}">${v || "(vacío)"}</span>` : v;
  }
  // Si el archivo del SIR o el respaldo vuelven a traer el valor original, el cambio del auditor se mantiene
  function aplicarCambios(a) {
    Object.entries(a.cambios || {}).forEach(([k, c]) => { if (a[k] !== c.nuevo) { c.antes = a[k]; a[k] = c.nuevo; } });
    return a;
  }
  const descCambios = (a) => Object.entries(a.cambios || {}).map(([k, c]) => `${CAMPOS_MOD[k].toLowerCase()}: ${c.antes || "(vacío)"} → ${c.nuevo || "(vacío)"}`).join("; ");
  // Las opciones del panel son las novedades de la matriz. «Equipo No creado» abre el formulario de activo nuevo,
  // «Modificar Marca Modelo Serie» el de modificación; las demás solo piden elegir el activo.
  const gTipo = () => ($('input[name="g-modo"]:checked') || {}).value || "Equipo No creado";
  const gModo = () => (/^equipo no creado/i.test(gTipo()) ? "nuevo" : gTipo() === "Modificar Marca Modelo Serie" ? "modificar" : "otra");
  $("#g-opciones").innerHTML = M.novedadesActivo.map((n) => `<label class="g-opcion${puntosNovedad(n) ? " resta" : ""}"><input type="radio" name="g-modo" value="${esc(n)}"${n === "Equipo No creado" ? " checked" : ""}>
    <span>${esc(n)}</span><small>${puntosNovedad(n) ? `−${puntosNovedad(n)}` : "0"} pts</small></label>`).join("");
  let gActivo = null, gFotos = []; // fotos pendientes de guardar: { id, blob, url }
  function renderGFotos() {
    $("#g-miniaturas").innerHTML = gFotos.map((f) => `<div class="foto" data-g-foto="${f.id}"><img src="${f.url}" alt="Foto nueva"><button type="button" data-g-quitar title="Quitar foto" aria-label="Quitar foto">✕</button></div>`).join("");
  }
  function limpiarGestion() {
    gFotos.forEach((f) => URL.revokeObjectURL(f.url)); gFotos = []; renderGFotos();
    $$("#g-nuevo input").forEach((i) => (i.value = ""));
    $("#g-obs").value = ""; $("#g-buscar").value = ""; gActivo = null;
    $$(".g-checks input").forEach((c) => (c.checked = false));
    renderGestion();
  }
  function renderGestion() {
    const modo = gModo(), elige = modo !== "nuevo";
    $("#g-elegir").hidden = !elige; $("#g-nuevo").hidden = elige; $("#g-modificar").hidden = modo !== "modificar" || !gActivo;
    $$(".g-opcion").forEach((l) => l.classList.toggle("sel", l.querySelector("input").checked));
    $("#g-guardar").textContent = modo === "nuevo" ? `Agregar activo (${gTipo()})` : modo === "modificar" ? "Guardar modificación" : `Registrar «${gTipo()}»`;
    $("#g-obs").placeholder = gTipo() === "Activo perdido" ? "¿Dónde debía estar? ¿Desde cuándo falta? (opcional)" : "Detalle para el informe (opcional)";
    $("#g-revertir").hidden = !(modo === "modificar" && gActivo && Object.keys(gActivo.cambios || {}).length);
    if (elige) buscarGestion();
    $("#g-actual").hidden = !gActivo;
    if (gActivo) {
      const a = gActivo;
      $("#g-actual").innerHTML = `<b>${valorMod(a, "nombre")}</b> · ${esc(a.codigo)}<br>Marca: ${valorMod(a, "marca") || "—"} · Modelo: ${valorMod(a, "modelo") || "—"} · Serie: ${valorMod(a, "serie") || "—"}
        ${a.novedad ? `<br><small class="muted">Novedad actual: ${esc(a.novedad)}</small>` : ""}`;
      // Un campo por cada casilla marcada, con el valor actual para editarlo
      $("#g-mod-campos").innerHTML = $$(".g-checks input:checked").map((c) => `<label>${CAMPOS_MOD[c.value]} nuevo
        <input data-mod="${c.value}" value="${esc(a[c.value] || "")}" placeholder="${CAMPOS_MOD[c.value]}"></label>`).join("");
    }
  }
  function buscarGestion() {
    if (gActivo) { $("#g-resultados").innerHTML = `<li><button type="button" class="ghost" data-g-cambiar>Elegir otro activo</button></li>`; return; }
    const q = $("#g-buscar").value.trim().toLowerCase();
    const lista = localActual ? activosLocal().activos.filter((a) => (enUso(a) || a.origen !== "SIR")
      && (!q || [a.codigo, a.erp, a.nombre, a.marca, a.modelo, a.serie].join(" ").toLowerCase().includes(q))).slice(0, 8) : [];
    $("#g-resultados").innerHTML = lista.length ? lista.map((a) => `<li><button type="button" class="ghost" data-g-activo="${a.id}">
      <b>${esc(a.nombre)}</b><small>${esc([a.codigo, [a.marca, a.modelo].filter(Boolean).join(" "), a.serie && `S: ${a.serie}`].filter(Boolean).join(" · "))}</small></button></li>`).join("")
      : `<li class="muted small">${localActual ? "Ningún activo coincide." : "Selecciona un local."}</li>`;
  }
  function abrirGestion(modo, idActivo) {
    $("#gestion").open = true;
    const tipo = modo === "modificar" ? "Modificar Marca Modelo Serie" : modo === "perdido" ? "Activo perdido" : modo === "nuevo" ? "Equipo No creado" : modo;
    $$('input[name="g-modo"]').forEach((r) => (r.checked = r.value === tipo));
    gActivo = idActivo ? activosLocal().activos.find((x) => x.id === idActivo) || null : null;
    renderGestion();
    $("#gestion").scrollIntoView({ behavior: "smooth", block: "start" });
  }
  $$('input[name="g-modo"]').forEach((r) => r.addEventListener("change", () => { gActivo = null; $("#g-buscar").value = ""; renderGestion(); }));
  $("#g-buscar").addEventListener("input", buscarGestion);
  $("#g-resultados").addEventListener("click", (e) => {
    if (e.target.closest("[data-g-cambiar]")) { gActivo = null; renderGestion(); $("#g-buscar").focus(); return; }
    const b = e.target.closest("[data-g-activo]"); if (!b) return;
    gActivo = activosLocal().activos.find((x) => x.id === b.dataset.gActivo); renderGestion();
  });
  $(".g-checks").addEventListener("change", renderGestion);
  // Mientras se escribe, el campo que cambia se pinta de amarillo pálido
  $("#g-mod-campos").addEventListener("input", (e) => {
    const i = e.target.closest("[data-mod]"); if (!i || !gActivo) return;
    i.classList.toggle("cambiado", i.value.trim() !== (gActivo[i.dataset.mod] || ""));
  });
  async function agregarGFotos(files) {
    for (const file of files) {
      if (!file.type.startsWith("image/")) continue;
      try { const blob = await reducirImagen(file); gFotos.push({ id: nuevoId(), blob, url: URL.createObjectURL(blob) }); } catch { toast(`No se pudo leer ${file.name}`); }
    }
    renderGFotos();
  }
  $("#g-agregar-foto").addEventListener("click", () => $("#g-input").click());
  $("#g-tomar").addEventListener("click", () => $("#g-camara").click());
  ["#g-input", "#g-camara"].forEach((s) => $(s).addEventListener("change", (e) => { agregarGFotos([...e.target.files]); e.target.value = ""; }));
  $("#g-miniaturas").addEventListener("click", (e) => {
    const box = e.target.closest("[data-g-foto]"); if (!box) return;
    if (e.target.closest("[data-g-quitar]")) { const f = gFotos.find((x) => x.id === box.dataset.gFoto); URL.revokeObjectURL(f.url); gFotos = gFotos.filter((x) => x !== f); renderGFotos(); return; }
    $("#f-grande").src = box.querySelector("img").src; $("#dlg-ver").showModal(); // clic = ampliar
  });
  async function guardarGFotos(idActivo) {
    const ids = [];
    for (const f of gFotos) { await fotosDB.put({ id: f.id, blob: f.blob, local: localActual.local, activo: idActivo, fecha: new Date().toISOString() }); ids.push(f.id); }
    return ids;
  }
  $("#g-guardar").addEventListener("click", async () => {
    if (!localActual) return toast("Selecciona un local primero");
    const modo = gModo(), st = activosLocal(), obs = $("#g-obs").value.trim(), hoyIso = new Date().toISOString().slice(0, 10);
    if (modo === "nuevo") {
      const v = (n) => $(`#g-nuevo [name="${n}"]`).value.trim();
      if (!v("nombre")) { toast("Escribe al menos el nombre del activo"); return $(`#g-nuevo [name="nombre"]`).focus(); }
      const a = { id: nuevoId(), codigo: v("codigo") || "SIN CÓDIGO", nombre: v("nombre"), marca: v("marca"), modelo: v("modelo"), serie: v("serie"),
        novedad: gTipo(), origen: "manual", estado: "Nuevo (no SIR)", creadoAuditor: hoyIso, obs, fotos: [] };
      a.fotos = await guardarGFotos(a.id);
      st.activos.push(a); guardar({ activos: st.activos });
      toast(`Activo agregado: ${a.nombre}`);
    } else {
      if (!gActivo) { toast("Primero elige el activo"); return $("#g-buscar").focus(); }
      const a = st.activos.find((x) => x.id === gActivo.id);
      if (modo === "modificar") {
        const nuevos = $$("#g-mod-campos [data-mod]").map((i) => [i.dataset.mod, i.value.trim()]).filter(([k, val]) => val !== (a[k] || ""));
        if (!$$(".g-checks input:checked").length) return toast("Marca qué vas a modificar: nombre, marca, modelo o serie");
        if (!nuevos.length) return toast("Escribe el valor nuevo (es igual al actual)");
        a.cambios = a.cambios || {};
        nuevos.forEach(([k, val]) => {
          const antes = a.cambios[k] ? a.cambios[k].antes : a[k] || "";
          if (val === antes) delete a.cambios[k]; else a.cambios[k] = { antes, nuevo: val, fecha: hoyIso };
          a[k] = val;
        });
        if (Object.keys(a.cambios).length && (!a.novedad || a.novedad === OK_ACT)) a.novedad = "Modificar Marca Modelo Serie";
        toast(`Modificado: ${a.nombre} (${nuevos.map(([k]) => CAMPOS_MOD[k].toLowerCase()).join(", ")})`);
      } else {
        a.novedad = gTipo();
        if (a.novedad === "Activo perdido") a.perdido = { fecha: hoyIso, obs }; else delete a.perdido;
        toast(`${a.nombre}: ${a.novedad}`);
      }
      if (obs) a.obs = obs;
      a.fotos = a.fotos.concat(await guardarGFotos(a.id));
      guardar({ activos: st.activos });
    }
    gFotos.forEach((f) => URL.revokeObjectURL(f.url)); gFotos = [];
    limpiarGestion(); renderActivos();
  });
  $("#g-limpiar").addEventListener("click", limpiarGestion);
  $("#g-revertir").addEventListener("click", () => {
    if (!gActivo || !confirm("¿Volver a los datos originales de este activo?")) return;
    const st = activosLocal(), a = st.activos.find((x) => x.id === gActivo.id);
    Object.entries(a.cambios || {}).forEach(([k, c]) => (a[k] = c.antes));
    delete a.cambios; if (a.novedad === "Modificar Marca Modelo Serie") a.novedad = "";
    guardar({ activos: st.activos }); gActivo = a; renderGestion(); renderActivos(); toast("Cambios revertidos");
  });
  $("#a-sir").addEventListener("input", () => { guardar({ sir: +$("#a-sir").value || 0 }); renderActivos(); });

  // ---------- Carga del archivo «Activos fijos» del SIR ----------
  // Tabla HTML con encabezados Código, ERP, Nombre, Marca, Modelo, Serie, …, Fecha Creación, Estado.
  // Tildes dañadas («REMODELACIÃ“N»): el SIR guardó UTF-8 leído como Windows-1252. Se vuelven a bytes y se decodifican bien.
  const CP1252 = { 0x20ac: 0x80, 0x201a: 0x82, 0x0192: 0x83, 0x201e: 0x84, 0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x02c6: 0x88, 0x2030: 0x89, 0x0160: 0x8a,
    0x2039: 0x8b, 0x0152: 0x8c, 0x017d: 0x8e, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93, 0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97, 0x02dc: 0x98,
    0x2122: 0x99, 0x0161: 0x9a, 0x203a: 0x9b, 0x0153: 0x9c, 0x017e: 0x9e, 0x0178: 0x9f };
  const utf8Estricto = new TextDecoder("utf-8", { fatal: true });
  function arreglarTexto(s) {
    if (!/[ÃÂ]/.test(s)) return s;
    const bytes = [];
    for (const ch of s) {
      const c = ch.codePointAt(0), b = c < 256 ? c : CP1252[c];
      if (b === undefined) return s;
      bytes.push(b);
    }
    try { return utf8Estricto.decode(new Uint8Array(bytes)); } catch { return s; }
  }
  function leerActivosSIR(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const todo = doc.body ? doc.body.textContent.replace(/\s+/g, " ") : "";
    const local = ((/Local:\s*(.+?)(?=\s*Usuario:|$)/.exec(todo) || [])[1] || "").trim();
    const fecha = (/Fecha:\s*(\d{4}-\d{2}-\d{2}|\d{1,2}\/\d{1,2}\/\d{2,4})/.exec(todo) || [])[1] || "";
    const norm = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
    let idx = null, filas = [], duplicados = 0;
    const vistos = new Set();
    doc.querySelectorAll("tr").forEach((tr) => {
      const c = [...tr.children].map((td) => arreglarTexto(td.textContent.replace(/\s+/g, " ").trim()));
      if (!idx) {
        const h = c.map(norm);
        if (h.includes("codigo") && h.includes("serie")) idx = { codigo: h.indexOf("codigo"), erp: h.indexOf("erp"), nombre: h.indexOf("nombre"), marca: h.indexOf("marca"),
          modelo: h.indexOf("modelo"), serie: h.indexOf("serie"), fecha: h.findIndex((x) => x.startsWith("fecha")), estado: h.indexOf("estado"), n: c.length };
        return;
      }
      if (c.length !== idx.n || !c[idx.codigo]) return;
      const v = (k) => (idx[k] >= 0 ? c[idx[k]] : "");
      const a = { codigo: v("codigo"), erp: v("erp"), nombre: v("nombre"), marca: v("marca"), modelo: v("modelo"), serie: v("serie"), fecha: v("fecha"), estado: v("estado") || "Activo" };
      const clave = `${a.codigo}|${a.serie}|${a.estado}`;
      if (vistos.has(clave)) { duplicados++; return; }
      vistos.add(clave); filas.push(a);
    });
    return { local, fecha, filas, duplicados };
  }
  async function cargarActivosSIR(file) {
    if (!file) return;
    let datos;
    try { datos = leerActivosSIR(await leerArchivoTexto(file)); } catch { datos = { filas: [] }; }
    if (!datos.filas.length) { toast("No reconozco el archivo: debe ser «Activos fijos» exportado del SIR"); return; }
    const codigo = (/^([A-Z0-9]{3,5})/i.exec(datos.local) || [])[1] || "";
    const delArchivo = codigo && locales.find((l) => l.local.toUpperCase().startsWith(codigo.toUpperCase() + " "));
    if (delArchivo && (!localActual || localActual.local !== delArchivo.local)) {
      if (!localActual || confirm(`El archivo es del local ${datos.local}. ¿Cambiar a ${delArchivo.local}?`)) seleccionarLocal(delArchivo.local, false);
    }
    if (!localActual) { toast("Selecciona primero el local del archivo"); return; }
    // Se conservan la novedad y las fotos de los activos que ya estaban (mismo código, serie y estado)
    // La clave usa los valores originales del SIR: si el auditor cambió la serie, igual se reconoce el activo
    const orig = (a, k) => (a.cambios && a.cambios[k] ? a.cambios[k].antes : a[k]);
    const st = activosLocal(), clave = (a) => `${orig(a, "codigo")}|${orig(a, "serie")}|${a.estado}`;
    const previos = new Map(st.activos.filter((a) => a.origen === "SIR").map((a) => [clave(a), a]));
    const nuevos = datos.filas.map((f) => { const p = previos.get(clave(f)); previos.delete(clave(f)); return aplicarCambios(Object.assign({ id: nuevoId(), novedad: "", fotos: [] }, p, f, { origen: "SIR" })); });
    const conservados = [...previos.values()].filter((a) => a.novedad || a.fotos.length); // ya no están en el SIR pero tienen trabajo hecho
    const manuales = st.activos.filter((a) => a.origen !== "SIR");
    const activos = nuevos.concat(conservados, manuales);
    const enActivo = new Set(datos.filas.filter((f) => f.estado !== "Transferido").map((f) => f.codigo)).size;
    guardar({ activos, sir: enActivo, archivoActivos: { nombre: file.name, local: datos.local, fecha: datos.fecha, activos: datos.filas.filter(enUso).length,
      transferidos: datos.filas.filter((f) => !enUso(f)).length, duplicados: datos.duplicados } });
    renderActivos();
    toast(`${datos.filas.filter(enUso).length} activos cargados de ${datos.local}`);
  }
  $("#a-subir").addEventListener("click", () => $("#a-archivo").click());
  $("#a-archivo").addEventListener("change", (e) => { cargarActivosSIR(e.target.files[0]); e.target.value = ""; });
  const zonaA = $("#a-carga");
  ["dragenter", "dragover"].forEach((t) => zonaA.addEventListener(t, (e) => { e.preventDefault(); zonaA.classList.add("arrastre"); }));
  ["dragleave", "drop"].forEach((t) => zonaA.addEventListener(t, () => zonaA.classList.remove("arrastre")));
  zonaA.addEventListener("drop", (e) => { e.preventDefault(); cargarActivosSIR(e.dataTransfer.files[0]); });

  // ---------- Respaldo .json de la app de inventario ----------
  // Formato: { v, items: [{ id, nombre, codigo, estado, local, ubicacion, categoria, marca, modelo, serie, notas, novedad, novedadDetalle,
  //   fotos: [dataURL], hist, libros: { "AAAA-MM": valor } }], sueltas: [{ id, local, nombre, foto }] }
  // Se cruza con el listado del SIR: el código de la app usa _D_ y el del SIR _I_, por eso se compara sin esa letra; si no, por serie.
  const codNorm = (c) => String(c || "").toUpperCase().replace(/_[A-Z]_/, "_").trim();
  const serieValida = (s) => !!s && !/^(S\/?[NSM]|N\/?A|SIN SERIE|NO TIENE|-+|0+)$/i.test(String(s).trim());
  const localDeCodigo = (c) => locales.find((l) => l.local.toUpperCase().startsWith(String(c).toUpperCase().trim() + " "));
  const limpiarNotas = (s) => String(s || "").split("\n").filter((l) => !/:\s*(undefined)?\s*$/.test(l)).join("\n").trim();
  let respaldo = null; // { nombre, porLocal: Map(nombreLocal → items), sueltas: Map(nombreLocal → fotos) }

  function cruzar(st, items) {
    const sir = st.activos.filter((a) => a.origen === "SIR");
    const porCod = new Map(sir.map((a) => [codNorm(a.codigo), a]));
    const porSerie = new Map(sir.filter((a) => serieValida(a.serie)).map((a) => [a.serie.toUpperCase().trim(), a]));
    return items.map((it) => ({ it, a: porCod.get(codNorm(it.codigo)) || (serieValida(it.serie) && porSerie.get(it.serie.toUpperCase().trim())) || null }));
  }
  function resumenJson() {
    const nombre = $("#jd-local").value, items = respaldo.porLocal.get(nombre) || [];
    const st = store.get(claveLocal(nombre), { activos: [] }); st.activos = st.activos || [];
    const pares = cruzar(st, items), enc = pares.filter((p) => p.a).length, sirEnUso = st.activos.filter((a) => a.origen === "SIR" && enUso(a)).length;
    const nFotos = items.reduce((s, x) => s + (x.fotos || []).length, 0) + (respaldo.sueltas.get(nombre) || []).length;
    $("#jd-resumen").innerHTML = [[items.length, "En la app"], [enc, "Coinciden con el SIR"], [items.length - enc, "Solo en la app"],
      [sirEnUso ? Math.max(0, sirEnUso - pares.filter((p) => p.a && enUso(p.a)).length) : "—", "Del SIR no están en la app"]]
      .map(([v, t]) => `<div><b>${v}</b><span>${t}</span></div>`).join("");
    $("#jd-nfotos").textContent = `(${nFotos} en este local)`;
    $("#jd-aviso").textContent = sirEnUso ? "" : "Este local aún no tiene el listado del SIR: se agregarán los activos de la app como lista. Si luego subes el .xls del SIR, se cruzan solos.";
  }
  async function dataUrlABlob(u) { return (await fetch(u)).blob(); }
  // Importa un local: actualiza los activos del SIR que coinciden, agrega los que solo están en la app y guarda las fotos en IndexedDB
  async function importarLocal(nombre, items, fotosSueltas, opciones) {
    const st = store.get(claveLocal(nombre), { marcados: [], activos: [], sir: 0 });
    st.activos = (st.activos || []).map((a) => Object.assign({ id: a.id || nuevoId(), fotos: [] }, a));
    let nFotos = 0, nuevos = 0, actualizados = 0;
    for (const { it, a: encontrado } of cruzar(st, items)) {
      let a = encontrado;
      if (!a) {
        if (!opciones.solo) continue;
        a = st.activos.find((x) => x.origen === "app" && x.appId === it.id);
        if (!a) { a = { id: nuevoId(), origen: "app", estado: "Solo en app", novedad: "", fotos: [] }; st.activos.push(a); nuevos++; }
        Object.assign(a, { appId: it.id, codigo: it.codigo, nombre: it.nombre, marca: it.marca, modelo: it.modelo, serie: it.serie });
        aplicarCambios(a);
      } else actualizados++;
      const libros = it.libros && Object.keys(it.libros).sort().pop();
      a.app = { id: it.id, estado: it.estado || "", ubicacion: it.ubicacion || "", novedad: !!it.novedad, detalle: it.novedadDetalle || "",
        notas: limpiarNotas(it.notas), hist: it.hist || [], libros: libros ? it.libros[libros] : null, librosMes: libros || "" };
      if (!a.novedad) {
        if (it.novedad) a.novedad = /marca|modelo|serie/i.test(it.novedadDetalle || "") ? "Modificar Marca Modelo Serie" : "";
        else if (opciones.ok && a.origen === "SIR" && enUso(a)) a.novedad = OK_ACT;
      }
      if (opciones.fotos) for (const [i, u] of (it.fotos || []).entries()) {
        const id = `app-${it.id}-${i}`; // id fijo: volver a importar no duplica fotos
        if (a.fotos.includes(id)) continue;
        try { await fotosDB.put({ id, blob: await dataUrlABlob(u), local: nombre, activo: a.id, fecha: new Date().toISOString() }); a.fotos.push(id); nFotos++; } catch { /* foto dañada */ }
      }
    }
    if (opciones.fotos) {
      st.fotosSueltas = st.fotosSueltas || [];
      for (const s of fotosSueltas) {
        const id = `suelta-${s.id}`;
        if (st.fotosSueltas.includes(id)) continue;
        try { await fotosDB.put({ id, blob: await dataUrlABlob(s.foto), local: nombre, activo: "sueltas", nombre: s.nombre, fecha: new Date().toISOString() }); st.fotosSueltas.push(id); nFotos++; } catch { /* foto dañada */ }
      }
    }
    st.respaldoApp = { nombre: respaldo.nombre, fecha: new Date().toLocaleDateString("es-EC"), items: items.length, actualizados, nuevos };
    store.set(claveLocal(nombre), st);
    return { actualizados, nuevos, nFotos };
  }
  async function leerRespaldo(file) {
    if (!file) return;
    toast("Leyendo respaldo…");
    let d;
    try { d = JSON.parse(await file.text()); } catch { return toast("El archivo .json no se pudo leer"); }
    if (!d || !Array.isArray(d.items)) return toast("No reconozco el archivo: debe ser el respaldo de la app de inventario");
    const porLocal = new Map(), sueltas = new Map(), sinLocal = new Set();
    d.items.forEach((it) => { const l = localDeCodigo(it.local); if (!l) return sinLocal.add(it.local); if (!porLocal.has(l.local)) porLocal.set(l.local, []); porLocal.get(l.local).push(it); });
    (d.sueltas || []).forEach((s) => { const l = localDeCodigo(s.local); if (!l) return; if (!sueltas.has(l.local)) sueltas.set(l.local, []); sueltas.get(l.local).push(s); });
    if (!porLocal.size) return toast("El respaldo no tiene activos de locales de la matriz");
    respaldo = { nombre: file.name, porLocal, sueltas };
    const nombres = [...porLocal.keys()].sort((a, b) => a.localeCompare(b, "es"));
    $("#jd-local").innerHTML = nombres.map((n) => `<option value="${esc(n)}">${esc(n)} · ${porLocal.get(n).length} activos</option>`).join("");
    $("#jd-local").value = localActual && porLocal.has(localActual.local) ? localActual.local : nombres[0];
    $("#jd-archivo").textContent = `${file.name} · ${d.items.length} activos de ${porLocal.size} locales${sinLocal.size ? ` (sin local en la matriz: ${[...sinLocal].join(", ")})` : ""}`;
    $("#jd-todos").hidden = porLocal.size < 2;
    $("#jd-todos").textContent = `Importar los ${porLocal.size} locales`;
    resumenJson();
    $("#dlg-json").showModal();
  }
  async function ejecutarImportacion(nombres) {
    const opciones = { ok: $("#jd-ok").checked, fotos: $("#jd-fotos").checked, solo: $("#jd-solo").checked };
    $("#jd-importar").disabled = $("#jd-todos").disabled = true;
    const tot = { actualizados: 0, nuevos: 0, nFotos: 0 };
    try {
      for (const n of nombres) {
        toast(`Importando ${n}…`);
        const r = await importarLocal(n, respaldo.porLocal.get(n), respaldo.sueltas.get(n) || [], opciones);
        Object.keys(tot).forEach((k) => (tot[k] += r[k]));
      }
    } finally { $("#jd-importar").disabled = $("#jd-todos").disabled = false; }
    $("#dlg-json").close();
    if (!localActual || !nombres.includes(localActual.local)) seleccionarLocal(nombres[0], false); else renderActivos();
    renderActivos();
    toast(`${tot.actualizados} activos enlazados, ${tot.nuevos} solo en la app, ${tot.nFotos} fotos importadas`);
  }
  $("#j-subir").addEventListener("click", () => $("#j-archivo").click());
  $("#j-archivo").addEventListener("change", (e) => { leerRespaldo(e.target.files[0]); e.target.value = ""; });
  $("#jd-local").addEventListener("change", resumenJson);
  $("#jd-cerrar").addEventListener("click", () => $("#dlg-json").close());
  $("#jd-importar").addEventListener("click", () => ejecutarImportacion([$("#jd-local").value]));
  $("#jd-todos").addEventListener("click", () => ejecutarImportacion([...respaldo.porLocal.keys()]));
  const zonaJ = $("#j-carga");
  ["dragenter", "dragover"].forEach((t) => zonaJ.addEventListener(t, (e) => { e.preventDefault(); zonaJ.classList.add("arrastre"); }));
  ["dragleave", "drop"].forEach((t) => zonaJ.addEventListener(t, () => zonaJ.classList.remove("arrastre")));
  zonaJ.addEventListener("drop", (e) => { e.preventDefault(); leerRespaldo(e.dataTransfer.files[0]); });
  $("#j-sueltas").addEventListener("click", () => abrirAsignar());

  // ---------- Fichas de activos (vista imprimible → «Guardar como PDF») ----------
  let urlsFichas = [];
  // Estado tal como lo dice el archivo: el del listado del SIR (Activo, Transferido, Por exportar, Exportado…);
  // si el activo vino solo del respaldo, el «estado:» de sus notas.
  // Novedades del auditor: equipo por crear, por modificar (qué campos) o perdido
  const esNovedadFicha = (a) => !!a.creadoAuditor || !!(a.cambios && Object.keys(a.cambios).length) || a.novedad === "Activo perdido";
  function estadoNovedad(a) {
    const baja = bajasLocal().find((b) => b.activoId === a.id);
    if (baja) return bajaPendiente(baja) ? "Dar de baja (pendiente de aprobación)" : "Dado de baja";
    if (a.creadoAuditor) return "Crear equipo";
    const campos = Object.keys(a.cambios || {}).map((k) => CAMPOS_MOD[k].toLowerCase());
    if (campos.length) return `Modificar ${unirY(campos)}`;
    if (a.novedad === "Activo perdido") return "Activo perdido";
    return "";
  }
  function estadoFicha(a) {
    const nov = estadoNovedad(a); if (nov) return nov;
    if (a.origen === "SIR" && a.estado) return a.estado;
    const m = /estado\s*:\s*([^\n]+)/i.exec((a.app && a.app.notas) || "");
    if (m && !/undefined/i.test(m[1])) return m[1].trim();
    return a.estado && a.estado !== "Solo en app" && a.estado !== "Manual" ? a.estado : "—";
  }
  function activosParaFichas() {
    const st = activosLocal(), que = $("#vf-que").value;
    if (que === "solo") return st.activos.filter((a) => fichasSolo.includes(a.id)).sort((x, y) => x.nombre.localeCompare(y.nombre, "es"));
    const lista = que === "filtro" ? ultimaVistaActivos : st.activos.filter((a) => enUso(a) || a.origen === "app");
    return (que === "novedades" ? st.activos.filter((a) => enUso(a) || a.origen !== "SIR").filter(esNovedadFicha)
      : que === "nov" ? lista.filter(conNovedadActivo) : que === "fotos" ? lista.filter((a) => a.fotos.length) : lista)
      .slice().sort((x, y) => x.nombre.localeCompare(y.nombre, "es"));
  }
  async function armarFichas() {
    if (!localActual) return;
    urlsFichas.forEach(URL.revokeObjectURL); urlsFichas = [];
    await cargarLibros().catch(() => null);
    const lista = activosParaFichas(), pp = +$("#vf-pp").value, st = activosLocal();
    const hoyTxt = new Date().toLocaleDateString("es-EC", { day: "numeric", month: "long", year: "numeric" });
    $("#vf-info").textContent = `${localActual.local} · ${lista.length} fichas`;
    // Hasta 3 fotos por ficha, igual que el formato de papel
    const fotos = await Promise.all(lista.map(async (a) => (await Promise.all(a.fotos.slice(0, 3).map((id) => fotosDB.get(id).catch(() => null))))
      .filter(Boolean).map((f) => { const u = URL.createObjectURL(f.blob); urlsFichas.push(u); return u; })));
    const ficha = (a, i) => {
      const x = a.app || {}, extra = a.fotos.length > 3 ? ` (+${a.fotos.length - 3} en la app)` : "";
      const notas = (x.notas || "").split("\n").filter((l) => !/^\s*estado\s*:/i.test(l)).join(" · "); // el estado ya va en su fila
      // Campo modificado (o activo nuevo): valor en amarillo pálido y, debajo, el valor anterior tachado
      const celda = (k) => {
        const c = a.cambios && a.cambios[k], v = esc(a[k] || "—");
        if (c) return `<span class="mod">${v}</span><span class="antes">Antes: ${esc(c.antes || "(vacío)")}</span>`;
        return a.creadoAuditor ? `<span class="mod">${v}</span>` : v;
      };
      const filas = [
        ["NOMBRE", celda("nombre")], ["MARCA", celda("marca")], ["MODELO", celda("modelo")], ["SERIE", celda("serie")],
        ["ESTADO", estadoNovedad(a) ? `<b class="${a.novedad === "Activo perdido" && !a.creadoAuditor && !a.cambios ? "perdido" : "mod"}">${esc(estadoFicha(a))}</b>` : `<b>${esc(estadoFicha(a))}</b>`],
      ];
      const bj = bajasLocal().find((b) => b.activoId === a.id), lb = libroDe(a, localActual.local);
      // valor en libros de la base contable (si la baja tiene un valor escrito a mano, manda ese)
      if (lb && !(bj && bj.valor !== "" && bj.valor != null)) filas.push(["VALOR EN LIBROS", `<b>${fmtUSD(valorLibro(lb) || 0)}</b> <small>a ${esc(mesLibros())}${lb.a != null ? ` · adquisición ${fmtUSD(lb.a)}` : ""}${lb.f ? ` · compra ${esc(lb.f)}` : ""}${lb.b ? ` · baja contable ${esc(lb.b)}` : ""}</small>`]);
      if (bj) filas.push(...(bj.valor !== "" && bj.valor != null ? [["VALOR EN LIBROS", `<b>${fmtUSD(+bj.valor)}</b>`]] : lb ? [] : [["VALOR EN LIBROS", "<b>—</b>"]]), ["MOTIVO DE BAJA", esc(bj.motivo + (bj.obs ? ` · ${bj.obs}` : ""))],
        ["CERTIFICA", esc(`Subgerente ${bj.gerente} · Jefe de área ${bj.jefe}`)]);
      if (a.obs) filas.push(["NOTA AUDITORÍA", `<div class="clamp">${esc(a.obs)}</div>`]);
      if (x.ubicacion) filas.push(["UBICACIÓN", esc(x.ubicacion)]);
      filas.push(["DATOS SIR", `<div class="clamp">${esc([a.erp && `ERP ${a.erp}`, a.fecha && `Creado ${a.fecha}`, x.libros != null && `Valor en libros ${fmtUSD(x.libros)} (${x.librosMes})`].filter(Boolean).join(" · ") || "—")}</div>`]);
      if (notas) filas.push(["OBSERVACIÓN", `<div class="clamp">${esc(notas)}</div>`]);
      return `<article class="fa">
        <button type="button" class="ghost sm fa-copiar" data-copiar-ficha title="Copiar esta ficha como imagen">📋 Copiar como imagen</button>
        <div class="fa-cab"><div class="tit"><h3>Ficha de activo fijo</h3><small>${esc(localActual.local)} · ${esc(hoyTxt)} · Ficha ${i + 1} de ${lista.length}</small></div>
          <div class="fa-cod"><b>Código del activo</b><span class="cod">${esc(a.codigo || "—")}</span></div></div>
        <div class="fa-sec">IMÁGENES DEL ACTIVO${extra}</div>
        <div class="fa-fotos">${fotos[i].length ? fotos[i].map((u, j) => `<figure><img src="${u}" alt="${esc(a.nombre)} foto ${j + 1}" data-ampliar><figcaption>Foto ${j + 1}</figcaption></figure>`).join("")
          : `<div class="sin">Sin fotos registradas</div>`}</div>
        <table>${filas.map(([k, v]) => `<tr><th>${k}:</th><td>${v}</td></tr>`).join("")}</table>
      </article>`;
    };
    // Portada con el resumen y el índice (estado tal como viene en el archivo)
    const porEstado = lista.reduce((m, a) => ((m[estadoFicha(a)] = (m[estadoFicha(a)] || 0) + 1), m), {});
    const portada = `<section class="vf-hoja portada">
      <h2>${$("#vf-que").value === "novedades" ? "Novedades de activos fijos" : $("#vf-que").value === "solo" ? "Activos por dar de baja" : "Inventario de activos fijos"}</h2>
      <div class="muted">${esc(localActual.local)} · Jefe de área: ${esc(localActual.jefe || "")} · Auditor: ${esc(usuario.nombre)} · ${esc(hoyTxt)}</div>
      <div class="kp"><div><b>${st.sir || 0}</b>Registrados SIR</div><div><b>${lista.length}</b>Fichas</div>
        <div><b>${lista.filter((a) => a.fotos.length).length}</b>Con fotos</div><div><b>${lista.filter((a) => !a.fotos.length).length}</b>Sin fotos</div></div>
      <p class="muted">Por estado: ${esc(Object.entries(porEstado).map(([k, n]) => `${k} ${n}`).join(" · "))}</p>
      <table><thead><tr><th>#</th><th>Código</th><th>Nombre</th><th>Marca / modelo</th><th>Serie</th><th>Estado</th><th>Fotos</th></tr></thead><tbody>
      ${lista.map((a, i) => `<tr><td>${i + 1}</td><td>${esc(a.codigo)}</td><td>${esc(a.nombre)}</td><td>${esc([a.marca, a.modelo].filter(Boolean).join(" "))}</td>
        <td>${esc(a.serie)}</td><td>${esc(estadoFicha(a))}</td><td>${a.fotos.length}</td></tr>`).join("")}</tbody></table></section>`;
    const hojas = [];
    for (let i = 0; i < lista.length; i += pp) hojas.push(`<section class="vf-hoja ${pp === 1 ? "una" : "dos"}">${lista.slice(i, i + pp).map((a, j) => ficha(a, i + j)).join("")}</section>`);
    $("#vf-hojas").innerHTML = lista.length ? portada + hojas.join("") : `<p class="muted" style="text-align:center">No hay activos con ese criterio.</p>`;
  }
  let fichasSolo = []; // activos puntuales (p. ej. los de baja) para la vista de fichas
  function abrirFichasDe(ids) { fichasSolo = ids; $("#vf-que").value = "solo"; abrirFichas(); }
  async function abrirFichas() {
    if (!localActual) return toast("Selecciona un local primero");
    if (!activosLocal().activos.length) return toast("Primero sube el listado de activos del SIR o el respaldo .json");
    $("#vista-fichas").hidden = false; document.body.style.overflow = "hidden";
    toast("Armando fichas…");
    await armarFichas();
  }
  function cerrarFichas() {
    $("#vista-fichas").hidden = true; document.body.style.overflow = "";
    urlsFichas.forEach(URL.revokeObjectURL); urlsFichas = []; $("#vf-hojas").innerHTML = "";
  }
  $("#a-fichas").addEventListener("click", () => { $("#vf-que").value = "todos"; abrirFichas(); });
  // PDF solo con las fichas de equipos por crear, por modificar y perdidos
  $("#a-fichas-nov").addEventListener("click", () => {
    if (localActual && !activosLocal().activos.some(esNovedadFicha)) return toast("No hay equipos por crear, modificar ni perdidos en este local");
    $("#vf-que").value = "novedades"; abrirFichas();
  });
  $("#vf-cerrar").addEventListener("click", cerrarFichas);
  ["#vf-que", "#vf-pp"].forEach((s) => $(s).addEventListener("change", armarFichas));
  $("#vf-pdf").addEventListener("click", async () => {
    toast("Preparando PDF…");
    await Promise.all($$("#vf-hojas img").map((i) => i.decode().catch(() => {}))); // que todas las fotos estén cargadas antes de imprimir
    const titulo = document.title;
    document.title = `${$("#vf-que").value === "novedades" ? "Novedades activos" : "Fichas activos"} ${localActual.local}`.replace(/[\\/:*?"<>|]/g, "_");
    document.body.classList.add("imprimiendo");
    window.print();
    document.body.classList.remove("imprimiendo"); document.title = titulo;
  });
  // Clic en cualquier foto de las fichas → se amplía
  $("#vf-hojas").addEventListener("click", async (e) => {
    const copiar = e.target.closest("[data-copiar-ficha]");
    if (copiar) {
      const art = copiar.closest(".fa"), cod = art.querySelector(".cod").textContent;
      const blob = await new Promise((ok) => dibujarFicha(art).toBlob(ok, "image/png"));
      try { await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]); toast("Ficha copiada como imagen: pégala con Ctrl+V"); }
      catch { descargar(blob, `Ficha ${cod}.png`.replace(/[\\/:*?"<>|]/g, "_")); toast("Tu navegador no permitió copiar: se descargó la imagen"); }
      return;
    }
    const img = e.target.closest("img[data-ampliar]"); if (!img) return;
    $("#f-grande").src = img.src; $("#f-grande").alt = img.alt; $("#dlg-ver").showModal();
  });
  // Dibuja una ficha (tal como se ve en pantalla) en un lienzo: encabezado, fotos y datos.
  // Lee la ficha ya armada para que la imagen sea igual a la vista y al PDF.
  function dibujarFicha(art) {
    const W = 900, S = ESCALA_COPIA, F = '"Segoe UI", Arial, sans-serif', NEGRO = "#1b1f24", GRIS = "#646d78", BORDE = "#c9ced6";
    const med = document.createElement("canvas").getContext("2d");
    const partir = (t, ancho, font) => {
      med.font = font; const out = []; let act = "";
      String(t).split(/\s+/).forEach((p) => { const pr = act ? `${act} ${p}` : p; if (med.measureText(pr).width > ancho && act) { out.push(act); act = p; } else act = pr; });
      if (act) out.push(act); return out.length ? out : [""];
    };
    const imgs = [...art.querySelectorAll(".fa-fotos img")].filter((i) => i.complete && i.naturalWidth);
    const KW = 170, VX = KW + 14, VW = W - VX - 16, fuenteV = `400 15px ${F}`;
    const filas = [...art.querySelectorAll("table tr")].map((tr) => {
      const td = tr.querySelector("td"), mod = td.querySelector(".mod"), antes = td.querySelector(".antes"), perdido = td.querySelector(".perdido");
      const valor = (mod ? mod.textContent : td.innerText.replace(antes ? antes.textContent : "", "")).trim();
      const lineas = partir(valor, VW, td.querySelector("b") ? `700 15px ${F}` : fuenteV);
      return { k: tr.querySelector("th").textContent, lineas, mod: !!mod, antes: antes ? antes.textContent : "", rojo: !!perdido, negrita: !!td.querySelector("b"),
        alto: lineas.length * 21 + (antes ? 18 : 0) + 14 };
    });
    const CAB = 96, SEC = 30, FOT = imgs.length ? 250 : 120;
    const H = CAB + SEC + FOT + filas.reduce((s, f) => s + f.alto, 0) + 2;
    const cv = document.createElement("canvas"); cv.width = W * S; cv.height = H * S;
    const g = cv.getContext("2d"); g.scale(S, S); g.imageSmoothingQuality = "high";
    const txt = (t, x, y, px, peso = 400, color = NEGRO, al = "left") => { g.font = `${peso} ${px}px ${F}`; g.fillStyle = color; g.textAlign = al; g.fillText(t, x, y); };
    const linea = (x1, y1, x2, y2, c = NEGRO, w = 1.5) => { g.strokeStyle = c; g.lineWidth = w; g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke(); };
    g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
    // Encabezado: barra roja, título y recuadro del código
    const CODW = 280;
    g.fillStyle = "#c8102e"; g.fillRect(0, 0, 10, CAB);
    txt(art.querySelector(".tit h3").textContent.toUpperCase(), 28, 44, 24, 700);
    txt(art.querySelector(".tit small").textContent, 28, 72, 14, 500, "#4a525c");
    g.fillStyle = "#f5f6f8"; g.fillRect(W - CODW, 0, CODW, CAB); linea(W - CODW, 0, W - CODW, CAB);
    txt("CÓDIGO DEL ACTIVO", W - CODW / 2, 36, 11, 700, GRIS, "center");
    txt(art.querySelector(".cod").textContent, W - CODW / 2, 68, 22, 700, NEGRO, "center");
    linea(0, CAB, W, CAB);
    // Sección de fotos
    g.fillStyle = "#eef0f3"; g.fillRect(0, CAB, W, SEC); linea(0, CAB + SEC, W, CAB + SEC, NEGRO, 1);
    txt(art.querySelector(".fa-sec").textContent, W / 2, CAB + 20, 13, 700, NEGRO, "center");
    const yF = CAB + SEC;
    if (imgs.length) {
      const gap = 12, bw = (W - 2 * 16 - 2 * gap) / 3, bh = 210;
      imgs.slice(0, 3).forEach((im, i) => {
        const x = 16 + i * (bw + gap), y = yF + 12;
        const k = Math.max(bw / im.naturalWidth, bh / im.naturalHeight), sw = bw / k, sh = bh / k; // recorte tipo «cover»
        g.drawImage(im, (im.naturalWidth - sw) / 2, (im.naturalHeight - sh) / 2, sw, sh, x, y, bw, bh);
        g.strokeStyle = "#b8bec7"; g.lineWidth = 1; g.strokeRect(x, y, bw, bh);
        txt(`Foto ${i + 1}`, x + bw / 2, y + bh + 18, 12, 400, GRIS, "center");
      });
    } else {
      g.setLineDash([4, 4]); g.strokeStyle = BORDE; g.strokeRect(16, yF + 12, W - 32, FOT - 24); g.setLineDash([]);
      txt("Sin fotos registradas", W / 2, yF + FOT / 2 + 5, 15, 400, GRIS, "center");
    }
    linea(0, yF + FOT, W, yF + FOT);
    // Datos
    let y = yF + FOT;
    filas.forEach((f) => {
      g.fillStyle = "#f5f6f8"; g.fillRect(0, y, KW, f.alto); linea(KW, y, KW, y + f.alto, BORDE, 1);
      txt(`${f.k}`, 14, y + 25, 13, 700);
      f.lineas.forEach((l, i) => {
        const yy = y + 25 + i * 21;
        if (f.mod) { g.font = `${f.negrita ? 700 : 400} 15px ${F}`; const w = g.measureText(l).width; g.fillStyle = "#fff3b0"; g.fillRect(VX - 4, yy - 16, w + 8, 22); }
        txt(l, VX, yy, 15, f.negrita ? 700 : 400, f.rojo ? "#b42318" : f.mod ? "#3d3000" : NEGRO);
      });
      if (f.antes) txt(f.antes, VX, y + 25 + f.lineas.length * 21, 12, 400, GRIS);
      y += f.alto; linea(0, y, W, y, BORDE, 1);
    });
    g.strokeStyle = NEGRO; g.lineWidth = 2; g.strokeRect(1, 1, W - 2, H - 2);
    return cv;
  }
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !$("#vista-fichas").hidden && !$("#dlg-ver").open) cerrarFichas(); });

  // ---------- Cumplimiento de activos y cuentas por cobrar (todos los locales de la zona del auditor) ----------
  // Cumplimiento por mes: matriz:<usuario>:cumpl:<AAAA-MM> = { local: { fecha, equipo, estatus } }
  // Cuentas por cobrar:   matriz:<usuario>:cxc = { local: { "AAAA-MM": valor } }
  const MES_NOM = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const ENF_ZONA = ["Cumplimiento Inventario Activos", "Análisis Cuentas por Cobrar"];
  const claveCu = (mes) => `matriz:${usuario.usuario}:cumpl:${mes}`;
  const claveCx = () => `matriz:${usuario.usuario}:cxc`;
  const mesCu = () => $("#cu-mes").value || hoy().slice(0, 7);
  const zonaCu = () => $("#cu-zona").value;
  const localesZona = () => locales.filter((l) => l.auditor === zonaCu()).sort((a, b) => a.local.localeCompare(b.local, "es"));
  const nombreMes = (ym) => MES_NOM[+ym.slice(5, 7) - 1];
  const mesesCx = (ym) => [2, 1, 0].map((k) => { const d = new Date(+ym.slice(0, 4), +ym.slice(5, 7) - 1 - k, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; });
  const fechaCortaCu = (iso) => { const d = new Date(iso + "T12:00:00"); return isNaN(d) ? "" : `${d.getDate()}-${d.toLocaleDateString("es-EC", { month: "short" }).replace(".", "")}`; };
  const estatusCu = (d) => d.estatus || (d.fecha ? "OK" : "No Realizado");

  function prepararCumplimiento() {
    if (!$("#cu-zona").options.length) llenar($("#cu-zona"), unicos("auditor"));
    if (usuario.auditor) { $("#cu-zona").value = usuario.auditor; $("#cu-zona").disabled = true; } // el auditor solo ve su zona
    else if (!$("#cu-zona").dataset.elegida) $("#cu-zona").value = (localActual && localActual.auditor) || $("#cu-zona").options[0].value;
    if (!$("#cu-mes").value) $("#cu-mes").value = hoy().slice(0, 7);
    renderCumplimiento(); renderCxc();
  }
  $("#cu-zona").addEventListener("change", () => { $("#cu-zona").dataset.elegida = "1"; renderCumplimiento(); renderCxc(); });
  $("#cu-mes").addEventListener("change", () => { renderCumplimiento(); renderCxc(); });

  function datosCumplimiento() {
    const datos = store.get(claveCu(mesCu()), {});
    const filas = localesZona().map((l) => { const d = datos[l.local] || {}; return { local: l.local, fecha: d.fecha || "", equipo: d.equipo || "", estatus: estatusCu(d) }; });
    const ok = filas.filter((f) => f.estatus === "OK").length;
    return { filas, ok, no: filas.length - ok, pct: filas.length ? Math.round((ok / filas.length) * 100) : 0 };
  }
  function renderCumplimiento() {
    const d = datosCumplimiento();
    $("#cu-resumen").textContent = `${zonaCu()} · ${nombreMes(mesCu())} ${mesCu().slice(0, 4)} · ${d.ok} de ${d.filas.length} locales cumplieron (${d.pct} %)${d.no ? ` · ${d.no} no realizado` : ""}`;
    $("#cu-tabla tbody").innerHTML = d.filas.map((f, i) => `<tr class="${f.estatus === "OK" ? "ok" : "no"}" data-local="${esc(f.local)}">
      <td>${i + 1}</td><td><b>${esc(f.local)}</b></td>
      <td><input type="date" data-cu="fecha" value="${esc(f.fecha)}" aria-label="Fecha de inventario de ${esc(f.local)}"></td>
      <td><input data-cu="equipo" value="${esc(f.equipo)}" placeholder="Administradores que realizaron el inventario" aria-label="Equipo administrativo de ${esc(f.local)}"></td>
      <td><select data-cu="estatus" aria-label="Estatus de ${esc(f.local)}"><option${f.estatus === "OK" ? " selected" : ""}>OK</option><option${f.estatus !== "OK" ? " selected" : ""}>No Realizado</option></select></td></tr>`).join("")
      || `<tr><td colspan="5" class="muted">No hay locales para esta zona.</td></tr>`;
  }
  const guardarCu = (e) => {
    const el = e.target.closest("[data-cu]"); if (!el) return;
    const local = el.closest("tr").dataset.local, datos = store.get(claveCu(mesCu()), {}), d = datos[local] || {};
    d[el.dataset.cu] = el.value;
    datos[local] = d; store.set(claveCu(mesCu()), datos);
    if (e.type === "change") renderCumplimiento(); // al poner la fecha pasa a «OK» si no se eligió estatus
  };
  $("#cu-tabla").addEventListener("input", guardarCu);
  $("#cu-tabla").addEventListener("change", guardarCu);

  function datosCxc() {
    const datos = store.get(claveCx(), {}), ms = mesesCx(mesCu());
    const val = (l, m) => { const v = (datos[l] || {})[m]; return v === "" || v == null ? null : +v; };
    const filas = localesZona().map((l) => ({ local: l.local, v: ms.map((m) => val(l.local, m)) }))
      .sort((a, b) => (b.v[2] ?? -1) - (a.v[2] ?? -1) || a.local.localeCompare(b.local, "es"));
    const totales = ms.map((_, i) => filas.reduce((s, f) => s + (f.v[i] || 0), 0));
    return { ms, filas, totales };
  }
  function renderCxc() {
    const d = datosCxc(), num = (v) => (v == null ? "" : v);
    $("#cx-tabla thead").innerHTML = `<tr><th>#</th><th>Local</th>${d.ms.map((m, i) => `<th class="num${i === 2 ? " ult" : ""}">${nombreMes(m)}</th>`).join("")}</tr>`;
    $("#cx-tabla tbody").innerHTML = d.filas.map((f, i) => `<tr data-local="${esc(f.local)}"><td>${i + 1}</td><td>${esc(f.local)}</td>
      ${d.ms.map((m, j) => `<td class="num${j === 2 ? " ult" : ""}"><input type="number" step="0.01" min="0" data-cx="${m}" value="${num(f.v[j])}" aria-label="${esc(f.local)} ${nombreMes(m)}"></td>`).join("")}</tr>`).join("")
      || `<tr><td colspan="5" class="muted">No hay locales para esta zona.</td></tr>`;
    renderCxcResumen(d);
  }
  // Totales, indicadores y barras (se actualizan mientras se escribe, sin reordenar la tabla)
  function renderCxcResumen(d = datosCxc()) {
    $("#cx-tabla tfoot").innerHTML = `<tr><td></td><td>Total</td>${d.totales.map((t) => `<td class="num">${fmtUSD(t)}</td>`).join("")}</tr>`;
    const [t0, t1, t2] = d.totales, varia = t1 ? (t2 - t1) / t1 : NaN;
    $("#cx-kpis").innerHTML = d.ms.map((m, i) => `<div class="${i === 2 ? "ult" : ""}"><b>${fmtUSD(d.totales[i])}</b><span>Total ${nombreMes(m)}${i === 2 && !isNaN(varia) ? ` · ${varia >= 0 ? "▲" : "▼"} ${fmtPct(Math.abs(varia))} vs ${nombreMes(d.ms[1])}` : ""}</span></div>`).join("");
    // Todos los locales de la zona entran en las barras; los que no tienen saldo aparecen en $0,00
    const barras = barrasCxc(d), max = Math.max(...barras.map((b) => b.v), 0);
    $("#cx-tit").textContent = `Cuentas por cobrar por local · ${nombreMes(d.ms[2])} (total ${fmtUSD(t2)})`;
    $("#cx-barras").innerHTML = barras.length ? barras.map((b) => `<div class="cx-fila${b.v ? "" : " cero"}" title="${esc(b.local)}: ${fmtUSD(b.v)} (${fmtPct(t2 ? b.v / t2 : 0)} del total)">
      <span class="nom">${esc(b.local)}</span><span class="pista"><span class="barra" style="display:block;width:${max ? (b.v / max) * 100 : 0}%"></span></span><span class="val">${fmtUSD(b.v)}</span></div>`).join("")
      : `<p class="muted small">No hay locales para esta zona.</p>`;
    void t0;
  }
  const barrasCxc = (d) => d.filas.map((f) => ({ local: f.local, v: f.v[2] || 0 })).sort((a, b) => b.v - a.v || a.local.localeCompare(b.local, "es"));
  $("#cx-tabla").addEventListener("input", (e) => {
    const el = e.target.closest("[data-cx]"); if (!el) return;
    const local = el.closest("tr").dataset.local, datos = store.get(claveCx(), {});
    datos[local] = Object.assign(datos[local] || {}, { [el.dataset.cx]: el.value }); store.set(claveCx(), datos);
    renderCxcResumen();
  });
  $("#cx-tabla").addEventListener("change", (e) => { if (e.target.closest("[data-cx]")) renderCxc(); }); // al salir del campo se reordena

  // Imagen del cuadro (como la matriz): encabezado Informe / Zona / Mes, tabla y, si hay, barras debajo
  function dibujarCuadroZona({ informe, cols, filas, pie, barras }) {
    const pad = 20, S = ESCALA_COPIA, F = '"Segoe UI", Arial, sans-serif', W = pad * 2 + cols.reduce((s, c) => s + c[1], 0);
    const RF = 30, CAB = 96, altoBar = barras && barras.items.length ? 46 + barras.items.length * 26 : 0;
    const H = pad + CAB + 32 + filas.length * RF + (pie ? RF : 0) + (altoBar ? altoBar + 20 : 0) + pad;
    const cv = document.createElement("canvas"); cv.width = W * S; cv.height = H * S;
    const g = cv.getContext("2d"); g.scale(S, S);
    const txt = (t, x, y, px = 14, peso = 400, color = "#1b1f24", al = "left") => { g.font = `${peso} ${px}px ${F}`; g.fillStyle = color; g.textAlign = al; g.fillText(String(t), x, y); };
    g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
    [["Informe:", informe, 20, 700], ["Zona:", zonaCu(), 15, 600], ["Mes:", `${nombreMes(mesCu())} ${mesCu().slice(0, 4)}`, 15, 600]].forEach(([k, v, px, peso], i) => {
      const y = pad + 24 + i * 28; txt(k, pad + 80, y, 13, 700, "#8a929c", "right"); txt(v, pad + 92, y, px, peso, i ? "#646d78" : "#1b1f24");
      if (i) { g.fillStyle = "#dfe3e8"; g.fillRect(pad + 92, y + 7, 420, 1); }
    });
    let y = pad + CAB, x = pad;
    cols.forEach(([t, w, al, dest]) => {
      if (dest) { const gr = g.createLinearGradient(x, 0, x + w, 0); gr.addColorStop(0, "#ffc629"); gr.addColorStop(1, "#fff1c4"); g.fillStyle = gr; } else g.fillStyle = "#eceef1";
      g.fillRect(x, y, w, 32);
      txt(t, al === "r" ? x + w - 10 : al === "c" ? x + w / 2 : x + 10, y + 21, 13, dest ? 800 : 600, "#1b1f24", al === "r" ? "right" : al === "c" ? "center" : "left");
      x += w;
    });
    y += 32;
    const celda = (c, x0, w, al, yy) => {
      if (c.bg) { g.fillStyle = c.bg; g.fillRect(x0, yy, w, RF); }
      txt(c.t, al === "r" ? x0 + w - 10 : al === "c" ? x0 + w / 2 : x0 + 10, yy + 20, c.px || 14, c.bold ? 700 : 400, c.color || "#1b1f24", al === "r" ? "right" : al === "c" ? "center" : "left");
    };
    filas.forEach((f) => {
      x = pad; f.forEach((c, j) => { celda(c, x, cols[j][1], cols[j][2], y); x += cols[j][1]; });
      y += RF; g.fillStyle = "#e6e8ec"; g.fillRect(pad, y - 1, W - 2 * pad, 1);
    });
    if (pie) { x = pad; pie.forEach((c, j) => { celda({ ...c, bold: true }, x, cols[j][1], cols[j][2], y); x += cols[j][1]; }); g.fillStyle = "#1b1f24"; g.fillRect(pad, y, W - 2 * pad, 1.5); y += RF; }
    if (altoBar) { // barras horizontales: una sola serie, un solo color, valor al final de cada barra
      y += 20; txt(barras.titulo, pad, y + 14, 15, 700);
      y += 30; const max = Math.max(...barras.items.map((b) => b.v)), xb = pad + 230, wb = W - pad - 110 - xb;
      barras.items.forEach((b) => {
        txt(b.nombre, pad, y + 14, 12.5, 400, "#1b1f24");
        const w = Math.max(2, (b.v / max) * wb); g.fillStyle = "#2a62c9";
        g.beginPath(); g.roundRect ? g.roundRect(xb, y + 2, w, 16, [0, 4, 4, 0]) : g.rect(xb, y + 2, w, 16); g.fill();
        txt(fmtUSD(b.v), W - pad, y + 15, 12.5, 700, "#1b1f24", "right");
        y += 26;
      });
    }
    return cv;
  }
  function imagenCumplimiento() {
    const d = datosCumplimiento(), NO = "#fff1c4";
    return dibujarCuadroZona({ informe: "Cumplimiento Inventario Activos",
      cols: [["#", 40, "c"], ["Local", 320, "l", true], ["Fecha inventario", 150, "c"], ["Equipo administrativo", 260, "c"], ["ESTATUS", 150, "c", true]],
      filas: d.filas.map((f, i) => [{ t: i + 1, color: "#8a929c" }, { t: f.local }, { t: fechaCortaCu(f.fecha) }, { t: f.equipo },
        { t: f.estatus, bg: f.estatus === "OK" ? "" : NO, bold: f.estatus !== "OK", color: f.estatus === "OK" ? "#1b1f24" : "#9a6300" }]),
      pie: [{ t: "" }, { t: `Cumplimiento: ${d.ok} de ${d.filas.length} locales` }, { t: "" }, { t: "" }, { t: `${d.pct} %` }] });
  }
  // Imagen del dashboard de cuentas por cobrar (lo que se pega en el informe): encabezado, total de los
  // tres meses y barras por local del último mes. Una sola serie y un solo color; el valor va al final de cada barra.
  function imagenCxc() {
    const d = datosCxc(), barras = barrasCxc(d), [, t1, t2] = d.totales, varia = t1 ? (t2 - t1) / t1 : NaN;
    const S = ESCALA_COPIA, F = '"Segoe UI", Arial, sans-serif', W = 900, pad = 24, FILA = 30;
    const H = pad + 96 + 112 + 44 + barras.length * FILA + pad;
    const cv = document.createElement("canvas"); cv.width = W * S; cv.height = H * S;
    const g = cv.getContext("2d"); g.scale(S, S);
    const txt = (t, x, y, px = 14, peso = 400, color = "#1b1f24", al = "left") => { g.font = `${peso} ${px}px ${F}`; g.fillStyle = color; g.textAlign = al; g.fillText(String(t), x, y); };
    const caja = (x, y, w, h, r, borde, ancho = 1) => { g.beginPath(); g.roundRect ? g.roundRect(x, y, w, h, r) : g.rect(x, y, w, h); g.fillStyle = "#fff"; g.fill(); g.strokeStyle = borde; g.lineWidth = ancho; g.stroke(); };
    g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
    // Encabezado como la matriz
    [["Informe:", "Análisis de Cuentas por Cobrar", 20, 700], ["Zona:", zonaCu(), 15, 600], ["Mes:", `${nombreMes(mesCu())} ${mesCu().slice(0, 4)}`, 15, 600]].forEach(([k, v, px, peso], i) => {
      const y = pad + 22 + i * 28; txt(k, pad + 70, y, 13, 700, "#8a929c", "right"); txt(v, pad + 82, y, px, peso, i ? "#646d78" : "#1b1f24");
    });
    // Totales de los tres meses (el último resaltado, con su variación)
    let y = pad + 96;
    const wk = (W - 2 * pad - 2 * 18) / 3;
    d.ms.forEach((m, i) => {
      const x = pad + i * (wk + 18), ult = i === 2;
      caja(x, y, wk, 92, 12, ult ? "#2a62c9" : "#dfe3e8", ult ? 1.5 : 1);
      txt(fmtUSD(d.totales[i]), x + 20, y + 42, 28, 700);
      txt(`Total ${nombreMes(m)}${ult && !isNaN(varia) ? ` · ${varia >= 0 ? "▲" : "▼"} ${fmtPct(Math.abs(varia))} vs ${nombreMes(d.ms[1])}` : ""}`, x + 20, y + 72, 14, 400, "#646d78");
    });
    // Barras por local
    y += 112;
    txt(`Cuentas por cobrar por local · ${nombreMes(d.ms[2])} (total ${fmtUSD(t2)})`, pad, y + 18, 16, 700, "#4a525c");
    y += 44;
    const max = Math.max(...barras.map((b) => b.v), 0), xb = pad + 300, wb = W - pad - 100 - xb;
    barras.forEach((b) => {
      g.font = `400 15px ${F}`; let nom = b.local; // nombres largos se recortan para no tocar la barra
      while (nom.length > 4 && g.measureText(nom).width > xb - pad - 14) nom = nom.slice(0, -2) + "…";
      txt(nom, pad, y + 15, 15, 400);
      const w = max && b.v ? Math.max(3, (b.v / max) * wb) : 3; // en cero queda una marca mínima
      g.fillStyle = b.v ? "#2a62c9" : "#c9ced6";
      g.beginPath(); g.roundRect ? g.roundRect(xb, y, w, 20, [0, 5, 5, 0]) : g.rect(xb, y, w, 20); g.fill();
      txt(fmtUSD(b.v), W - pad, y + 15, 15, 700, b.v ? "#1b1f24" : "#8a929c", "right");
      y += FILA;
    });
    return cv;
  }
  async function copiarCanvas(cv, nombre) {
    const blob = await new Promise((ok) => cv.toBlob(ok, "image/png"));
    try { await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]); toast("Copiado como imagen: pégalo con Ctrl+V"); }
    catch { descargar(blob, `${nombre}.png`.replace(/[\\/:*?"<>|]/g, "_")); toast("Tu navegador no permitió copiar: se descargó la imagen"); }
  }
  $("#cu-copiar").addEventListener("click", () => copiarCanvas(imagenCumplimiento(), `Cumplimiento activos ${zonaCu()} ${mesCu()}`));
  $("#cx-copiar").addEventListener("click", () => copiarCanvas(imagenCxc(), `Cuentas por cobrar ${zonaCu()} ${mesCu()}`));

  // ---------- Baja de activos ----------
  // st.bajas = [{ id, activoId, codigo, nombre, marca, modelo, serie, valor, motivo, obs, autoriza, trat, gerente, jefe,
  //   estado: "Pendiente" | "Aprobado", fecha, fechaAprob }]. Quién autoriza, certifica y el jefe se recuerdan por local (st.bajaDatos).
  // Al aprobarse, la baja sale de «Activos por dar de baja» y queda en el historial de aprobadas.
  let bActivo = null, bUrls = [], bLocal = null;
  const bajasLocal = () => (localActual ? estadoLocal(localActual.local).bajas || [] : []);
  const bajaPendiente = (b) => b.estado !== "Aprobado";
  function renderBajas() {
    if (!libros && !librosIntentado) { librosIntentado = true; cargarLibros(); } // para proponer el valor en libros
    if (!localActual) { $("#b-local").textContent = "Selecciona un local."; $("#b-lista").innerHTML = ""; return; }
    const st = estadoLocal(localActual.local), d = st.bajaDatos || {};
    if (bLocal !== localActual.local) { // al cambiar de local se limpian los datos del formulario
      bLocal = localActual.local; bActivo = null; $("#b-buscar").value = ""; $("#b-valor").value = ""; $("#b-obs").value = "";
      $("#b-autoriza").value = d.autoriza || ""; $("#b-gerente").value = d.gerente || ""; $("#b-jefe").value = d.jefe || localActual.jefe || "";
      if (d.trat) $("#b-trat").value = d.trat;
    }
    $("#b-local").textContent = `${localActual.local} · Elige el activo que se va a dar de baja.`;
    buscarBaja(); renderBajaFicha();
    const bajas = st.bajas || [], pend = bajas.filter(bajaPendiente), apr = bajas.filter((b) => !bajaPendiente(b));
    const item = (b) => `<div class="b-item ${bajaPendiente(b) ? "pendiente" : "aprobado"}" data-baja="${b.id}">
      <div><b>${bajaPendiente(b) ? "Activo por dar de baja" : "Dado de baja"}: ${esc(b.nombre)}</b>
        <small>${esc([b.codigo, [b.marca, b.modelo].filter(Boolean).join(" "), b.serie && `S: ${b.serie}`].filter(Boolean).join(" · "))}</small>
        <small>Valor en libros: ${b.valor === "" || b.valor == null ? "—" : fmtUSD(+b.valor)} · ${esc(b.motivo)} · Certifica: ${esc(b.gerente || "—")} · Jefe de área: ${esc(b.jefe || "—")}${!bajaPendiente(b) && b.fechaAprob ? ` · Aprobado el ${esc(b.fechaAprob)}` : ""}</small></div>
      <select data-b-estado aria-label="Estado de la baja de ${esc(b.nombre)}"><option${bajaPendiente(b) ? " selected" : ""}>Pendiente</option><option${!bajaPendiente(b) ? " selected" : ""}>Aprobado</option></select>
      <div class="b-acc"><button type="button" class="ghost sm" data-b-editar title="Editar datos">✏️</button><button type="button" class="ghost sm" data-b-pdf>📄 Ficha</button><button type="button" class="ghost sm" data-b-quitar title="Quitar de la lista" aria-label="Quitar">✕</button></div></div>`;
    $("#b-lista").innerHTML = pend.length ? pend.map(item).join("") : `<p class="b-vacio">No hay activos pendientes de baja.</p>`;
    $("#b-aprobados-box").hidden = !apr.length;
    $("#b-aprobados-tit").textContent = `Bajas aprobadas (${apr.length})`;
    $("#b-aprobados").innerHTML = apr.map(item).join("");
  }
  function buscarBaja() {
    if (bActivo) { $("#b-resultados").innerHTML = `<li><button type="button" class="ghost" data-b-cambiar>Elegir otro activo</button></li>`; return; }
    const q = $("#b-buscar").value.trim().toLowerCase();
    if (!q || !localActual) { $("#b-resultados").innerHTML = ""; return; }
    const lista = activosLocal().activos.filter((a) => (enUso(a) || a.origen !== "SIR")
      && [a.codigo, a.erp, a.nombre, a.marca, a.modelo, a.serie].join(" ").toLowerCase().includes(q)).slice(0, 8);
    $("#b-resultados").innerHTML = lista.length ? lista.map((a) => `<li><button type="button" class="ghost" data-b-activo="${a.id}">
      <b>${esc(a.nombre)}</b><small>${esc([a.codigo, [a.marca, a.modelo].filter(Boolean).join(" "), a.serie && `S: ${a.serie}`].filter(Boolean).join(" · "))}</small></button></li>`).join("")
      : `<li class="muted small">Ningún activo coincide.</li>`;
  }
  async function renderBajaFicha() {
    bUrls.forEach(URL.revokeObjectURL); bUrls = [];
    const hay = !!bActivo;
    $("#b-ficha").hidden = !hay; $("#b-fotos-box").hidden = !hay; $("#b-ficha-pdf").disabled = !hay; $("#b-agregar").disabled = !hay;
    if (!hay) return;
    const a = activosLocal().activos.find((x) => x.id === bActivo.id) || bActivo, b = bajasLocal().find((x) => x.activoId === a.id);
    $("#b-agregar").textContent = b ? "Actualizar datos de la baja" : "Agregar a activos por dar de baja";
    $("#b-ficha").innerHTML = `<div class="bf-cab"><div><b>${esc(a.nombre)}</b><br><small class="muted">${esc(localActual.local)}</small></div>
      <div class="bf-cod"><small class="muted">CÓDIGO</small><br>${esc(a.codigo || "—")}</div></div>
      <table>${[["Marca", a.marca], ["Modelo", a.modelo], ["Serie", a.serie], ["Estado SIR", a.estado],
        ["Datos SIR", [a.erp && `ERP ${a.erp}`, a.fecha && `Creado ${a.fecha}`].filter(Boolean).join(" · ")],
        ["Libros (app)", a.app && a.app.libros != null ? `${fmtUSD(a.app.libros)} (${a.app.librosMes})` : ""]]
        .filter(([, v]) => v).map(([k, v]) => `<tr><th>${k}</th><td>${esc(v)}</td></tr>`).join("")}</table>`;
    const fotos = (await Promise.all((a.fotos || []).map((id) => fotosDB.get(id).catch(() => null)))).filter(Boolean);
    $("#b-miniaturas").innerHTML = fotos.map((f) => { const u = URL.createObjectURL(f.blob); bUrls.push(u); return `<div class="foto"><img src="${u}" alt="Foto de ${esc(a.nombre)}"></div>`; }).join("");
  }
  function elegirActivoBaja(id) {
    bActivo = activosLocal().activos.find((x) => x.id === id) || null;
    const b = bActivo && bajasLocal().find((x) => x.activoId === id);
    // Si ya está en la lista se cargan sus datos; si no, se propone el valor en libros que trae la app
    const lb = bActivo && libroDe(bActivo, localActual.local); // base contable del admin; si no, lo que trae la app
    $("#b-valor").value = b ? b.valor ?? "" : lb ? valorLibro(lb) ?? "" : bActivo && bActivo.app && bActivo.app.libros != null ? bActivo.app.libros : "";
    $("#b-obs").value = b ? b.obs || "" : "";
    if (b) { $("#b-motivo").value = b.motivo; $("#b-autoriza").value = b.autoriza; $("#b-gerente").value = b.gerente; $("#b-jefe").value = b.jefe; $("#b-trat").value = b.trat || "Estimado Señor"; }
    renderBajas();
  }
  $("#b-buscar").addEventListener("input", buscarBaja);
  $("#b-resultados").addEventListener("click", (e) => {
    if (e.target.closest("[data-b-cambiar]")) { bActivo = null; renderBajas(); $("#b-buscar").focus(); return; }
    const b = e.target.closest("[data-b-activo]"); if (b) elegirActivoBaja(b.dataset.bActivo);
  });
  $("#b-miniaturas").addEventListener("click", (e) => { const img = e.target.closest("img"); if (!img) return; $("#f-grande").src = img.src; $("#dlg-ver").showModal(); });
  async function agregarFotosBaja(files) {
    if (!bActivo) return;
    const st = activosLocal(), a = st.activos.find((x) => x.id === bActivo.id);
    for (const file of files) {
      if (!file.type.startsWith("image/")) continue;
      try { const blob = await reducirImagen(file), id = nuevoId(); await fotosDB.put({ id, blob, local: localActual.local, activo: a.id, fecha: new Date().toISOString() }); a.fotos = (a.fotos || []).concat(id); }
      catch { toast(`No se pudo leer ${file.name}`); }
    }
    guardar({ activos: st.activos }); bActivo = a; renderBajaFicha(); renderActivos();
  }
  $("#b-agregar-foto").addEventListener("click", () => $("#b-input").click());
  $("#b-tomar").addEventListener("click", () => $("#b-camara").click());
  ["#b-input", "#b-camara"].forEach((s) => $(s).addEventListener("change", (e) => { agregarFotosBaja([...e.target.files]); e.target.value = ""; }));
  $("#b-agregar").addEventListener("click", () => {
    if (!bActivo || !localActual) return;
    const st = estadoLocal(localActual.local), bajas = st.bajas || [], a = activosLocal().activos.find((x) => x.id === bActivo.id);
    const datos = { autoriza: $("#b-autoriza").value.trim(), gerente: $("#b-gerente").value.trim(), jefe: $("#b-jefe").value.trim(), trat: $("#b-trat").value };
    if (!datos.autoriza || !datos.gerente || !datos.jefe) return toast("Completa quién autoriza, el subgerente que certifica y el jefe de área");
    let b = bajas.find((x) => x.activoId === a.id);
    const nuevo = !b;
    if (!b) { b = { id: nuevoId(), activoId: a.id, estado: "Pendiente", fecha: hoy() }; bajas.push(b); }
    Object.assign(b, { codigo: a.codigo, nombre: a.nombre, marca: a.marca, modelo: a.modelo, serie: a.serie, valor: $("#b-valor").value, motivo: $("#b-motivo").value, obs: $("#b-obs").value.trim() }, datos);
    guardar({ bajas, bajaDatos: datos });
    toast(nuevo ? `${a.nombre}: agregado a activos por dar de baja` : `${a.nombre}: datos de la baja actualizados`);
    bActivo = null; $("#b-buscar").value = ""; $("#b-valor").value = ""; $("#b-obs").value = "";
    renderBajas(); renderActivos();
  });
  const editarBaja = (id, fn) => { const st = estadoLocal(localActual.local), bajas = st.bajas || [], b = bajas.find((x) => x.id === id); if (!b) return; fn(b, bajas); guardar({ bajas }); renderBajas(); renderActivos(); };
  ["#b-lista", "#b-aprobados"].forEach((s) => {
    $(s).addEventListener("change", (e) => {
      const sel = e.target.closest("[data-b-estado]"); if (!sel) return;
      const id = sel.closest("[data-baja]").dataset.baja;
      editarBaja(id, (b) => { b.estado = sel.value; b.fechaAprob = sel.value === "Aprobado" ? hoy() : ""; });
      toast(sel.value === "Aprobado" ? "Baja aprobada: ya no aparece en activos por dar de baja" : "La baja volvió a pendiente");
    });
    $(s).addEventListener("click", (e) => {
      const it = e.target.closest("[data-baja]"); if (!it) return;
      const b = bajasLocal().find((x) => x.id === it.dataset.baja); if (!b) return;
      if (e.target.closest("[data-b-pdf]")) return abrirFichasDe([b.activoId]);
      if (e.target.closest("[data-b-editar]")) { elegirActivoBaja(b.activoId); $("#b-buscar").scrollIntoView({ behavior: "smooth", block: "center" }); return; }
      if (e.target.closest("[data-b-quitar]") && confirm(`¿Quitar ${b.nombre} de la lista de bajas?`)) editarBaja(b.id, (x, bajas) => bajas.splice(bajas.indexOf(x), 1));
    });
  });
  $("#b-ficha-pdf").addEventListener("click", () => bActivo && abrirFichasDe([bActivo.id]));
  $("#b-pdf-todas").addEventListener("click", () => {
    const ids = bajasLocal().filter(bajaPendiente).map((b) => b.activoId);
    if (!ids.length) return toast("No hay activos pendientes de baja");
    abrirFichasDe(ids);
  });

  // ---------- Asignar fotos sueltas a activos ----------
  // Las fotos de la cámara se llaman IMG_AAAAMMDD_HHMMSS: se ordenan por hora y se agrupan las tomadas seguidas
  // (en los respaldos reales, las de un mismo equipo están a ≤ 15 s y al pasar al siguiente equipo hay 20 s o más).
  // Flujo: grupo preseleccionado → escribir código → Enter → siguiente grupo.
  const gapGrupo = () => (+$("#as-gap").value || 15) * 1000;
  let asg = null; // { fotos: [{ id, nombre, url, ts }], sel: Set(id), ultimo: índice, urls: [] }
  const tsDeNombre = (n) => { const m = /(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})/.exec(n || ""); return m ? new Date(+m[1], m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime() : null; };
  function gruposAsg() {
    const g = [];
    asg.fotos.forEach((f, i) => {
      const prev = asg.fotos[i - 1];
      if (!prev || f.ts == null || prev.ts == null || f.ts - prev.ts > gapGrupo()) g.push([]);
      g[g.length - 1].push(f);
    });
    return g;
  }
  async function abrirAsignar() {
    if (!localActual) return toast("Selecciona un local primero");
    const ids = activosLocal().fotosSueltas || [];
    if (!ids.length) return toast("Este local no tiene fotos sueltas");
    toast("Cargando fotos…");
    const recs = (await Promise.all(ids.map((id) => fotosDB.get(id).catch(() => null)))).filter(Boolean);
    asg = { fotos: [], sel: new Set(), ultimo: -1, urls: [] };
    asg.fotos = recs.map((r) => { const url = URL.createObjectURL(r.blob); asg.urls.push(url); return { id: r.id, nombre: r.nombre || r.id, url, ts: tsDeNombre(r.nombre) }; })
      .sort((a, b) => (a.ts ?? 0) - (b.ts ?? 0) || a.nombre.localeCompare(b.nombre));
    const primero = gruposAsg()[0] || [];
    primero.forEach((f) => asg.sel.add(f.id));
    renderAsignar();
    $("#dlg-asignar").showModal();
    $("#as-buscar").value = ""; buscarAsg(); $("#as-buscar").focus();
  }
  function renderAsignar() {
    const g = gruposAsg(), hora = (f) => (f.ts ? new Date(f.ts).toLocaleTimeString("es-EC", { hour: "2-digit", minute: "2-digit", second: "2-digit" }) : f.nombre);
    const dia = (f) => (f.ts ? new Date(f.ts).toLocaleDateString("es-EC", { day: "numeric", month: "short" }) : "");
    $("#as-sub").textContent = `${localActual.local} · ${asg.fotos.length} fotos sin asignar en ${g.length} grupos`;
    $("#as-fotos").innerHTML = asg.fotos.length ? g.map((grupo, gi) => `
      <div class="as-grupo"><span>Grupo ${gi + 1} · ${esc(dia(grupo[0]))} ${esc(hora(grupo[0]))} · ${grupo.length} foto${grupo.length > 1 ? "s" : ""}</span>
        <button type="button" class="ghost sm" data-grupo="${gi}">Seleccionar grupo</button></div>` +
      grupo.map((f) => `<button type="button" class="as-foto${asg.sel.has(f.id) ? " sel" : ""}" data-id="${f.id}" title="${esc(f.nombre)}" aria-pressed="${asg.sel.has(f.id)}">
        <span class="as-zoom" data-zoom title="Ampliar" aria-hidden="true">🔍</span><img src="${f.url}" alt="Foto ${esc(f.nombre)}" loading="lazy"><span>${esc(hora(f))}</span></button>`).join("")).join("")
      : `<p class="muted">¡Listo! No quedan fotos sueltas en este local.</p>`;
    $("#as-nsel").textContent = `${asg.sel.size} seleccionada${asg.sel.size === 1 ? "" : "s"}`;
  }
  function buscarAsg() {
    const q = $("#as-buscar").value.trim().toLowerCase(), sinFotos = $("#as-sinfotos").checked;
    const lista = activosLocal().activos.filter((a) => (enUso(a) || a.origen === "app")
      && (!sinFotos || !a.fotos.length)
      && (!q || [a.codigo, a.erp, a.nombre, a.marca, a.modelo, a.serie].join(" ").toLowerCase().includes(q)))
      .sort((x, y) => (q && x.codigo.toLowerCase().startsWith(q) ? -1 : 0) - (q && y.codigo.toLowerCase().startsWith(q) ? -1 : 0) || x.nombre.localeCompare(y.nombre, "es"))
      .slice(0, 12);
    $("#as-resultados").innerHTML = lista.length ? lista.map((a) => `<li><button type="button" class="ghost" data-activo="${a.id}">
      <b>${esc(a.nombre)}</b><small>${esc([a.codigo, [a.marca, a.modelo].filter(Boolean).join(" "), a.serie && `S: ${a.serie}`].filter(Boolean).join(" · "))}${a.fotos.length ? ` · 📷 ${a.fotos.length}` : ""}</small></button></li>`).join("")
      : `<li class="muted small">Ningún activo coincide.</li>`;
  }
  function asignarA(idActivo) {
    if (!asg.sel.size) return toast("Selecciona al menos una foto");
    const st = activosLocal(), a = st.activos.find((x) => x.id === idActivo); if (!a) return;
    const ids = asg.fotos.filter((f) => asg.sel.has(f.id)).map((f) => f.id), quitadas = asg.fotos.filter((f) => asg.sel.has(f.id));
    const idxPrimera = asg.fotos.indexOf(quitadas[0]);
    a.fotos = a.fotos.concat(ids);
    const sueltas = (st.fotosSueltas || []).filter((id) => !ids.includes(id));
    guardar({ activos: st.activos, fotosSueltas: sueltas });
    asg.fotos = asg.fotos.filter((f) => !asg.sel.has(f.id));
    // Se preselecciona el grupo que sigue a las fotos asignadas
    asg.sel.clear(); asg.ultimo = -1;
    const siguiente = asg.fotos[Math.min(idxPrimera, asg.fotos.length - 1)];
    if (siguiente) (gruposAsg().find((g) => g.includes(siguiente)) || []).forEach((f) => asg.sel.add(f.id));
    renderAsignar(); renderActivos();
    $("#as-buscar").value = ""; buscarAsg(); $("#as-buscar").focus();
    const sel = $(".as-foto.sel"); if (sel) sel.scrollIntoView({ block: "nearest" });
    toast(`${ids.length} foto${ids.length > 1 ? "s" : ""} asignada${ids.length > 1 ? "s" : ""} a ${a.nombre}`, "Deshacer", () => {
      const s2 = activosLocal(), a2 = s2.activos.find((x) => x.id === idActivo);
      if (a2) a2.fotos = a2.fotos.filter((id) => !ids.includes(id));
      guardar({ activos: s2.activos, fotosSueltas: (s2.fotosSueltas || []).concat(ids) });
      if (asg && $("#dlg-asignar").open) {
        asg.fotos = asg.fotos.concat(quitadas).sort((x, y) => (x.ts ?? 0) - (y.ts ?? 0) || x.nombre.localeCompare(y.nombre));
        asg.sel = new Set(ids); renderAsignar();
      }
      renderActivos();
    });
  }
  $("#as-fotos").addEventListener("click", (e) => {
    const z = e.target.closest("[data-zoom]");
    if (z) { $("#f-grande").src = z.parentElement.querySelector("img").src; $("#dlg-ver").showModal(); return; }
    const g = e.target.closest("[data-grupo]");
    if (g) { asg.sel = new Set(gruposAsg()[+g.dataset.grupo].map((f) => f.id)); renderAsignar(); return; }
    const b = e.target.closest(".as-foto"); if (!b) return;
    const i = asg.fotos.findIndex((f) => f.id === b.dataset.id);
    if (e.shiftKey && asg.ultimo >= 0) {
      const [a, z] = [Math.min(i, asg.ultimo), Math.max(i, asg.ultimo)];
      asg.fotos.slice(a, z + 1).forEach((f) => asg.sel.add(f.id));
    } else if (asg.sel.has(b.dataset.id)) asg.sel.delete(b.dataset.id); else asg.sel.add(b.dataset.id);
    asg.ultimo = i; renderAsignar();
  });
  $("#as-fotos").addEventListener("dblclick", (e) => {
    const b = e.target.closest(".as-foto"); if (!b) return;
    $("#f-grande").src = b.querySelector("img").src; $("#dlg-ver").showModal();
  });
  $("#as-buscar").addEventListener("input", buscarAsg);
  $("#as-sinfotos").addEventListener("change", buscarAsg);
  $("#as-gap").addEventListener("change", () => { // al cambiar el corte se vuelve a seleccionar el grupo de la primera foto seleccionada
    const ref = asg.fotos.find((f) => asg.sel.has(f.id)) || asg.fotos[0];
    asg.sel = new Set(((ref && gruposAsg().find((g) => g.includes(ref))) || []).map((f) => f.id));
    renderAsignar(); $("#as-buscar").focus();
  });
  $("#as-buscar").addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    const primero = $("#as-resultados [data-activo]"); if (primero) asignarA(primero.dataset.activo);
  });
  $("#as-resultados").addEventListener("click", (e) => { const b = e.target.closest("[data-activo]"); if (b) asignarA(b.dataset.activo); });
  $("#as-limpiar").addEventListener("click", () => { asg.sel.clear(); renderAsignar(); });
  $("#as-eliminar").addEventListener("click", async () => {
    if (!asg.sel.size || !confirm(`¿Eliminar ${asg.sel.size} foto(s) sueltas? No se puede deshacer.`)) return;
    const ids = [...asg.sel];
    await Promise.all(ids.map((id) => fotosDB.del(id).catch(() => {})));
    guardar({ fotosSueltas: (activosLocal().fotosSueltas || []).filter((id) => !ids.includes(id)) });
    asg.fotos = asg.fotos.filter((f) => !asg.sel.has(f.id)); asg.sel.clear(); renderAsignar(); renderActivos();
  });
  $("#as-cerrar").addEventListener("click", () => $("#dlg-asignar").close());
  $("#dlg-asignar").addEventListener("close", () => { if (asg) asg.urls.forEach(URL.revokeObjectURL); asg = null; });

  // ---------- Fotos de activos (IndexedDB: las imágenes no caben en localStorage) ----------
  const fotosDB = (() => {
    let abierta;
    const abrir = () => abierta || (abierta = new Promise((ok, mal) => {
      const r = indexedDB.open("matriz-fotos", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("fotos", { keyPath: "id" });
      r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error);
    }));
    const tx = async (modo, fn) => {
      const db = await abrir();
      return new Promise((ok, mal) => { const t = db.transaction("fotos", modo), req = fn(t.objectStore("fotos")); t.oncomplete = () => ok(req.result); t.onerror = () => mal(t.error); });
    };
    return { put: (f) => tx("readwrite", (s) => s.put(f)), get: (id) => tx("readonly", (s) => s.get(id)), del: (id) => tx("readwrite", (s) => s.delete(id)) };
  })();
  async function reducirImagen(file, max = 1600) {
    const bmp = await createImageBitmap(file);
    const k = Math.min(1, max / Math.max(bmp.width, bmp.height)), c = document.createElement("canvas");
    c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    return new Promise((ok) => c.toBlob(ok, "image/jpeg", 0.8));
  }
  let activoFotos = null, urlsFotos = [];
  const liberarUrls = () => { urlsFotos.forEach(URL.revokeObjectURL); urlsFotos = []; };
  // Destino de las fotos: un activo, o las «fotos sueltas» del local (vienen del respaldo de la app sin activo asignado)
  function destinoFotos() {
    const st = activosLocal();
    if (activoFotos === "__sueltas") return { nombre: "Fotos sueltas del local", sub: "Fotos del respaldo de la app que no están asignadas a un activo", lista: st.fotosSueltas || [], fijar: (l) => guardar({ fotosSueltas: l }) };
    const a = st.activos.find((x) => x.id === activoFotos); if (!a) return null;
    return { nombre: a.nombre, sub: [a.codigo, a.marca, a.modelo, a.serie && `S: ${a.serie}`, a.novedad].filter(Boolean).join(" · "), lista: a.fotos,
      fijar: (l) => { a.fotos = l; guardar({ activos: st.activos }); } };
  }
  async function renderFotos() {
    const a = destinoFotos(); if (!a) return;
    $("#f-titulo").textContent = `Fotos · ${a.nombre}`;
    $("#f-sub").textContent = a.sub;
    liberarUrls();
    const fotos = (await Promise.all(a.lista.map((id) => fotosDB.get(id).catch(() => null)))).filter(Boolean);
    $("#f-grid").innerHTML = fotos.length ? fotos.map((f) => { const u = URL.createObjectURL(f.blob); urlsFotos.push(u);
      return `<div class="foto" data-foto="${f.id}"><img src="${u}" alt="Foto de ${esc(a.nombre)}"><button type="button" data-borrar-foto title="Eliminar foto" aria-label="Eliminar foto">✕</button></div>`; }).join("")
      : `<p class="muted">Sin fotos todavía. Agrega imágenes desde tu equipo o toma una foto con la cámara.</p>`;
  }
  function abrirFotos(id) { activoFotos = id; renderFotos(); $("#dlg-fotos").showModal(); }
  async function agregarFotos(files) {
    if (!files.length || !activoFotos) return;
    toast("Guardando fotos…");
    const ids = [];
    for (const file of files) {
      if (!file.type.startsWith("image/")) continue;
      try { const blob = await reducirImagen(file), id = nuevoId(); await fotosDB.put({ id, blob, local: localActual.local, activo: activoFotos, fecha: new Date().toISOString() }); ids.push(id); }
      catch { toast(`No se pudo leer ${file.name}`); }
    }
    const dst = destinoFotos(); dst.fijar(dst.lista.concat(ids));
    renderFotos(); renderActivos(); if (ids.length) toast(`${ids.length} foto(s) agregada(s)`);
  }
  $("#f-agregar").addEventListener("click", () => $("#f-input").click());
  $("#f-tomar").addEventListener("click", () => $("#f-camara").click());
  ["#f-input", "#f-camara"].forEach((s) => $(s).addEventListener("change", (e) => { agregarFotos([...e.target.files]); e.target.value = ""; }));
  $("#f-grid").addEventListener("click", async (e) => {
    const box = e.target.closest("[data-foto]"); if (!box) return;
    if (e.target.closest("[data-borrar-foto]")) {
      if (!confirm("¿Eliminar esta foto?")) return;
      await fotosDB.del(box.dataset.foto).catch(() => {});
      const dst = destinoFotos(); dst.fijar(dst.lista.filter((f) => f !== box.dataset.foto)); renderFotos(); renderActivos(); return;
    }
    $("#f-grande").src = box.querySelector("img").src; $("#dlg-ver").showModal();
  });
  $("#f-cerrar").addEventListener("click", () => $("#dlg-fotos").close());
  $("#dlg-fotos").addEventListener("close", () => { liberarUrls(); activoFotos = null; });
  $("#f-cerrar-ver").addEventListener("click", () => $("#dlg-ver").close());
  // La foto ampliada se cierra con clic sobre ella o fuera de ella (además de Esc y «Cerrar»)
  $("#dlg-ver").addEventListener("click", (e) => { if (e.target.id === "dlg-ver") $("#dlg-ver").close(); });

  // ---------- Zoom de la foto ampliada (sirve para todas las fotos de la app) ----------
  // Rueda del mouse = acercar/alejar hacia donde apunta el cursor; arrastrar = mover; doble clic = volver a ajustar.
  const zoom = { s: 1, x: 0, y: 0, arrastre: null };
  const Z_MAX = 8;
  function aplicarZoom() {
    if (zoom.s <= 1) { zoom.s = 1; zoom.x = 0; zoom.y = 0; }
    const v = $("#f-visor"), img = $("#f-grande");
    // que la imagen no se salga del recuadro al mover
    zoom.x = Math.min(0, Math.max(v.clientWidth - img.clientWidth * zoom.s, zoom.x));
    zoom.y = Math.min(0, Math.max(v.clientHeight - img.clientHeight * zoom.s, zoom.y));
    img.style.transform = `translate(${zoom.x}px, ${zoom.y}px) scale(${zoom.s})`;
    $("#z-nivel").textContent = `${Math.round(zoom.s * 100)}%`;
    v.classList.toggle("zoom", zoom.s > 1);
  }
  function zoomEn(factor, px, py) {
    const v = $("#f-visor");
    if (px == null) { px = v.clientWidth / 2; py = v.clientHeight / 2; }
    const ns = Math.min(Z_MAX, Math.max(1, zoom.s * factor));
    zoom.x = px - (px - zoom.x) * (ns / zoom.s);
    zoom.y = py - (py - zoom.y) * (ns / zoom.s);
    zoom.s = ns; aplicarZoom();
  }
  const resetZoom = () => { zoom.s = 1; zoom.x = 0; zoom.y = 0; aplicarZoom(); };
  $("#f-grande").addEventListener("load", resetZoom);
  $("#dlg-ver").addEventListener("close", resetZoom);
  $("#f-visor").addEventListener("wheel", (e) => {
    e.preventDefault();
    const r = $("#f-visor").getBoundingClientRect();
    zoomEn(e.deltaY < 0 ? 1.2 : 1 / 1.2, e.clientX - r.left, e.clientY - r.top);
  }, { passive: false });
  $("#f-visor").addEventListener("pointerdown", (e) => {
    if (zoom.s <= 1) return;
    zoom.arrastre = { x: e.clientX - zoom.x, y: e.clientY - zoom.y };
    $("#f-visor").setPointerCapture(e.pointerId); $("#f-visor").classList.add("moviendo");
  });
  $("#f-visor").addEventListener("pointermove", (e) => {
    if (!zoom.arrastre) return;
    zoom.x = e.clientX - zoom.arrastre.x; zoom.y = e.clientY - zoom.arrastre.y; aplicarZoom();
  });
  ["pointerup", "pointercancel"].forEach((t) => $("#f-visor").addEventListener(t, () => { zoom.arrastre = null; $("#f-visor").classList.remove("moviendo"); }));
  $("#f-visor").addEventListener("dblclick", (e) => {
    if (zoom.s > 1) return resetZoom();
    const r = $("#f-visor").getBoundingClientRect(); zoomEn(2.5, e.clientX - r.left, e.clientY - r.top);
  });
  $("#z-mas").addEventListener("click", () => zoomEn(1.5));
  $("#z-menos").addEventListener("click", () => zoomEn(1 / 1.5));
  $("#z-reset").addEventListener("click", resetZoom);
  // Arrastrar imágenes sobre la ventana de fotos
  $("#dlg-fotos").addEventListener("dragover", (e) => e.preventDefault());
  $("#dlg-fotos").addEventListener("drop", (e) => { e.preventDefault(); agregarFotos([...e.dataTransfer.files]); });

  // ---------- Producto (Inventario Total / Arqueo de Producto) ----------
  function fmtPct(v) { return isNaN(v) ? "—" : (v * 100).toLocaleString("es-EC", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%"; }
  function fmtUSD(v) { return (v < 0 ? "-$" : "$") + Math.abs(v || 0).toLocaleString("es-EC", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  const UNIDADES = M.unidades || [];
  const producto = () => Object.assign(productoVacio(), localActual ? estadoLocal(localActual.local).producto : {});
  const guardarProducto = (p) => guardar({ producto: p });
  const sumaLista = (lista) => lista.reduce((s, x) => s + (parseFloat(x.total) || 0), 0);
  const filaLlena = (x) => x.articulo || parseFloat(x.total) || parseFloat(x.cantidad);

  function renderProducto() {
    const p = producto();
    $("#p-tipo").value = p.tipo;
    $("#p-local").value = localActual ? localActual.local : "";
    $("#p-venta").value = p.venta;
    $("#p-fecha").value = p.fecha || $("#sel-fecha").value || hoy();
    const a = p.archivo;
    $("#p-archivo-info").textContent = a
      ? `Cargado «${a.nombre}» · ${a.local} · ${a.desde} al ${a.hasta} · ${a.articulos} artículos: ${a.nFal} faltantes (${fmtUSD(a.totFal)}) y ${a.nSob} sobrantes (${fmtUSD(a.totSob)}). Se muestra el top ${a.top}; puedes editar cantidades y valores.`
      : "Sube el archivo .xls exportado del SIR (o arrástralo aquí) para llenar automáticamente el TOP de sobrantes y faltantes.";
    if (a) $("#p-top").value = a.top;
    const co = p.costos || {};
    $$("[data-costo]").forEach((i) => (i.value = co[i.dataset.costo] ?? ""));
    renderDifCostos(co);
    $$(".ef-items").forEach((i) => (i.value = p.items));
    $$(".ef-error").forEach((i) => (i.value = p.error));
    ["sobrantes", "faltantes"].forEach((k) => {
      const filas = p[k].slice();
      while (filas.length < 3) filas.push({ articulo: "", unidad: "", cantidad: "", total: "" });
      $(`#p-${k} tbody`).innerHTML = filas.map((x, i) => `<tr data-lista="${k}" data-i="${i}"${x.grupo ? ` class="de-archivo" title="Del archivo SIR · ${esc(x.grupo)}"` : ""}>
        <td class="col-quitar"><button type="button" class="quitar" data-quitar title="Eliminar fila" aria-label="Eliminar fila ${esc(x.articulo || i + 1)}">✕</button></td>
        <td><input data-campo="articulo" value="${esc(x.articulo)}" placeholder="Artículo"></td>
        <td><select data-campo="unidad"><option value=""></option>${(x.unidad && !UNIDADES.includes(x.unidad) ? [x.unidad] : []).concat(UNIDADES).map((u) => `<option${u === x.unidad ? " selected" : ""}>${esc(u)}</option>`).join("")}</select></td>
        <td class="num"><input data-campo="cantidad" type="number" step="0.01" value="${esc(x.cantidad)}"></td>
        <td class="num"><input data-campo="total" type="number" step="0.01" value="${esc(x.total)}"></td>
        <td class="num costo"></td><td class="num pct"></td></tr>`).join("");
    });
    actualizarProducto();
  }
  // Recalcula costos, porcentajes, efectividad y el dashboard sin redibujar las filas (no se pierde el foco)
  function actualizarProducto() {
    const p = producto(), venta = parseFloat(p.venta) || 0;
    ["sobrantes", "faltantes"].forEach((k) => {
      $$(`#p-${k} tbody tr`).forEach((tr) => {
        const x = p[k][+tr.dataset.i] || {}, tot = parseFloat(x.total), cant = parseFloat(x.cantidad);
        tr.querySelector(".costo").textContent = tot && cant ? fmtUSD(tot / cant) : "";
        tr.querySelector(".pct").textContent = tot && venta ? fmtPct(tot / venta) : "";
      });
      const corto = k === "sobrantes" ? "sob" : "fal", total = sumaLista(p[k]);
      $(`#p-${corto}-total`).textContent = fmtUSD(total);
      $(`#p-${corto}-pct`).textContent = venta ? fmtPct(total / venta) : "—";
    });
    const e = efectividadDe(p);
    $$(".ef-valor").forEach((b) => { b.textContent = fmtPct(e.ef / 100); b.className = "ef-valor" + (isNaN(e.ef) ? "" : " " + nivel(e.ef)); });
    $$(".ef-puntos").forEach((s) => (s.textContent = isNaN(e.ef) ? "Ingresa ítems y error" : `${e.error} error(es) en ${e.items} ítems`));
    renderScore();
    if (!$('section[data-panel="producto"]').hidden) renderDashboard();
  }

  function renderDifCostos(co) {
    ["adm", "aud", "per"].forEach((k) => {
      const d = difCosto(co[k + "Ideal"], co[k + "Real"]), td = $(`[data-dif="${k}"]`);
      td.textContent = d ? `${fmtP(Math.abs(d.v))} ${d.tipo}` : "—";
      td.className = "num dif" + (d && d.v ? " " + d.tipo : "");
    });
  }
  const evCosto = (e) => {
    const i = e.target.closest("[data-costo]"); if (!i) return;
    if (!localActual) { toast("Selecciona un local primero"); return; }
    const p = producto(); p.costos = Object.assign({}, p.costos, { [i.dataset.costo]: i.value });
    guardarProducto(p); renderDifCostos(p.costos);
  };
  $(".costos").addEventListener("input", evCosto);
  $(".costos").addEventListener("change", evCosto);
  $("#p-tipo").addEventListener("change", () => { const p = producto(); p.tipo = $("#p-tipo").value; guardarProducto(p); });
  $("#p-venta").addEventListener("input", () => { const p = producto(); p.venta = $("#p-venta").value; guardarProducto(p); actualizarProducto(); });
  $("#p-fecha").addEventListener("change", () => { const p = producto(); p.fecha = $("#p-fecha").value; guardarProducto(p); });
  $("#p-local").addEventListener("change", () => {
    const l = locales.find((x) => x.local.toLowerCase() === $("#p-local").value.trim().toLowerCase());
    if (l) seleccionarLocal(l.local, false); else toast("Local no encontrado");
  });
  // Ítems / Error aparecen en Resultado y en Producto; ambos editan el mismo dato
  document.addEventListener("input", (e) => {
    const campo = e.target.classList.contains("ef-items") ? "items" : e.target.classList.contains("ef-error") ? "error" : null;
    if (!campo) return;
    if (!localActual) { toast("Selecciona un local primero"); return; }
    const p = producto(); p[campo] = e.target.value; guardarProducto(p);
    $$(`.ef-${campo}`).forEach((i) => { if (i !== e.target) i.value = e.target.value; });
    actualizarProducto();
  });
  ["sobrantes", "faltantes"].forEach((k) => {
    const tb = $(`#p-${k} tbody`);
    const editar = (e) => {
      const el = e.target.closest("[data-campo]"); if (!el) return;
      if (!localActual) { toast("Selecciona un local primero"); return; }
      const tr = el.closest("tr"), p = producto(), i = +tr.dataset.i;
      while (p[k].length <= i) p[k].push({ articulo: "", unidad: "", cantidad: "", total: "" });
      p[k][i][el.dataset.campo] = el.value; guardarProducto(p); actualizarProducto();
    };
    tb.addEventListener("input", editar);
    tb.addEventListener("change", editar);
    tb.addEventListener("click", (e) => {
      if (!e.target.closest("[data-quitar]") || !localActual) return;
      const p = producto(), i = +e.target.closest("tr").dataset.i;
      if (i >= p[k].length) { toast("Fila vacía"); return; } // filas de relleno que aún no tienen datos
      const [quitada] = p[k].splice(i, 1), local = localActual.local;
      guardarProducto(p); renderProducto();
      toast(`Fila eliminada: ${quitada.articulo || "sin nombre"}`, "Deshacer", () => {
        if (!localActual || localActual.local !== local) return;
        const q = producto(); q[k].splice(Math.min(i, q[k].length), 0, quitada); guardarProducto(q); renderProducto();
        toast("Fila recuperada");
      });
    });
  });
  $$("[data-agregar]").forEach((b) => b.addEventListener("click", () => {
    if (!localActual) { toast("Selecciona un local primero"); return; }
    const p = producto(), k = b.dataset.agregar;
    const n = Math.max(p[k].length, $$(`#p-${k} tbody tr`).length);
    while (p[k].length <= n) p[k].push({ articulo: "", unidad: "", cantidad: "", total: "" });
    guardarProducto(p); renderProducto();
    const filas = $$(`#p-${k} tbody tr`); filas[filas.length - 1].querySelector("input").focus();
  }));
  $("#p-registrar").addEventListener("click", () => registrarMes(false));

  // ---------- Carga del reporte «Usos y consumos» del SIR ----------
  // El .xls del SIR es una página HTML con una tabla. Cada artículo trae 23 celdas: grupo, nombre, unidad,
  // máximo, mínimo y 9 pares costo/cantidad; el último par es la Diferencia.
  // Regla de auditoría: diferencia positiva = FALTANTE, diferencia negativa = SOBRANTE.
  const UNIDAD_SIR = { KG: "Kilos", KILO: "Kilos", KILOS: "Kilos", LITRO: "Litros", LITROS: "Litros", LT: "Litros", UNIDAD: "Unidad", UND: "Unidad",
    FUNDA: "Fundas", FDA1KG: "Fundas", GALON: "Galones", GL: "Galones", LBS: "Libras", LIBRA: "Libras", GRAMO: "Gramos", GR: "Gramos",
    PAQUETE: "Paquetes", PAQ: "Paquetes", CAJA: "Caja" };
  const numSIR = (s) => { const v = parseFloat(String(s).replace(/[$\s]/g, "").replace(/\./g, "").replace(",", ".")); return isNaN(v) ? 0 : v; };
  const r2 = (v) => Math.round(v * 100) / 100;
  const isoDe = (dmy) => { const m = /(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(dmy || ""); return m ? `${m[3].length === 2 ? "20" + m[3] : m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` : ""; };

  function leerUsosConsumos(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const todo = doc.body ? doc.body.textContent.replace(/\s+/g, " ") : "";
    const local = (/Local:\s*(.+?)(?=\s*Usuario:|$)/.exec(todo) || [])[1] || "";
    const fechas = /Fecha Desde:\s*(\S+)\s*Hasta:\s*(\S+)/.exec(todo) || [];
    let grupo = "";
    const articulos = [];
    doc.querySelectorAll("tr").forEach((tr) => {
      const c = [...tr.children].map((td) => td.textContent.replace(/\s+/g, " ").trim());
      if (c.length !== 23 || !c[1] || /total/i.test(c[3]) || c[0] === "Grupos") return;
      if (c[0]) grupo = c[0];
      articulos.push({ grupo, articulo: c[1], unidadSIR: c[2], dif$: numSIR(c[21]), difCant: numSIR(c[22]) });
    });
    return { local: local.trim(), desde: fechas[1] || "", hasta: fechas[2] || "", articulos };
  }
  async function leerArchivoTexto(file) {
    const buf = await file.arrayBuffer();
    let t = new TextDecoder("utf-8").decode(buf);
    if (t.includes("�")) t = new TextDecoder("windows-1252").decode(buf); // exportes con tildes en ANSI
    return t;
  }
  function aFila(x) {
    const u = x.unidadSIR.toUpperCase().replace(/\s+/g, "");
    return { articulo: x.articulo, unidad: UNIDAD_SIR[u] || x.unidadSIR, cantidad: String(r2(Math.abs(x.difCant))), total: String(r2(Math.abs(x.dif$))), grupo: x.grupo };
  }
  async function cargarUsosConsumos(file) {
    if (!file) return;
    let datos;
    try { datos = leerUsosConsumos(await leerArchivoTexto(file)); } catch { datos = { articulos: [] }; }
    if (!datos.articulos.length) { toast("No reconozco el archivo: debe ser «Usos y consumos» exportado del SIR"); return; }
    // Local del archivo (p. ej. «K004 _ CCI») → local de la matriz con el mismo código
    const codigo = (/^([A-Z0-9]{3,5})/i.exec(datos.local) || [])[1] || "";
    const delArchivo = codigo && locales.find((l) => l.local.toUpperCase().startsWith(codigo.toUpperCase() + " "));
    if (delArchivo && (!localActual || localActual.local !== delArchivo.local)) {
      if (!localActual || confirm(`El archivo es del local ${datos.local}. ¿Cambiar a ${delArchivo.local}?`)) seleccionarLocal(delArchivo.local, false);
    }
    if (!localActual) { toast("Selecciona primero el local del archivo"); return; }
    const p = producto();
    if ((p.sobrantes.some(filaLlena) || p.faltantes.some(filaLlena)) && !confirm("Ya hay sobrantes o faltantes en este local. ¿Reemplazarlos con el TOP del archivo?")) return;
    const top = +$("#p-top").value || 10;
    const fal = datos.articulos.filter((x) => x.dif$ > 0).sort((a, b) => b.dif$ - a.dif$);
    const sob = datos.articulos.filter((x) => x.dif$ < 0).sort((a, b) => a.dif$ - b.dif$);
    p.faltantes = fal.slice(0, top).map(aFila);
    p.sobrantes = sob.slice(0, top).map(aFila);
    if (datos.hasta) p.fecha = isoDe(datos.hasta);
    p.archivo = { nombre: file.name, local: datos.local, desde: datos.desde, hasta: datos.hasta, articulos: datos.articulos.length, top,
      nFal: fal.length, nSob: sob.length, totFal: r2(fal.reduce((s, x) => s + x.dif$, 0)), totSob: r2(-sob.reduce((s, x) => s + x.dif$, 0)),
      cand: { fal: fal.slice(0, 20).map(aFila), sob: sob.slice(0, 20).map(aFila) } }; // para cambiar el top sin volver a subir
    guardarProducto(p); renderProducto();
    toast(`Top ${top}: ${p.faltantes.length} faltantes y ${p.sobrantes.length} sobrantes cargados`);
  }
  $("#p-top").addEventListener("change", () => {
    const p = producto(), top = +$("#p-top").value;
    if (!p.archivo || !p.archivo.cand) return;
    if (!confirm(`¿Cambiar al top ${top}? Se vuelven a tomar los valores del archivo y se pierden los cambios hechos a mano.`)) { $("#p-top").value = p.archivo.top; return; }
    p.faltantes = p.archivo.cand.fal.slice(0, top).map((x) => Object.assign({}, x));
    p.sobrantes = p.archivo.cand.sob.slice(0, top).map((x) => Object.assign({}, x));
    p.archivo.top = top; guardarProducto(p); renderProducto();
  });
  $("#p-subir").addEventListener("click", () => $("#p-archivo").click());
  $("#p-archivo").addEventListener("change", (e) => { cargarUsosConsumos(e.target.files[0]); e.target.value = ""; });
  const zona = $("#p-carga");
  ["dragenter", "dragover"].forEach((t) => zona.addEventListener(t, (e) => { e.preventDefault(); zona.classList.add("arrastre"); }));
  ["dragleave", "drop"].forEach((t) => zona.addEventListener(t, () => zona.classList.remove("arrastre")));
  zona.addEventListener("drop", (e) => { e.preventDefault(); cargarUsosConsumos(e.dataTransfer.files[0]); });

  // ---------- Dashboard de pastel: venta neta vs sobrantes y faltantes TOP ----------
  $$('input[name="d-alcance"]').forEach((i) => i.addEventListener("change", renderDashboard));
  function datosDashboard(alcance) {
    const todos = (alcance || $('input[name="d-alcance"]:checked').value) === "todos";
    if (!todos) {
      const p = producto();
      return { venta: parseFloat(p.venta) || 0, sob: sumaLista(p.sobrantes), fal: sumaLista(p.faltantes), n: localActual ? 1 : 0,
        sub: localActual ? `${localActual.local} · ${p.tipo}` : "Elige un local" };
    }
    const pref = `matriz:${usuario.usuario}:`, nombres = new Set(locales.map((l) => l.local));
    let venta = 0, sob = 0, fal = 0, n = 0;
    const filas = [];
    Object.keys(localStorageSeguro()).forEach((k) => {
      if (!k.startsWith(pref) || !nombres.has(k.slice(pref.length))) return;
      const p = store.get(k, {}).producto; if (!p) return;
      const v = parseFloat(p.venta) || 0, s = sumaLista(p.sobrantes || []), f = sumaLista(p.faltantes || []);
      if (!v && !s && !f) return;
      venta += v; sob += s; fal += f; n++;
      filas.push({ local: k.slice(pref.length), venta: v, sob: s, fal: f });
    });
    filas.sort((a, b) => a.local.localeCompare(b.local, "es"));
    return { todos: true, venta, sob, fal, n, filas, sub: `${n} local(es) con datos de producto registrados por ${usuario.nombre}` };
  }
  function renderDashboard() {
    const d = datosDashboard();
    const resto = Math.max(0, d.venta - d.sob - d.fal);
    const base = d.venta > 0 ? Math.max(d.venta, d.sob + d.fal) : d.sob + d.fal;
    const pct = (v) => (d.venta > 0 ? v / d.venta : NaN);
    $("#d-sub").textContent = d.sub;
    $("#d-kpis").innerHTML = [
      [fmtUSD(d.venta), "Venta neta"],
      [fmtUSD(d.fal), `Faltante TOP · ${fmtPct(pct(d.fal))}`],
      [fmtUSD(d.sob), `Sobrante TOP · ${fmtPct(pct(d.sob))}`],
      [fmtUSD(d.sob - d.fal), `Diferencia neta · ${fmtPct(pct(d.sob - d.fal))}`],
    ].map(([v, t]) => `<div><b>${esc(v)}</b><span>${esc(t)}</span></div>`).join("");
    const rebanadas = [
      { k: "fal", nombre: "Faltante TOP", v: d.fal, color: "var(--c-fal)" },
      { k: "sob", nombre: "Sobrante TOP", v: d.sob, color: "var(--c-sob)" },
      { k: "resto", nombre: "Resto de la venta neta", v: resto, color: "var(--c-resto)" },
    ].filter((r) => r.v > 0);
    $("#d-leyenda").innerHTML = rebanadas.length ? rebanadas.map((r) => `<li><i style="background:${r.color}"></i><span>${esc(r.nombre)}</span>
      <b>${fmtUSD(r.v)}</b><em>${d.venta > 0 ? fmtPct(r.v / d.venta) : fmtPct(r.v / base)}</em></li>`).join("")
      + (d.venta > 0 && d.sob + d.fal > d.venta ? `<li class="muted small">Sobrantes + faltantes superan la venta neta.</li>` : "")
      : `<li class="muted">Ingresa la venta neta y los productos TOP para ver el gráfico.</li>`;
    const box = $("#d-pie"), R = 100, C = 110;
    if (!rebanadas.length) { box.innerHTML = `<svg viewBox="0 0 220 220" role="img" aria-label="Sin datos"><circle cx="${C}" cy="${C}" r="${R}" class="pie-vacio"/></svg>`; return; }
    let ang = -Math.PI / 2, s = `<svg viewBox="0 0 220 220" role="img" aria-label="Venta neta: faltante ${fmtPct(pct(d.fal))}, sobrante ${fmtPct(pct(d.sob))}">`;
    rebanadas.forEach((r, i) => {
      const frac = r.v / base, a2 = ang + frac * 2 * Math.PI;
      const tip = `${r.nombre}: ${fmtUSD(r.v)} (${fmtPct(d.venta > 0 ? r.v / d.venta : frac)})`;
      if (frac >= 0.9999) s += `<circle data-i="${i}" cx="${C}" cy="${C}" r="${R}" fill="${r.color}" class="rebanada"><title>${esc(tip)}</title></circle>`;
      else {
        const x1 = C + R * Math.cos(ang), y1 = C + R * Math.sin(ang), x2 = C + R * Math.cos(a2), y2 = C + R * Math.sin(a2);
        s += `<path data-i="${i}" class="rebanada" fill="${r.color}" d="M${C},${C} L${x1},${y1} A${R},${R} 0 ${a2 - ang > Math.PI ? 1 : 0} 1 ${x2},${y2} Z"><title>${esc(tip)}</title></path>`;
      }
      // Porcentaje dentro de la rebanada cuando hay espacio
      if (frac >= 0.07 && r.k !== "resto") {
        const am = (ang + a2) / 2;
        s += `<text class="pie-txt" x="${C + R * 0.62 * Math.cos(am)}" y="${C + R * 0.62 * Math.sin(am) + 4}" text-anchor="middle">${fmtPct(d.venta > 0 ? r.v / d.venta : frac)}</text>`;
      }
      ang = a2;
    });
    box.innerHTML = s + "</svg>";
  }

  // ---------- Reporte para copiar / pegar en el correo ----------
  // Colores fijos de fondo claro: el correo se lee sobre blanco aunque la página esté en modo oscuro
  const RC = { fal: "#d9622b", sob: "#2a62c9", resto: "#d5dae1", txt: "#1b1f24", muted: "#646d78", borde: "#dfe3e8", fondo: "#f5f6f8", ok: "#1f7a4d", warn: "#9a6300", bad: "#b42318" };
  const FUENTE = '"Segoe UI", Arial, sans-serif';

  function datosReporte(alcance) {
    const d = datosDashboard(alcance), p = producto();
    const resto = Math.max(0, d.venta - d.sob - d.fal);
    const base = d.venta > 0 ? Math.max(d.venta, d.sob + d.fal) : d.sob + d.fal;
    const pv = (v) => (d.venta > 0 ? v / d.venta : base ? v / base : NaN);
    const rebanadas = [["Faltante TOP", d.fal, RC.fal], ["Sobrante TOP", d.sob, RC.sob], ["Resto de la venta neta", resto, RC.resto]]
      .map(([nombre, v, color]) => ({ nombre, v, color, pct: pv(v), frac: base ? v / base : 0 }));
    return {
      d, p, todos: !!d.todos, pv, rebanadas, base, efx: efectividadDe(p),
      titulo: d.todos ? "Producto · todos mis locales" : `${p.tipo} · ${localActual ? localActual.local : ""}`,
      sobFilas: p.sobrantes.filter(filaLlena), falFilas: p.faltantes.filter(filaLlena),
      fecha: fechaLarga((p.costos && p.costos.audFecha) || p.fecha || $("#sel-fecha").value || hoy()),
    };
  }

  function dibujarReporte(alcance) {
    const r = datosReporte(alcance), d = r.d;
    if (!base0(r)) return null;
    const W = 960, pad = 28, S = ESCALA_COPIA, ANCHO = W - 2 * pad, GAP = 24;
    // Filas de texto ya formateadas: se miden antes de dibujar para que ningún número se corte
    const venta = d.venta;
    const filaProd = (x) => {
      const t = parseFloat(x.total) || 0, c = parseFloat(x.cantidad) || 0;
      return [x.articulo, x.unidad || "", c ? c.toLocaleString("es-EC", { maximumFractionDigits: 2 }) : "", fmtUSD(t), c && t ? fmtUSD(t / c) : "", venta && t ? fmtPct(t / venta) : ""];
    };
    const cab = r.todos ? [["Local", "l"], ["Venta neta", "r"], ["Faltante $", "r"], ["%", "r"], ["Sobrante $", "r"], ["%", "r"]]
      : [["Artículo", "l"], ["Unidad", "l"], ["Cantidad", "r"], ["Total", "r"], ["Costo U.", "r"], ["% venta", "r"]];
    const bloques = r.todos
      ? [{ titulo: "Detalle por local", color: RC.txt, resumen: "", filas: d.filas.map((f) => [f.local, fmtUSD(f.venta), fmtUSD(f.fal), f.venta ? fmtPct(f.fal / f.venta) : "—", fmtUSD(f.sob), f.venta ? fmtPct(f.sob / f.venta) : "—"]) }]
      : [{ titulo: "Sobrantes producto TOP", color: RC.sob, resumen: `${fmtPct(r.pv(d.sob))} · ${fmtUSD(d.sob)}`, filas: r.sobFilas.map(filaProd) },
         { titulo: "Faltante producto TOP", color: RC.fal, resumen: `${fmtPct(r.pv(d.fal))} · ${fmtUSD(d.fal)}`, filas: r.falFilas.map(filaProd) }];
    const med = document.createElement("canvas").getContext("2d");
    const ancho = (t, px, peso) => { med.font = `${peso} ${px}px ${FUENTE}`; return med.measureText(String(t)).width; };
    const todas = bloques.flatMap((b) => b.filas);
    const necesario = cab.map(([n], j) => Math.ceil(Math.max(ancho(n.toUpperCase(), 11, 700), ...todas.map((f) => ancho(f[j], 13, 400)))) + 18);
    // Lado a lado solo si todo (incluido el nombre completo) cabe en media página; si no, una tabla debajo de otra
    const mitad = (ANCHO - GAP) / 2;
    const ladoALado = bloques.length === 2 && necesario.reduce((s, w) => s + w, 0) <= mitad;
    const anchoTabla = ladoALado ? mitad : ANCHO;
    const total = necesario.reduce((s, w) => s + w, 0);
    let col;
    if (total <= anchoTabla) { // sobra espacio: 40% al nombre y 60% repartido entre las demás columnas
      const extra = anchoTabla - total, n = cab.length - 1;
      col = cab.map(([nom, al], j) => [nom, necesario[j] + (j === 0 ? extra * 0.4 : (extra * 0.6) / n), al]);
    } else { // no cabe: los números se mantienen completos y solo se recorta el nombre
      const otras = total - necesario[0];
      col = cab.map(([nom, al], j) => [nom, j === 0 ? Math.max(120, anchoTabla - otras) : necesario[j], al]);
    }
    const altoBloque = (b) => 12 + 28 + Math.max(b.filas.length, 1) * 26;
    const altoTablas = ladoALado ? Math.max(...bloques.map(altoBloque)) : bloques.reduce((s, b) => s + altoBloque(b) + 44, -44);
    const lineaCostos = r.todos ? "" : textoCostos(r.p.costos || {});
    const H = 96 + 300 + 30 + altoTablas + 40 + 60 + (lineaCostos ? 26 : 0);
    const cv = document.createElement("canvas");
    cv.width = W * S; cv.height = H * S;
    const g = cv.getContext("2d"); g.scale(S, S); g.imageSmoothingQuality = "high";
    const font = (px, peso = 400) => (g.font = `${peso} ${px}px ${FUENTE}`);
    const texto = (t, x, y, { px = 14, peso = 400, color = RC.txt, alinear = "left", max } = {}) => {
      font(px, peso); g.fillStyle = color; g.textAlign = alinear; g.textBaseline = "alphabetic";
      let s = String(t);
      if (max) while (s.length > 1 && g.measureText(s).width > max) s = s.slice(0, -2) + "…";
      g.fillText(s, x, y);
    };
    const caja = (x, y, w, h, rad, relleno, borde) => {
      g.beginPath(); g.roundRect ? g.roundRect(x, y, w, h, rad) : g.rect(x, y, w, h);
      if (relleno) { g.fillStyle = relleno; g.fill(); }
      if (borde) { g.strokeStyle = borde; g.lineWidth = 1; g.stroke(); }
    };
    g.fillStyle = "#ffffff"; g.fillRect(0, 0, W, H);
    // Encabezado
    g.fillStyle = "#c8102e"; g.fillRect(0, 0, W, 6);
    texto("Venta neta vs. sobrantes y faltantes TOP", pad, 42, { px: 22, peso: 700 });
    texto(`${r.titulo}  ·  ${r.fecha}  ·  ${usuario.nombre}`, pad, 66, { px: 14, color: RC.muted, max: W - 2 * pad });
    g.fillStyle = RC.borde; g.fillRect(pad, 80, W - 2 * pad, 1);
    // Pastel
    const y0 = 96, cx = pad + 140, cy = y0 + 140, R = 130;
    let ang = -Math.PI / 2;
    r.rebanadas.filter((x) => x.v > 0).forEach((x) => {
      const a2 = ang + x.frac * 2 * Math.PI;
      g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, R, ang, a2); g.closePath();
      g.fillStyle = x.color; g.fill(); g.strokeStyle = "#fff"; g.lineWidth = 2; g.stroke();
      if (x.frac >= 0.07 && x.color !== RC.resto) {
        const am = (ang + a2) / 2;
        texto(fmtPct(x.pct), cx + R * 0.62 * Math.cos(am), cy + R * 0.62 * Math.sin(am) + 6, { px: 16, peso: 700, color: "#fff", alinear: "center" });
      }
      ang = a2;
    });
    // Indicadores 2x2
    const xR = pad + 310, wR = W - pad - xR, wK = (wR - 12) / 2;
    [[fmtUSD(d.venta), "Venta neta"], [fmtUSD(d.fal), `Faltante TOP · ${fmtPct(r.pv(d.fal))}`],
     [fmtUSD(d.sob), `Sobrante TOP · ${fmtPct(r.pv(d.sob))}`], [fmtUSD(d.sob - d.fal), `Diferencia neta · ${fmtPct(r.pv(d.sob - d.fal))}`]]
      .forEach(([v, t], i) => {
        const x = xR + (i % 2) * (wK + 12), y = y0 + Math.floor(i / 2) * 82;
        caja(x, y, wK, 70, 10, RC.fondo, RC.borde);
        texto(v, x + 16, y + 34, { px: 22, peso: 700 });
        texto(t, x + 16, y + 56, { px: 13, color: RC.muted });
      });
    // Leyenda
    r.rebanadas.forEach((x, i) => {
      const y = y0 + 186 + i * 28;
      caja(xR, y - 12, 14, 14, 3, x.color);
      texto(x.nombre, xR + 24, y, { px: 15 });
      texto(fmtUSD(x.v), W - pad - 90, y, { px: 15, peso: 700, alinear: "right" });
      texto(fmtPct(x.pct), W - pad, y, { px: 15, color: RC.muted, alinear: "right" });
    });
    // Tablas
    const tabla = (x, y, titulo, color, resumen, filas) => {
      texto(titulo, x, y, { px: 17, peso: 700, color });
      if (resumen) texto(resumen, x + col.reduce((s, c) => s + c[1], 0), y, { px: 13, peso: 600, color: RC.muted, alinear: "right" });
      let yy = y + 12;
      caja(x, yy, col.reduce((s, c) => s + c[1], 0), 28, 4, RC.fondo);
      let xx = x;
      col.forEach(([n, w, al]) => { texto(n.toUpperCase(), al === "r" ? xx + w - 8 : xx + 8, yy + 19, { px: 11, peso: 700, color: RC.muted, alinear: al === "r" ? "right" : "left" }); xx += w; });
      yy += 28;
      filas.forEach((f) => {
        xx = x;
        f.forEach((v, j) => { const [, w, al] = col[j]; texto(v, al === "r" ? xx + w - 8 : xx + 8, yy + 18, { px: 13, alinear: al === "r" ? "right" : "left", max: w - 12 }); xx += w; });
        yy += 26; g.fillStyle = RC.borde; g.fillRect(x, yy - 1, col.reduce((s, c) => s + c[1], 0), 1);
      });
      if (!filas.length) texto("Sin registros", x + 8, yy + 18, { px: 13, color: RC.muted });
    };
    let yT = y0 + 300 + 30;
    bloques.forEach((b, i) => {
      if (ladoALado) tabla(i === 0 ? pad : W - pad - anchoTabla, yT, b.titulo, b.color, b.resumen, b.filas);
      else { tabla(pad, yT, b.titulo, b.color, b.resumen, b.filas); yT += altoBloque(b) + 44; }
    });
    // Pie de página
    const yF = H - 30;
    g.fillStyle = RC.borde; g.fillRect(pad, yF - (lineaCostos ? 52 : 26), W - 2 * pad, 1);
    if (lineaCostos) texto(lineaCostos, pad, yF - 26, { px: 14, peso: 600, max: W - 2 * pad });
    if (!r.todos && !isNaN(r.efx.ef)) {
      texto(`Efectividad inventarios totales: ${fmtPct(r.efx.ef / 100)}  (${r.efx.items} ítems, ${r.efx.error} error(es))`, pad, yF, { px: 14, peso: 600, color: r.efx.ef >= 90 ? RC.ok : r.efx.ef >= 75 ? RC.warn : RC.bad });
    }
    texto("Auditoría · Grupo KFC Ecuador", W - pad, yF, { px: 12, color: RC.muted, alinear: "right" });
    return cv;
  }
  const base0 = (r) => r.d.venta > 0 || r.d.sob > 0 || r.d.fal > 0;
  // El reporte va en el correo cuando el enfoque es de producto y hay datos del local
  const conReporte = () => catCorreo() === "producto" && !!localActual && base0(datosReporte("local"));

  // Versión en tablas HTML con estilos en línea (Outlook usa Word para mostrar el correo: nada de flex ni grid)
  function htmlReporte(imgSrc, alcance) {
    const r = datosReporte(alcance), d = r.d;
    const td = (v, extra = "") => `<td style="padding:5px 8px;border-bottom:1px solid ${RC.borde};font-size:10pt;${extra.includes("right") ? "white-space:nowrap;" : ""}${extra}">${esc(v)}</td>`;
    const th = (v, der) => `<th style="padding:6px 8px;background:${RC.fondo};color:${RC.muted};font-size:8.5pt;text-transform:uppercase;text-align:${der ? "right" : "left"}">${esc(v)}</th>`;
    const tablaProd = (titulo, color, filas, total) => `
      <p style="margin:14px 0 6px;font-size:12pt;font-weight:bold;color:${color}">${esc(titulo)}
        <span style="font-size:10pt;color:${RC.muted};font-weight:normal">&nbsp; Afectación al costo ${fmtPct(r.pv(total))} · ${fmtUSD(total)}</span></p>
      <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font-family:${FUENTE}">
        <tr>${th("Artículo")}${th("Unidad")}${th("Cantidad", 1)}${th("Total", 1)}${th("Costo U.", 1)}${th("% venta", 1)}</tr>
        ${filas.length ? filas.map((x) => {
          const t = parseFloat(x.total) || 0, c = parseFloat(x.cantidad) || 0, R = "text-align:right";
          return `<tr>${td(x.articulo)}${td(x.unidad || "")}${td(c ? c.toLocaleString("es-EC", { maximumFractionDigits: 2 }) : "", R)}${td(fmtUSD(t), R)}${td(c && t ? fmtUSD(t / c) : "", R)}${td(d.venta && t ? fmtPct(t / d.venta) : "", R)}</tr>`;
        }).join("") : `<tr>${td("Sin registros", `color:${RC.muted}`)}</tr>`}
      </table>`;
    const kpi = (v, t) => `<td style="padding:10px 14px;border:1px solid ${RC.borde};background:${RC.fondo};width:50%">
      <div style="font-size:15pt;font-weight:bold;color:${RC.txt}">${esc(v)}</div><div style="font-size:9pt;color:${RC.muted}">${esc(t)}</div></td>`;
    // Barra apilada en tabla: se ve aunque el correo bloquee la imagen
    const barra = r.rebanadas.filter((x) => x.frac > 0).map((x) => `<td style="background:${x.color};height:12px;width:${Math.max(1, Math.round(x.frac * 100))}%;font-size:1px">&nbsp;</td>`).join("");
    const leyenda = r.rebanadas.map((x) => `<tr><td style="padding:3px 0;width:18px"><div style="width:12px;height:12px;background:${x.color}"></div></td>
      <td style="padding:3px 6px;font-size:10pt">${esc(x.nombre)}</td><td style="padding:3px 6px;font-size:10pt;font-weight:bold;text-align:right">${fmtUSD(x.v)}</td>
      <td style="padding:3px 0 3px 8px;font-size:10pt;color:${RC.muted};text-align:right">${fmtPct(x.pct)}</td></tr>`).join("");
    const detalle = r.todos
      ? `<p style="margin:14px 0 6px;font-size:12pt;font-weight:bold">Detalle por local</p>
         <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font-family:${FUENTE}">
         <tr>${th("Local")}${th("Venta neta", 1)}${th("Faltante", 1)}${th("%", 1)}${th("Sobrante", 1)}${th("%", 1)}</tr>
         ${d.filas.map((f) => `<tr>${td(f.local)}${td(fmtUSD(f.venta), "text-align:right")}${td(fmtUSD(f.fal), "text-align:right")}${td(f.venta ? fmtPct(f.fal / f.venta) : "—", "text-align:right")}${td(fmtUSD(f.sob), "text-align:right")}${td(f.venta ? fmtPct(f.sob / f.venta) : "—", "text-align:right")}</tr>`).join("")}</table>`
      : tablaProd("Sobrantes producto TOP", RC.sob, r.sobFilas, d.sob) + tablaProd("Faltante producto TOP", RC.fal, r.falFilas, d.fal);
    const lc = r.todos ? "" : textoCostos(r.p.costos || {});
    const ef = (lc ? `<p style="margin:12px 0 0;font-size:10pt;font-weight:bold">${esc(lc)}</p>` : "") + (!r.todos && !isNaN(r.efx.ef)
      ? `<p style="margin:12px 0 0;font-size:10pt;font-weight:bold;color:${r.efx.ef >= 90 ? RC.ok : r.efx.ef >= 75 ? RC.warn : RC.bad}">Efectividad inventarios totales: ${fmtPct(r.efx.ef / 100)} (${r.efx.items} ítems, ${r.efx.error} error(es))</p>` : "");
    return `<div style="font-family:${FUENTE};color:${RC.txt};max-width:720px">
      <table cellpadding="0" cellspacing="0" style="width:100%;border-top:4px solid #c8102e"><tr><td style="padding:10px 0 2px;font-size:15pt;font-weight:bold">Venta neta vs. sobrantes y faltantes TOP</td></tr>
      <tr><td style="padding:0 0 10px;font-size:10pt;color:${RC.muted}">${esc(r.titulo)} · ${esc(r.fecha)} · ${esc(usuario.nombre)}</td></tr></table>
      <table cellpadding="0" cellspacing="0" style="width:100%"><tr>
        ${imgSrc ? `<td style="width:230px;vertical-align:top;padding-right:16px"><img src="${imgSrc}" width="220" height="220" alt="Gráfico de pastel: faltante ${fmtPct(r.pv(d.fal))}, sobrante ${fmtPct(r.pv(d.sob))}" style="display:block"></td>` : ""}
        <td style="vertical-align:top">
          <table cellpadding="0" cellspacing="6" style="width:100%"><tr>${kpi(fmtUSD(d.venta), "Venta neta")}${kpi(fmtUSD(d.fal), `Faltante TOP · ${fmtPct(r.pv(d.fal))}`)}</tr>
          <tr>${kpi(fmtUSD(d.sob), `Sobrante TOP · ${fmtPct(r.pv(d.sob))}`)}${kpi(fmtUSD(d.sob - d.fal), `Diferencia neta · ${fmtPct(r.pv(d.sob - d.fal))}`)}</tr></table>
          <table cellpadding="0" cellspacing="0" style="width:100%;margin-top:8px"><tr>${barra}</tr></table>
          <table cellpadding="0" cellspacing="0" style="margin-top:6px">${leyenda}</table>
        </td></tr></table>
      ${detalle}${ef}</div>`;
  }
  // Pastel suelto (para incrustar en la versión de tabla)
  function pastelPNG(alcance) {
    const r = datosReporte(alcance), S = 3, T = 220, cv = document.createElement("canvas");
    cv.width = T * S; cv.height = T * S;
    const g = cv.getContext("2d"); g.scale(S, S); g.imageSmoothingQuality = "high"; g.fillStyle = "#fff"; g.fillRect(0, 0, T, T);
    let ang = -Math.PI / 2;
    r.rebanadas.filter((x) => x.v > 0).forEach((x) => {
      const a2 = ang + x.frac * 2 * Math.PI;
      g.beginPath(); g.moveTo(110, 110); g.arc(110, 110, 104, ang, a2); g.closePath();
      g.fillStyle = x.color; g.fill(); g.strokeStyle = "#fff"; g.lineWidth = 2; g.stroke();
      if (x.frac >= 0.07 && x.color !== RC.resto) {
        const am = (ang + a2) / 2;
        g.font = `700 14px ${FUENTE}`; g.fillStyle = "#fff"; g.textAlign = "center";
        g.fillText(fmtPct(x.pct), 110 + 64 * Math.cos(am), 110 + 64 * Math.sin(am) + 5);
      }
      ang = a2;
    });
    return cv.toDataURL("image/png");
  }
  const textoReporte = () => {
    const r = datosReporte(), d = r.d;
    return [`Venta neta vs. sobrantes y faltantes TOP — ${r.titulo}`, `Venta neta: ${fmtUSD(d.venta)}`,
      `Faltante TOP: ${fmtUSD(d.fal)} (${fmtPct(r.pv(d.fal))})`, `Sobrante TOP: ${fmtUSD(d.sob)} (${fmtPct(r.pv(d.sob))})`,
      `Diferencia neta: ${fmtUSD(d.sob - d.fal)}`].join("\n");
  };
  const canvasBlob = (cv) => new Promise((ok) => cv.toBlob(ok, "image/png"));
  function descargar(blob, nombre) {
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  const nombreArchivo = () => `Reporte producto ${datosReporte().titulo}`.replace(/[\\/:*?"<>|·]/g, "_").replace(/\s+/g, " ").trim();
  function sinDatos() { toast("Ingresa la venta neta o los productos TOP primero"); }

  $("#r-copiar-img").addEventListener("click", async () => {
    const cv = dibujarReporte(); if (!cv) return sinDatos();
    const blob = await canvasBlob(cv);
    try { await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]); toast("Imagen copiada: pégala en el correo con Ctrl+V"); }
    catch { descargar(blob, nombreArchivo() + ".png"); toast("Tu navegador no permitió copiar: se descargó la imagen"); }
  });
  $("#r-copiar-tabla").addEventListener("click", async () => {
    if (!base0(datosReporte())) return sinDatos();
    const html = htmlReporte(pastelPNG());
    try {
      await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([textoReporte()], { type: "text/plain" }) })]);
      toast("Tabla copiada: pégala en el correo con Ctrl+V");
    } catch {
      const div = document.createElement("div"); div.contentEditable = "true"; div.innerHTML = html;
      div.style.cssText = "position:fixed;left:-9999px;top:0;background:#fff"; document.body.appendChild(div);
      const rg = document.createRange(); rg.selectNodeContents(div); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(rg);
      document.execCommand("copy"); sel.removeAllRanges(); div.remove(); toast("Tabla copiada: pégala en el correo con Ctrl+V");
    }
  });
  $("#r-descargar").addEventListener("click", async () => {
    const cv = dibujarReporte(); if (!cv) return sinDatos();
    descargar(await canvasBlob(cv), nombreArchivo() + ".png"); toast("Imagen descargada");
  });

  // ---------- Tendencias ----------
  const MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  const MESES_LARGO = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const META = 90, ALERTA = 75;
  const RES = M.resumenMes || { mes: "", locales: [] };
  const claveHist = (l) => `matriz:${usuario.usuario}:hist:${l}`;
  const nivel = (v) => (v >= META ? "ok" : v >= ALERTA ? "warn" : "bad");
  const fmtMes = (ym) => { const [a, m] = ym.split("-"); return `${MESES_LARGO[+m - 1]} ${a}`; };

  // Historial de un local: lo registrado en la página + la nota del mes que trae la matriz Excel
  function historial(nombre) {
    const h = store.get(claveHist(nombre), {});
    const ex = RES.locales.find((x) => x.local === nombre);
    if (ex && ex.nota != null) {
      const prev = h[RES.mes] || {};
      if (prev.score == null) h[RES.mes] = Object.assign({}, prev, { score: ex.nota, fuente: "Excel" });
    }
    return h;
  }
  // Un proceso marcado «No cumple» necesita su observación antes de registrar o enviar
  function faltanObservaciones() {
    const falta = calcular().sinObs;
    if (!falta.length) return false;
    toast(`Escribe la observación de: ${falta.map((c) => c.area).join(", ")}`);
    irA("informe");
    const ta = $(`#checklist .proc-obs[data-n="${falta[0].n}"]`);
    ta.scrollIntoView({ block: "center" }); ta.focus();
    return true;
  }
  function registrarMes(silencioso) {
    if (!localActual) { toast("Selecciona un local primero"); return; }
    if (!silencioso && faltanObservaciones()) return;
    const r = calcular(), st = estadoLocal(localActual.local), co = producto().costos || {};
    const mes = ($("#sel-fecha").value || hoy()).slice(0, 7);
    const h = store.get(claveHist(localActual.local), {});
    h[mes] = Object.assign(h[mes] || {}, {
      score: r.score, novedades: r.items.length, efectividad: isNaN(r.ef) ? null : r.ef,
      activosNov: st.activos.filter(conNovedadActivo).length,
      informe: $("#sel-informe").value, fuente: "Página",
      venta: parseFloat(producto().venta) || null, sobrantes: sumaLista(producto().sobrantes), faltantes: sumaLista(producto().faltantes),
      // mermas, lunch y costo real pasan a la tabla de Tendencias (solo si se llenaron en Producto)
      ...Object.fromEntries([["mermas", co.merma], ["lunch", co.lunch], ["costo", co.audReal || co.perReal]]
        .filter(([, v]) => v !== "" && v != null).map(([k, v]) => [k, +v])),
    });
    store.set(claveHist(localActual.local), h);
    if (!silencioso) toast(`Registrado en ${fmtMes(mes)}: ${r.score}/100`);
  }
  $("#btn-registrar").addEventListener("click", () => registrarMes(false));

  $("#t-local").addEventListener("change", () => {
    const l = locales.find((x) => x.local.toLowerCase() === $("#t-local").value.trim().toLowerCase());
    if (l) { seleccionarLocal(l.local, false); renderTendencias(); } else toast("Local no encontrado");
  });
  $("#t-anio").addEventListener("change", renderTendencias);

  function aniosDisponibles() {
    const set = new Set([new Date().getFullYear()]);
    if (RES.mes) set.add(+RES.mes.slice(0, 4));
    if (usuario) Object.keys(localStorageSeguro()).forEach((k) => {
      if (k.startsWith(`matriz:${usuario.usuario}:hist:`)) Object.keys(store.get(k, {})).forEach((m) => set.add(+m.slice(0, 4)));
    });
    return [...set].sort((a, b) => b - a);
  }
  function localStorageSeguro() { try { return localStorage; } catch { return {}; } }

  function renderTendencias() {
    const sel = $("#t-anio"), anios = aniosDisponibles(), previo = +sel.value;
    sel.innerHTML = anios.map((a) => `<option>${a}</option>`).join("");
    sel.value = anios.includes(previo) ? previo : (RES.mes && anios.includes(+RES.mes.slice(0, 4)) ? +RES.mes.slice(0, 4) : anios[0]);
    const anio = +sel.value;
    renderMapa(anio);
    if (!localActual) {
      $("#t-kpis").innerHTML = ""; $("#t-chart").innerHTML = `<p class="muted">Elige un local para ver su evolución.</p>`;
      $("#t-tabla tbody").innerHTML = ""; return;
    }
    $("#t-local").value = localActual.local;
    $("#t-titulo").textContent = `Puntaje del check list por mes · ${localActual.local}`;
    const h = historial(localActual.local);
    const filas = MESES.map((_, i) => { const k = `${anio}-${String(i + 1).padStart(2, "0")}`; return { k, i, d: h[k] || {} }; });
    const conNota = filas.filter((f) => f.d.score != null);

    // KPIs
    const prom = conNota.length ? conNota.reduce((s, f) => s + f.d.score, 0) / conNota.length : null;
    const ult = conNota[conNota.length - 1], ant = conNota[conNota.length - 2];
    const mejor = conNota.reduce((a, f) => (!a || f.d.score > a.d.score ? f : a), null);
    const peor = conNota.reduce((a, f) => (!a || f.d.score < a.d.score ? f : a), null);
    const delta = ult && ant ? ult.d.score - ant.d.score : null;
    $("#t-kpis").innerHTML = [
      [ult ? `${ult.d.score}` : "—", ult ? `Último puntaje (${MESES_LARGO[ult.i]})` : "Último puntaje"],
      [delta == null ? "—" : (delta > 0 ? "▲ +" : delta < 0 ? "▼ " : "= ") + delta, ant ? `Cambio vs ${MESES_LARGO[ant.i]}` : "Cambio vs mes anterior"],
      [prom == null ? "—" : prom.toFixed(1), `Promedio ${anio}`],
      [mejor ? `${mejor.d.score} · ${MESES[mejor.i]}` : "—", "Mejor mes"],
      [peor ? `${peor.d.score} · ${MESES[peor.i]}` : "—", "Peor mes"],
      [`${conNota.length} / 12`, "Meses registrados"],
    ].map(([v, t]) => `<div><b>${esc(v)}</b><span>${esc(t)}</span></div>`).join("");

    renderGrafico(filas);

    // Tabla mensual (indicadores editables como en la hoja Novedades)
    const num = (v, suf = "") => (v == null || v === "" ? "—" : `${v}${suf}`);
    $("#t-tabla tbody").innerHTML = filas.map((f) => {
      const d = f.d, inp = (campo, step) => `<input type="number" step="${step}" data-mes="${f.k}" data-campo="${campo}" value="${d[campo] ?? ""}" aria-label="${campo} ${MESES_LARGO[f.i]}">`;
      return `<tr class="${d.score == null ? "vacio" : ""}"><td>${MESES_LARGO[f.i]}</td>
        <td class="num">${d.score == null ? "—" : `<span class="celda ${nivel(d.score)}">${d.score}</span>`}</td>
        <td class="num">${num(d.novedades)}</td><td class="num">${num(d.efectividad, "%")}</td><td class="num">${num(d.activosNov)}</td>
        <td>${esc(d.fuente || "")}</td>
        <td class="num">${inp("cxc", "0.01")}</td><td class="num">${inp("mermas", "0.01")}</td>
        <td class="num">${inp("lunch", "0.01")}</td><td class="num">${inp("costo", "0.01")}</td></tr>`;
    }).join("");
  }
  $("#t-tabla").addEventListener("change", (e) => {
    const i = e.target.closest("input[data-mes]"); if (!i || !localActual) return;
    const h = store.get(claveHist(localActual.local), {});
    h[i.dataset.mes] = Object.assign(h[i.dataset.mes] || {}, { [i.dataset.campo]: i.value === "" ? null : +i.value });
    store.set(claveHist(localActual.local), h);
    toast("Guardado");
  });

  function renderGrafico(filas) {
    const W = 760, H = 260, m = { t: 24, r: 56, b: 30, l: 36 };
    const datos = filas.filter((f) => f.d.score != null);
    const box = $("#t-chart");
    if (!datos.length) {
      box.innerHTML = `<svg viewBox="0 0 ${W} 120" role="img" aria-label="Sin datos"><text class="vacio" x="${W / 2}" y="64" text-anchor="middle">Sin puntajes registrados en este año. Usa «Registrar en tendencias» en la pestaña Informe.</text></svg>`;
      return;
    }
    const minV = Math.min(...datos.map((f) => f.d.score));
    const y0 = Math.max(0, Math.min(60, Math.floor((minV - 5) / 10) * 10)), y1 = 100;
    const x = (i) => m.l + (i * (W - m.l - m.r)) / 11;
    const y = (v) => m.t + ((y1 - v) * (H - m.t - m.b)) / (y1 - y0);
    let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Puntaje mensual del check list">`;
    for (let v = y0; v <= y1; v += 10) s += `<line class="grid" x1="${m.l}" x2="${W - m.r}" y1="${y(v)}" y2="${y(v)}"/><text class="eje" x="${m.l - 8}" y="${y(v) + 4}" text-anchor="end">${v}</text>`;
    [[META, "Meta 90"], [ALERTA, "Alerta 75"]].forEach(([v, t]) => { if (v >= y0) s += `<line class="meta" x1="${m.l}" x2="${W - m.r}" y1="${y(v)}" y2="${y(v)}"/><text class="eje" x="${W - m.r + 6}" y="${y(v) + 4}">${t}</text>`; });
    MESES.forEach((n, i) => (s += `<text class="eje" x="${x(i)}" y="${H - 8}" text-anchor="middle">${n}</text>`));
    // Línea por tramos de meses consecutivos (un mes sin dato corta la línea)
    let tramo = [];
    const cerrar = () => { if (tramo.length > 1) s += `<polyline class="linea" points="${tramo.join(" ")}"/>`; tramo = []; };
    filas.forEach((f) => { if (f.d.score == null) cerrar(); else tramo.push(`${x(f.i)},${y(f.d.score)}`); });
    cerrar();
    const ult = datos[datos.length - 1];
    datos.forEach((f) => (s += `<circle class="punto" cx="${x(f.i)}" cy="${y(f.d.score)}" r="4.5"/>`));
    s += `<text class="valor" x="${x(ult.i)}" y="${y(ult.d.score) - 10}" text-anchor="middle">${ult.d.score}</text>`;
    s += `<line class="cruz" id="t-cruz" x1="0" x2="0" y1="${m.t}" y2="${H - m.b}" visibility="hidden"/>`;
    filas.forEach((f) => (s += `<rect data-i="${f.i}" x="${x(f.i) - (W - m.l - m.r) / 22}" y="${m.t}" width="${(W - m.l - m.r) / 11}" height="${H - m.t - m.b}" fill="transparent"/>`));
    s += `</svg><div class="tip" id="t-tip" hidden></div>`;
    box.innerHTML = s;

    const svg = box.querySelector("svg"), tip = $("#t-tip"), cruz = $("#t-cruz");
    svg.addEventListener("mousemove", (e) => {
      const r = e.target.closest("rect[data-i]"); if (!r) return;
      const f = filas[+r.dataset.i], d = f.d;
      cruz.setAttribute("x1", x(f.i)); cruz.setAttribute("x2", x(f.i)); cruz.setAttribute("visibility", "visible");
      const fila = (k, v) => `<div class="fila"><span>${k}</span><span>${esc(v)}</span></div>`;
      tip.innerHTML = `<b>${MESES_LARGO[f.i]} ${f.k.slice(0, 4)}</b>` + (d.score == null ? `<div class="fila"><span>Sin registro</span></div>` :
        fila("Puntaje", `${d.score}/100`) + fila("Novedades", d.novedades ?? "—") + fila("Efectividad", d.efectividad != null ? d.efectividad + "%" : "—") + fila("Fuente", d.fuente || "—"));
      tip.hidden = false;
      const bw = box.clientWidth, px = (x(f.i) / W) * bw;
      tip.style.left = Math.min(Math.max(px + 12, 0), bw - tip.offsetWidth) + "px";
      tip.style.top = "8px";
      if (px + 12 + tip.offsetWidth > bw) tip.style.left = Math.max(0, px - tip.offsetWidth - 12) + "px";
    });
    svg.addEventListener("mouseleave", () => { tip.hidden = true; cruz.setAttribute("visibility", "hidden"); });
  }

  function renderMapa(anio) {
    const nombres = new Set();
    if (usuario.auditor) locales.filter((l) => l.auditor === usuario.auditor).forEach((l) => nombres.add(l.local));
    if (RES.mes.startsWith(String(anio))) RES.locales.forEach((l) => nombres.add(l.local));
    Object.keys(localStorageSeguro()).forEach((k) => {
      const p = `matriz:${usuario.usuario}:hist:`;
      if (k.startsWith(p) && Object.keys(store.get(k, {})).some((m) => m.startsWith(String(anio)))) nombres.add(k.slice(p.length));
    });
    const lista = [...nombres].sort((a, b) => a.localeCompare(b, "es"));
    $("#t-resumen-txt").textContent = lista.length
      ? `${lista.length} locales · ${anio} · verde ≥ ${META}, ámbar ≥ ${ALERTA}, rojo < ${ALERTA}. Clic en un local para ver su gráfico.`
      : "Aún no hay locales con puntajes registrados.";
    $("#t-mapa thead").innerHTML = `<tr><th>Local</th>${MESES.map((n) => `<th>${n}</th>`).join("")}</tr>`;
    $("#t-mapa tbody").innerHTML = lista.map((n) => {
      const h = historial(n);
      return `<tr data-local="${esc(n)}"><td>${esc(n)}</td>${MESES.map((_, i) => {
        const d = h[`${anio}-${String(i + 1).padStart(2, "0")}`];
        return d && d.score != null ? `<td><span class="celda ${nivel(d.score)}" title="${esc(MESES_LARGO[i])}: ${d.score}/100 (${esc(d.fuente || "")})">${d.score}</span></td>` : `<td><span class="celda na">·</span></td>`;
      }).join("")}</tr>`;
    }).join("");
  }
  $("#t-mapa tbody").addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-local]"); if (!tr) return;
    seleccionarLocal(tr.dataset.local, false); renderTendencias(); window.scrollTo({ top: 0, behavior: "smooth" });
  });

  // ---------- Correo Outlook ----------
  // Tres correos distintos: Informe (check list), Activos y Producto. Cada asunto de la hoja «Asunto» pertenece a uno.
  const categoriaDe = (enfoque) => (/baja/i.test(enfoque) ? "baja" : /activo/i.test(enfoque) ? "activos" : /producto/i.test(enfoque) ? "producto" : "informe");
  const catCorreo = () => $('input[name="m-cat"]:checked').value;
  const DESC_CAT = {
    informe: "Check list de procesos: resultado global, procesos que no cumplen con su observación y observaciones generales.",
    activos: "Inventario de activos fijos: verificados, novedades, equipos modificados, perdidos y nuevos. «Borrador .eml» incluye las fotos de los activos con novedad.",
    producto: "Inventario o arqueo de producto: costos ideal y real, efectividad y sobrantes/faltantes de mayor incidencia. «Borrador .eml» incluye el reporte gráfico.",
    baja: "Solicitud de autorización para dar de baja los activos pendientes de la pestaña «Baja de activos». «Borrador .eml» incluye las fotos de los equipos.",
  };
  function llenarAsuntos() {
    const cat = catCorreo(), previo = $("#m-enfoque").value;
    const opciones = M.asuntos.map((a) => a.enfoque).filter((e) => categoriaDe(e) === cat);
    $("#m-enfoque").innerHTML = opciones.map((e) => `<option>${esc(e)}</option>`).join("");
    if (opciones.includes(previo)) $("#m-enfoque").value = previo;
    $("#m-cat-desc").textContent = DESC_CAT[cat];
    $("#m-saludo").closest("label").hidden = cat === "baja"; // la baja usa el tratamiento de la pestaña «Baja de activos»
  }
  function elegirCorreo(enfoque) {
    $(`input[name="m-cat"][value="${categoriaDe(enfoque)}"]`).checked = true;
    llenarAsuntos(); $("#m-enfoque").value = enfoque;
    delete $("#m-cuerpo").dataset.editado;
  }
  $$('input[name="m-cat"]').forEach((r) => r.addEventListener("change", () => { llenarAsuntos(); delete $("#m-cuerpo").dataset.editado; prepararCorreo(true); }));
  llenarAsuntos();
  const fechaLarga = (iso) => { const d = new Date(iso + "T12:00:00"); return isNaN(d) ? iso : d.toLocaleDateString("es-EC", { day: "2-digit", month: "long", year: "numeric" }); };

  function asuntoPara(enfoque) {
    const a = M.asuntos.find((x) => x.enfoque === enfoque);
    let base = a ? a.asunto.replace(/\s*-\s*k004 cci\s*$/i, "") : enfoque;
    if (/producto/i.test(enfoque) && localActual) base = producto().tipo;
    if (ENF_ZONA.includes(enfoque)) return `${base} - ${zonaCu() || (localActual && localActual.auditor) || ""}`; // reporte de toda la zona
    return localActual ? `${base} - ${localActual.local}` : base;
  }
  const fechaCorta = (iso) => { const d = new Date(iso + "T12:00:00"); return isNaN(d) ? iso : d.toLocaleDateString("es-EC", { day: "numeric", month: "long" }); };
  const unirY = (xs) => (xs.length > 1 ? `${xs.slice(0, -1).join(", ")} y ${xs[xs.length - 1]}` : xs[0] || "");
  const FIRMA = "Saludos cordiales,";

  // Análisis de los dos sobrantes y faltantes de mayor monto (ordenados aunque la tabla se haya editado)
  function analisisTop(p) {
    const venta = parseFloat(p.venta) || 0;
    const item = (x) => {
      const t = parseFloat(x.total) || 0, c = parseFloat(x.cantidad) || 0;
      const u = (x.unidad || "").toLowerCase(), plural = c === 1 || !u || /s$/.test(u) ? u : /[aeiou]$/.test(u) ? u + "s" : u + "es";
      const cant = c ? `${c.toLocaleString("es-EC", { maximumFractionDigits: 2 })} ${plural}`.trim() + ", " : "";
      return `${x.articulo || "(sin nombre)"} (${cant}${fmtUSD(t)}${venta ? ` – ${fmtPct(t / venta)} de la venta` : ""})`;
    };
    const frase = (k, nombre) => {
      const filas = p[k].filter(filaLlena).sort((a, b) => (parseFloat(b.total) || 0) - (parseFloat(a.total) || 0));
      if (!filas.length) return null;
      const tot = sumaLista(filas), dos = filas.slice(0, 2), sumDos = sumaLista(dos);
      const peso = tot && filas.length > 2 ? `; juntos representan el ${fmtPct(sumDos / tot)} del ${nombre} TOP (${fmtUSD(tot)})` : "";
      return `${nombre === "faltante" ? "Faltantes" : "Sobrantes"} de mayor incidencia: ${unirY(dos.map(item))}${peso}.`;
    };
    return [frase("faltantes", "faltante"), frase("sobrantes", "sobrante")].filter(Boolean);
  }

  // Cada correo sigue el formato del área: saludo, fecha de corte, qué se ejecutó y una viñeta por hallazgo.
  // Solo cambia el contenido: Informe (check list), Activos o Producto.
  function generarCuerpo(enfoque) {
    if (!localActual) return "Selecciona un local para generar el mensaje.";
    const l = localActual, nombre = (l.jefe || "").trim().split(/\s+/)[0] || "", hora = new Date().getHours();
    if (categoriaDe(enfoque) === "baja") return cuerpoBaja(); // la solicitud de baja va dirigida a quien autoriza, con su propio saludo
    if (ENF_ZONA.includes(enfoque)) return cuerpoZona(enfoque);
    const L = [`${$("#m-saludo").value}${nombre ? " " + nombre : ""}, ${hora < 12 ? "buenos días" : hora < 19 ? "buenas tardes" : "buenas noches"}`, ""];
    const B = (t) => L.push(`• ${t}`);
    ({ informe: cuerpoInforme, activos: cuerpoActivos, producto: cuerpoProducto })[categoriaDe(enfoque)](L, B);
    L.push("", FIRMA, usuario.nombre, "Auditoría · Grupo KFC Ecuador");
    return L.join("\n");
  }
  // Reportes de toda la zona: cumplimiento del inventario de activos y análisis de cuentas por cobrar
  function cuerpoZona(enfoque) {
    if (!$("#cu-zona").options.length) prepararCumplimiento();
    const hora = new Date().getHours(), mes = `${nombreMes(mesCu())} ${mesCu().slice(0, 4)}`;
    const L = [`Estimados, ${hora < 12 ? "buenos días" : hora < 19 ? "buenas tardes" : "buenas noches"}`, ""];
    const B = (t) => L.push(`• ${t}`);
    if (enfoque === "Cumplimiento Inventario Activos") {
      const d = datosCumplimiento(), no = d.filas.filter((f) => f.estatus !== "OK").map((f) => f.local);
      L.push(`Se comparte el cumplimiento del Inventario Total de Activos del mes de ${mes} en los locales de la zona ${zonaCu()}.`, "");
      B(`Cumplieron ${d.ok} de ${d.filas.length} locales (${d.pct} %).`);
      if (no.length) B(`No realizaron el inventario: ${unirY(no)}.`); else B("Todos los locales realizaron el inventario de activos. ¡Felicitaciones!");
    } else {
      const d = datosCxc(), [, t1, t2] = d.totales, varia = t1 ? (t2 - t1) / t1 : NaN;
      const top = d.filas.filter((f) => f.v[2] > 0).slice(0, 3);
      L.push(`Se comparte el análisis de cuentas por cobrar de ${unirY(d.ms.map(nombreMes))} de los locales de la zona ${zonaCu()}.`, "");
      B(`Total por cobrar en ${nombreMes(d.ms[2])}: ${fmtUSD(t2)}${isNaN(varia) ? "" : ` (${varia >= 0 ? "aumenta" : "disminuye"} ${fmtPct(Math.abs(varia))} frente a ${nombreMes(d.ms[1])}, ${fmtUSD(t1)})`}.`);
      if (top.length) B(`Locales con mayor saldo: ${unirY(top.map((f) => `${f.local} (${fmtUSD(f.v[2])})`))}.`);
      B("Seguimiento y gestión de cobro de los saldos pendientes.");
    }
    B("Se adjunta el cuadro con el detalle por local.");
    L.push("", FIRMA, usuario.nombre, "Auditoría · Grupo KFC Ecuador");
    return L.join("\n");
  }
  const esCorreoZona = () => ENF_ZONA.includes($("#m-enfoque").value);
  // 🗑️ Baja: solicitud de autorización a quien aprueba, con los activos pendientes de baja
  function cuerpoBaja() {
    const pend = bajasLocal().filter(bajaPendiente);
    if (!pend.length) return "No hay activos pendientes de baja en este local. Agrégalos en la pestaña «Baja de activos».";
    const d = pend[pend.length - 1]; // datos de la última baja registrada (quién autoriza, certifica y jefe)
    const total = pend.reduce((s, b) => s + (parseFloat(b.valor) || 0), 0);
    const L = [`${d.trat || "Estimado Señor"} @${d.autoriza},`, ""];
    L.push("Su autorización para dar de baja los siguientes activos fijos que se encuentran deteriorados o en desuso, los mismos que estorban en el local.", "");
    pend.forEach((b) => L.push(`• ${b.codigo} ${b.nombre}${b.marca ? " / " + b.marca : ""}${b.modelo ? " " + b.modelo : ""}${b.serie ? " / S: " + b.serie : ""}`
      + ` — ${b.motivo}; valor en libros ${b.valor === "" || b.valor == null ? "sin dato" : fmtUSD(+b.valor)}${b.obs ? `; ${b.obs}` : ""}`));
    L.push("", `Esta información ha sido verificada y certificada por el Subgerente ${d.gerente} y comunicado al Jefe de área @${d.jefe}. `
      + (total === 0 ? "Los equipos se encuentran con VALOR CERO EN LIBROS." : `Los equipos suman un valor en libros de ${fmtUSD(total)}.`) + " Se adjunta imagen de equipos.");
    L.push("", FIRMA, usuario.nombre, "Auditoría · Grupo KFC Ecuador");
    return L.join("\n");
  }
  // Fotos de los activos pendientes de baja para el borrador (máximo 20)
  async function fotosBajaCorreo() {
    const st = activosLocal(), imgs = [];
    for (const b of bajasLocal().filter(bajaPendiente)) {
      const a = st.activos.find((x) => x.id === b.activoId); if (!a) continue;
      for (const id of a.fotos || []) {
        if (imgs.length >= 20) break;
        const f = await fotosDB.get(id).catch(() => null);
        if (f) imgs.push({ cid: `baja${imgs.length}@matriz`, nombre: `baja${imgs.length + 1}.jpg`, tipo: "image/jpeg", b64: await blobB64(f.blob), activo: a, baja: b });
      }
    }
    return imgs;
  }
  // 📋 Informe: check list de procesos
  function cuerpoInforme(L, B) {
    const st = estadoLocal(localActual.local), r = calcular();
    const tipo = /visita/i.test($("#sel-informe").value) ? "visita parcial" : "chequeo total";
    const evaluado = Object.keys(r.evaluacion).length > 0 || r.marcados.length > 0;
    L.push(`Con fecha de corte ${fechaCorta($("#sel-fecha").value || hoy())} se ha ejecutado el CHECKLIST de los esquemas de control (${tipo}) en el local ${localActual.local}.`, "");
    // La visita parcial no lleva calificación: solo las novedades y los puntos perdidos
    if (!evaluado) B("El check list aún no tiene procesos evaluados.");
    else if (tipo === "chequeo total") B(`El resultado global de la Evaluación es de ${r.score} %.`);
    else if (r.items.length) B(`Puntos perdidos en la visita: -${100 - r.score}.`);
    r.items.forEach((c) => B(c.obs ? `${c.area}: ${c.obs}` : c.activos ? `${c.desc} (${c.cant} activo${c.cant > 1 ? "s" : ""}: ${unirY(c.activos.map((a) => a.nombre))}; -${c.puntosTot} puntos)` : c.desc || c.area));
    if (r.items.length) B("Para su análisis adjunto el cuadro de novedades del check list.");
    // Observaciones generales: una viñeta por línea, tal como las escribe el auditor
    (st.observaciones || "").split("\n").map((s) => s.trim().replace(/^[•*\-]\s*/, "")).filter(Boolean).forEach(B);
    if (r.cumplen.length) B(`${TXT_CUMPLE} en: ${r.cumplen.map((c) => c.area).join(", ")}.`);
  }
  // 🏷️ Activos: inventario de activos fijos
  function cuerpoActivos(L, B) {
    const st = activosLocal(), lista = st.activos;
    const base = lista.filter(enUso), verificados = base.filter((a) => a.novedad), pendientes = base.filter((a) => !a.novedad);
    const conNov = lista.filter(conNovedadActivo), sinNov = base.filter((a) => a.novedad === OK_ACT);
    const fecha = (st.archivoActivos && isoDe(st.archivoActivos.fecha)) || st.archivoActivos?.fecha;
    L.push(`Con fecha de corte ${fechaCorta(/^\d{4}-\d{2}-\d{2}$/.test(fecha || "") ? fecha : $("#sel-fecha").value || hoy())} se ha ejecutado el Inventario total de activos fijos en el local ${localActual.local}.`, "");
    if (!lista.length) { B("Aún no se ha cargado el listado de activos del SIR."); return; }
    B(`Activos registrados en el SIR: ${st.sir || 0}; verificados: ${verificados.length}${pendientes.length ? `; pendientes de verificar: ${pendientes.length}` : ""}.`);
    B(conNov.length ? `Activos sin novedad: ${sinNov.length}; con novedad: ${conNov.length}.` : `Todos los activos verificados están sin novedad (${sinNov.length}).`);
    // Por tipo de novedad: «Activo perdido: CAJA y TABLET»
    const porTipo = conNov.reduce((m, a) => ((m[a.novedad] = m[a.novedad] || []).push(a.nombre), m), {});
    Object.entries(porTipo).forEach(([n, xs]) => B(`${n}: ${unirY(xs)}.`));
    const dc = datosCuadroActivos();
    B(dc.total < 0 ? `Puntos perdidos en activos: ${dc.total} (se descuentan en la evaluación del check list).` : "Felicitaciones: no hay puntos perdidos en activos.");
    const app = lista.filter((a) => a.app && a.origen !== "app").length, noApp = base.filter((a) => a.origen === "SIR" && !a.app).length;
    if (st.respaldoApp) B(`Cruce con la app de inventario: ${app} activos encontrados${noApp ? `; ${noApp} del SIR no están registrados en la app` : ""}.`);
    if (conNov.length) {
      L.push("", "Detalle de activos con novedad:");
      conNov.forEach((a) => L.push(`  - ${a.codigo} ${a.nombre}${a.marca ? " / " + a.marca : ""}${a.modelo ? " " + a.modelo : ""}${a.serie ? " / S: " + a.serie : ""} → ${a.novedad}`
        + `${descCambios(a) ? ` (${descCambios(a)})` : ""}${a.creadoAuditor ? " (activo nuevo, creado por auditoría)" : ""}${a.obs ? ` — ${a.obs}` : ""}`
        + `${a.fotos && a.fotos.length ? ` (${a.fotos.length} foto${a.fotos.length > 1 ? "s" : ""})` : ""}`));
    }
  }
  // 📦 Producto: inventario total o arqueo
  function cuerpoProducto(L, B) {
    const p = producto(), co = p.costos || {}, r = calcular(), venta = parseFloat(p.venta) || 0;
    const tipo = p.tipo === "Arqueo de Producto" ? "Arqueo de producto" : "Inventario total de producto";
    L.push(`Con fecha de corte ${fechaCorta(co.audFecha || p.fecha || $("#sel-fecha").value || hoy())} se ha ejecutado el ${tipo} en el local ${localActual.local}.`, "");
    const hayTop = p.sobrantes.some(filaLlena) || p.faltantes.some(filaLlena);
    const dCosto = difCosto(co.audIdeal, co.audReal) || difCosto(co.perIdeal, co.perReal);
    if (dCosto) {
      const extra = [["merma", co.merma], ["lunch", co.lunch]].filter(([, v]) => v !== "" && v != null).map(([n, v]) => `${n} ${fmtP(v)}`);
      B(`A la fecha dejamos certificado los costos: IDEAL ${fmtP(dCosto.ideal)} - REAL ${fmtP(dCosto.real)}, ${dCosto.v ? `${dCosto.tipo} de ${fmtP(Math.abs(dCosto.v))}` : "sin diferencia"}.${extra.length ? ` (${extra.join(", ")}).` : ""}`);
    }
    if (venta || hayTop) {
      const sob = sumaLista(p.sobrantes), fal = sumaLista(p.faltantes), pv = (v) => (venta ? ` (${fmtPct(v / venta)} de la venta)` : "");
      B(`${venta ? `Venta neta ${fmtUSD(venta)}: ` : ""}faltante TOP ${fmtUSD(fal)}${pv(fal)} y sobrante TOP ${fmtUSD(sob)}${pv(sob)}; diferencia neta ${fmtUSD(sob - fal)}.`);
    }
    if (!isNaN(r.ef)) B(`Efectividad de inventarios totales: ${fmtPct(r.ef / 100)} (${r.efItems} ítems, ${r.efError} error(es)).`);
    analisisTop(p).forEach(B);
    if (hayTop) B("Para su análisis adjunto al email sobrantes y faltantes de mayor incidencia.");
    if (!dCosto && !hayTop && !venta) B("Aún no se han ingresado costos ni sobrantes/faltantes en la pestaña Producto.");
  }
  // El saludo (Estimada / Estimado …) lo elige el auditor y se recuerda por jefe de área
  const claveSaludo = () => `matriz:${usuario.usuario}:saludo`;
  $("#m-saludo").addEventListener("change", () => {
    if (localActual && localActual.jefe) { const g = store.get(claveSaludo(), {}); g[localActual.jefe] = $("#m-saludo").value; store.set(claveSaludo(), g); }
    delete $("#m-cuerpo").dataset.editado; prepararCorreo(true);
  });
  function prepararCorreo(forzar) {
    const enf = $("#m-enfoque").value;
    if (localActual) $("#m-saludo").value = store.get(claveSaludo(), {})[localActual.jefe] || "Estimado/a";
    $("#m-asunto").value = asuntoPara(enf);
    if (forzar || !$("#m-cuerpo").dataset.editado) $("#m-cuerpo").value = generarCuerpo(enf);
    const guardados = store.get(`matriz:${usuario.usuario}:para`, {});
    if (localActual && !$("#m-para").value) $("#m-para").value = guardados[localActual.local] || "";
    renderPreview();
  }
  function renderPreview() {
    $("#m-preview").innerHTML = `<div class="p-head"><b>Para:</b> ${esc($("#m-para").value) || "<span class='muted'>(sin destinatario)</span>"}<br>
      ${$("#m-cc").value ? `<b>CC:</b> ${esc($("#m-cc").value)}<br>` : ""}<b>Asunto:</b> ${esc($("#m-asunto").value)}</div>${esc($("#m-cuerpo").value)}`;
    // Reportes de zona: el cuadro como imagen (va en el .eml)
    if (esCorreoZona()) {
      const cv = $("#m-enfoque").value === ENF_ZONA[0] ? imagenCumplimiento() : imagenCxc();
      $("#m-preview").insertAdjacentHTML("beforeend", `<div class="preview-reporte preview-blanco"><p class="muted small">Cuadro de la zona: va incluido en «Borrador .eml». Con «Abrir en Outlook» pégalo con
        <button type="button" class="ghost sm" id="m-copiar-zona">Copiar cuadro</button> y Ctrl+V.</p><img alt="Cuadro de la zona" src="${cv.toDataURL("image/png")}"></div>`);
      $("#m-copiar-zona").addEventListener("click", () => ($("#m-enfoque").value === ENF_ZONA[0] ? $("#cu-copiar") : $("#cx-copiar")).click());
    }
    // Informe: el cuadro del check list va dentro del borrador .eml; aquí se muestra cómo queda
    if (catCorreo() === "informe" && !esCorreoZona() && localActual && calcular().items.length) {
      $("#m-preview").insertAdjacentHTML("beforeend", `<div class="preview-reporte preview-blanco"><p class="muted small">Cuadro de novedades: va incluido en «Borrador .eml». Con «Abrir en Outlook» pégalo con
        <button type="button" class="ghost sm" id="m-copiar-res">Copiar cuadro</button> y Ctrl+V.</p><img alt="Cuadro de resultado del check list" src="${dibujarResultado().toDataURL("image/png")}"></div>`);
      $("#m-copiar-res").addEventListener("click", () => $("#res-copiar").click());
    }
    if (conReporte()) {
      const cv = dibujarReporte("local");
      $("#m-preview").insertAdjacentHTML("beforeend", `<div class="preview-reporte"><p class="muted small">Reporte gráfico: va incluido en «Borrador .eml». Con «Abrir en Outlook» o «Outlook Web» pégalo tú con
        <button type="button" class="ghost sm" id="m-copiar-reporte">Copiar reporte</button> y Ctrl+V.</p><img alt="Reporte de producto" src="${cv.toDataURL("image/png")}"></div>`);
      $("#m-copiar-reporte").addEventListener("click", () => $("#r-copiar-img").click());
    }
    const largo = mailtoUrl().length;
    $("#m-aviso").textContent = largo > 1900
      ? "El mensaje es largo: Outlook de escritorio puede recortarlo con «Abrir en Outlook». Usa «Borrador .eml» para el texto completo."
      : "«Abrir en Outlook» abre un mensaje nuevo en tu Outlook de escritorio listo para revisar y enviar.";
  }
  $("#m-enfoque").addEventListener("change", () => prepararCorreo(true));
  $("#m-cuerpo").addEventListener("input", () => { $("#m-cuerpo").dataset.editado = "1"; renderPreview(); });
  ["#m-para", "#m-cc", "#m-asunto"].forEach((s) => $(s).addEventListener("input", renderPreview));
  $("#m-para").addEventListener("change", () => {
    if (!localActual) return;
    const g = store.get(`matriz:${usuario.usuario}:para`, {}); g[localActual.local] = $("#m-para").value;
    store.set(`matriz:${usuario.usuario}:para`, g);
  });
  $("#btn-regenerar").addEventListener("click", () => { delete $("#m-cuerpo").dataset.editado; prepararCorreo(true); });

  const listaCorreos = (s) => s.split(/[;,]\s*/).map((x) => x.trim()).filter(Boolean).join(";");
  function mailtoUrl() {
    const p = [];
    if ($("#m-cc").value) p.push("cc=" + encodeURIComponent(listaCorreos($("#m-cc").value)));
    p.push("subject=" + encodeURIComponent($("#m-asunto").value));
    p.push("body=" + encodeURIComponent($("#m-cuerpo").value.replace(/\n/g, "\r\n")));
    return `mailto:${encodeURIComponent(listaCorreos($("#m-para").value)).replace(/%3B/g, ";").replace(/%40/g, "@")}?${p.join("&")}`;
  }
  function validar() {
    if (!localActual) { toast("Selecciona un local primero"); return false; }
    if (!$("#m-para").value.trim()) { toast("Agrega al menos un destinatario en «Para»"); $("#m-para").focus(); return false; }
    if (faltanObservaciones()) return false;
    registrarMes(true); // el informe enviado queda en Tendencias
    return true;
  }
  $("#btn-outlook").addEventListener("click", () => { if (validar()) { window.location.href = mailtoUrl(); toast("Abriendo Outlook…"); } });
  $("#btn-owa").addEventListener("click", () => {
    if (!validar()) return;
    const q = new URLSearchParams({ to: listaCorreos($("#m-para").value), subject: $("#m-asunto").value, body: $("#m-cuerpo").value });
    if ($("#m-cc").value) q.set("cc", listaCorreos($("#m-cc").value));
    window.open("https://outlook.office.com/mail/deeplink/compose?" + q.toString().replace(/\+/g, "%20"), "_blank", "noopener");
  });
  $("#btn-copiar").addEventListener("click", async () => {
    const txt = `Asunto: ${$("#m-asunto").value}\n\n${$("#m-cuerpo").value}`;
    try { await navigator.clipboard.writeText(txt); toast("Mensaje copiado"); }
    catch { const t = $("#m-cuerpo"); t.select(); document.execCommand("copy"); toast("Mensaje copiado"); }
  });
  const blobB64 = (blob) => new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(String(r.result).split(",")[1]); r.readAsDataURL(blob); });
  // Fotos de los activos con novedad, para incrustarlas en el borrador (máximo 20 para no hacer el correo enorme)
  async function fotosParaCorreo() {
    const conFotos = activosLocal().activos.filter((a) => conNovedadActivo(a) && a.fotos.length);
    const imgs = [];
    for (const a of conFotos) for (const id of a.fotos) {
      if (imgs.length >= 20) break;
      const f = await fotosDB.get(id).catch(() => null);
      if (f) imgs.push({ cid: `foto${imgs.length}@matriz`, nombre: `foto${imgs.length + 1}.jpg`, tipo: "image/jpeg", b64: await blobB64(f.blob), activo: a });
    }
    return imgs;
  }
  $("#btn-eml").addEventListener("click", async () => {
    if (!validar()) return;
    const b64 = (s) => { const b = new TextEncoder().encode(s); let t = ""; for (let i = 0; i < b.length; i += 8192) t += String.fromCharCode(...b.subarray(i, i + 8192)); return btoa(t); };
    const lineas = (s) => s.replace(/.{76}/g, "$&\r\n");
    const aHtml = (t) => esc(t).replace(/\n/g, "<br>");
    // Reporte de producto y fotos de activos van después de las viñetas, antes de la firma
    const ls = $("#m-cuerpo").value.split("\n"), iFirma = ls.indexOf(FIRMA), marca = "@@ANEXOS@@";
    if (iFirma >= 0) ls.splice(iFirma, 0, marca, ""); else ls.push("", marca);
    const imagenes = [];
    let anexos = "";
    if (conReporte()) {
      anexos += htmlReporte("cid:pastel@matriz", "local");
      imagenes.push({ cid: "pastel@matriz", nombre: "pastel.png", tipo: "image/png", b64: pastelPNG("local").split(",")[1] });
    }
    if (esCorreoZona()) { // reporte de la zona: solo su cuadro como imagen
      const cv = $("#m-enfoque").value === ENF_ZONA[0] ? imagenCumplimiento() : imagenCxc();
      anexos += `<p style="margin:14px 0 6px"><img src="cid:zona@matriz" width="760" alt="Cuadro de la zona" style="display:block;max-width:100%"></p>`;
      imagenes.push({ cid: "zona@matriz", nombre: "cuadro.png", tipo: "image/png", b64: cv.toDataURL("image/png").split(",")[1] });
    }
    if (catCorreo() === "informe" && !esCorreoZona() && localActual && calcular().items.length) { // el cuadro del check list como imagen incrustada
      anexos += `<p style="margin:14px 0 6px"><img src="cid:resultado@matriz" width="760" alt="Cuadro de resultado del check list" style="display:block;max-width:100%"></p>`;
      imagenes.push({ cid: "resultado@matriz", nombre: "resultado.png", tipo: "image/png", b64: dibujarResultado().toDataURL("image/png").split(",")[1] });
    }
    if (catCorreo() === "baja") { // imágenes de los equipos a dar de baja
      const fotos = await fotosBajaCorreo();
      if (fotos.length) {
        anexos += `<p style="margin:16px 0 8px;font-size:12pt;font-weight:bold">Imágenes de los equipos</p>` + fotos.map((f) =>
          `<div style="display:inline-block;vertical-align:top;width:300px;margin:0 12px 12px 0;font-family:Segoe UI,Arial,sans-serif">
            <img src="cid:${f.cid}" width="300" style="display:block;border:1px solid #dfe3e8" alt="${esc(f.activo.nombre)}">
            <div style="font-size:9pt;color:#1b1f24;margin-top:4px"><b>${esc(f.activo.nombre)}</b> · ${esc(f.activo.codigo)}${f.activo.serie ? ` · S: ${esc(f.activo.serie)}` : ""}<br><span style="color:#646d78">${esc(f.baja.motivo)}</span></div></div>`).join("");
        imagenes.push(...fotos);
      }
    }
    if (catCorreo() === "activos" && !esCorreoZona()) {
      anexos += htmlCuadroActivos(true); // el cuadro de novedades y puntos, como en la matriz
      const fotos = await fotosParaCorreo();
      if (fotos.length) {
        anexos += `<p style="margin:16px 0 8px;font-size:12pt;font-weight:bold">Registro fotográfico de activos con novedad</p>` + fotos.map((f) =>
          `<div style="display:inline-block;vertical-align:top;width:300px;margin:0 12px 12px 0;font-family:Segoe UI,Arial,sans-serif">
            <img src="cid:${f.cid}" width="300" style="display:block;border:1px solid #dfe3e8" alt="${esc(f.activo.nombre)}">
            <div style="font-size:9pt;color:#1b1f24;margin-top:4px"><b>${esc(f.activo.nombre)}</b> · ${esc(f.activo.codigo)}${f.activo.serie ? ` · S: ${esc(f.activo.serie)}` : ""}<br><span style="color:#b42318">${esc(f.activo.novedad)}</span></div></div>`).join("");
        imagenes.push(...fotos);
      }
    }
    const cuerpo = aHtml(ls.join("\n")).replace(anexos ? marca : `${marca}<br>`, anexos);
    const html = `<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:11pt">${cuerpo}</body></html>`;
    const cab = [
      `To: ${listaCorreos($("#m-para").value).replace(/;/g, ", ")}`,
      $("#m-cc").value ? `Cc: ${listaCorreos($("#m-cc").value).replace(/;/g, ", ")}` : null,
      `Subject: =?UTF-8?B?${b64($("#m-asunto").value)}?=`, "X-Unsent: 1", "MIME-Version: 1.0",
    ].filter((x) => x !== null);
    const limite = "----=_matriz_" + Date.now();
    const eml = (imagenes.length
      ? [...cab, `Content-Type: multipart/related; boundary="${limite}"; type="text/html"`, "",
        `--${limite}`, "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: base64", "", lineas(b64(html)),
        ...imagenes.flatMap((im) => [`--${limite}`, `Content-Type: ${im.tipo}; name="${im.nombre}"`, "Content-Transfer-Encoding: base64", `Content-ID: <${im.cid}>`,
          `Content-Disposition: inline; filename="${im.nombre}"`, "", lineas(im.b64)]),
        `--${limite}--`, ""]
      : [...cab, "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: base64", "", lineas(b64(html))]).join("\r\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([eml], { type: "message/rfc822" }));
    a.download = `${$("#m-asunto").value.replace(/[\\/:*?"<>|]/g, "_")}.eml`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    toast("Borrador descargado: ábrelo y se abrirá en Outlook");
  });

  // ---------- Manuales de auditoría ----------
  // Todos los usuarios ven los manuales publicados en la carpeta de la app (manuales/manuales.js + los GIF).
  // Solo el admin sube, edita y publica. Mientras edita, los cambios viven en este navegador (IndexedDB «matriz-manuales»)
  // y «Publicar para todos» los escribe en la carpeta «manuales» de la app para que cualquier equipo con esa carpeta los vea.
  const esAdmin = () => !!usuario && usuario.rol === "admin";
  const manualesDB = (() => {
    let abierta;
    const abrir = () => abierta || (abierta = new Promise((ok, mal) => {
      const r = indexedDB.open("matriz-manuales", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("manuales", { keyPath: "id" });
      r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error);
    }));
    const tx = async (modo, fn) => { const db = await abrir(); return new Promise((ok, mal) => { const t = db.transaction("manuales", modo), req = fn(t.objectStore("manuales")); t.oncomplete = () => ok(req.result); t.onerror = () => mal(t.error); }); };
    return { all: () => tx("readonly", (s) => s.getAll()), put: (m) => tx("readwrite", (s) => s.put(m)), del: (id) => tx("readwrite", (s) => s.delete(id)) };
  })();
  const CLAVE_BORRADOS = "matriz:manuales:borrados"; // publicados que el admin quitó y aún no republica
  let mnUrls = [], mnLista = [];
  const publicados = () => (window.MANUALES || []).map((m) => ({ ...m, publicado: true }));
  async function cargarManuales() {
    const borrados = store.get(CLAVE_BORRADOS, []);
    const locales = await manualesDB.all().catch(() => []);
    if (!esAdmin()) { // auditores: lo publicado en la carpeta + lo que el admin subió en este mismo equipo (solo lectura)
      const ids = new Set(locales.map((l) => l.id));
      return locales.concat(publicados().filter((p) => !ids.has(p.id) && !borrados.includes(p.id)));
    }
    for (const p of publicados()) { // lo publicado entra a la lista del admin para poder editarlo
      if (!borrados.includes(p.id) && !locales.some((l) => l.id === p.id)) { await manualesDB.put(p); locales.push(p); }
    }
    return locales;
  }
  const ordenar = (xs) => xs.slice().sort((a, b) => (a.seccion || "").localeCompare(b.seccion || "", "es") || (a.orden || 0) - (b.orden || 0) || a.titulo.localeCompare(b.titulo, "es"));
  async function renderManuales() {
    mnUrls.forEach(URL.revokeObjectURL); mnUrls = [];
    const admin = esAdmin();
    mnLista = ordenar(await cargarManuales());
    $("#mn-admin").hidden = !admin;
    const pend = mnLista.filter((m) => m.pendiente || !m.publicado).length + store.get(CLAVE_BORRADOS, []).length;
    $("#mn-aviso").hidden = !admin || !pend;
    $("#mn-aviso").textContent = `${pend} cambio(s) sin publicar: solo se ven en este equipo hasta que pulses «Publicar para todos».`;
    $("#mn-info").textContent = admin
      ? "Sube GIF (o una carpeta: cada subcarpeta es una sección), edita títulos y publica. Los auditores solo pueden verlos."
      : mnLista.length ? `${mnLista.length} manual(es) · clic en un GIF para verlo en grande y hacer zoom.` : "Aún no hay manuales publicados.";
    const q = $("#mn-buscar").value.trim().toLowerCase();
    const vista = mnLista.filter((m) => !q || [m.titulo, m.seccion, m.descripcion].join(" ").toLowerCase().includes(q));
    const secciones = vista.reduce((g, m) => ((g[m.seccion || "General"] = g[m.seccion || "General"] || []).push(m), g), {});
    const src = (m) => { if (m.blob) { const u = URL.createObjectURL(m.blob); mnUrls.push(u); return u; } return m.archivo; };
    $("#mn-secciones").innerHTML = Object.keys(secciones).length ? Object.entries(secciones).map(([sec, xs]) => `<div class="mn-sec">
      <h3>${esc(sec)} <small>${xs.length} manual(es)</small></h3>
      <div class="mn-grid">${xs.map((m) => `<figure class="mn-card" data-mn="${m.id}" style="margin:0">
        <div class="mn-img" title="Ver en grande"><img src="${esc(src(m))}" alt="${esc(m.titulo)}" loading="lazy"></div>
        <figcaption class="mn-txt"><b>${esc(m.titulo)}</b>${m.descripcion ? `<span>${esc(m.descripcion)}</span>` : ""}
          ${admin && (m.pendiente || !m.publicado) ? `<span class="mn-badge">● Sin publicar</span>` : ""}</figcaption>
        <div class="mn-share">
          <button type="button" class="ghost sm" data-mn-descargar title="Descargar el manual">⬇ Descargar</button>
        </div>
        ${admin ? `<div class="mn-edit">
          <input data-mn-campo="titulo" value="${esc(m.titulo)}" placeholder="Título" aria-label="Título">
          <input data-mn-campo="seccion" value="${esc(m.seccion || "")}" placeholder="Sección" aria-label="Sección" list="mn-lista-sec">
          <textarea data-mn-campo="descripcion" rows="2" placeholder="Descripción (opcional)" aria-label="Descripción">${esc(m.descripcion || "")}</textarea>
          <div class="fila"><span><button type="button" class="ghost sm" data-mn-mover="-1" title="Subir">↑</button> <button type="button" class="ghost sm" data-mn-mover="1" title="Bajar">↓</button></span>
            <button type="button" class="ghost sm" data-mn-quitar title="Eliminar manual">✕ Eliminar</button></div></div>` : ""}
      </figure>`).join("")}</div></div>`).join("") + (admin ? `<datalist id="mn-lista-sec">${[...new Set(mnLista.map((m) => m.seccion).filter(Boolean))].map((s) => `<option value="${esc(s)}">`).join("")}</datalist>` : "")
      : `<p class="muted">${q ? "Ningún manual coincide con la búsqueda." : admin ? "Sube GIF con los botones de arriba." : "Aún no hay manuales publicados."}</p>`;
  }
  $("#mn-buscar").addEventListener("input", renderManuales);

  // El GIF de un manual como archivo. Los publicados se leen de «manuales/<archivo>.js» (una copia en base64),
  // porque al abrir la app con doble clic (file://) el navegador no permite leer archivos de la carpeta con fetch.
  async function blobManual(m) {
    if (m.blob) return m.blob;
    window.__mnDatos = window.__mnDatos || {};
    if (!window.__mnDatos[m.id]) await new Promise((ok, mal) => { const s = document.createElement("script"); s.src = `${m.archivo}.js`; s.onload = ok; s.onerror = mal; document.head.appendChild(s); });
    const url = window.__mnDatos[m.id]; if (!url) throw new Error("sin datos");
    const [cab, b64] = url.split(","), tipo = (cab.match(/data:([^;]+)/) || [])[1] || "image/gif", bin = atob(b64), u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return new Blob([u8], { type: tipo });
  }
  // Siempre como archivo .gif (image/gif) para que llegue animado a WhatsApp y Outlook
  async function archivoManual(m) {
    const b = await blobManual(m);
    let ext = String(m.ext || (m.archivo || "").split(".").pop() || (b.type || "").split("/")[1] || "gif").toLowerCase().replace("jpeg", "jpg");
    if (/gif/.test(b.type) || !/^(png|jpg|webp)$/.test(ext)) ext = "gif";
    const tipo = ext === "gif" ? "image/gif" : b.type || `image/${ext}`;
    return new File([b], `${(m.titulo || "manual").replace(/[\\/:*?"<>|]/g, "_")}.${ext}`, { type: tipo });
  }
  // «Descargar como»: GIF animado o imagen fija (PNG). Se recuerda en este navegador.
  const formatoMn = () => ($("#mn-formato").value === "imagen" ? "imagen" : "gif");
  $("#mn-formato").value = store.get("matriz:mn-formato", "gif") === "imagen" ? "imagen" : "gif";
  $("#mn-formato").addEventListener("change", () => store.set("matriz:mn-formato", formatoMn()));
  const nombreMn = (m, ext) => `${(m.titulo || "manual").replace(/[\\/:*?"<>|]/g, "_")}.${ext}`;
  // Imagen fija: el primer cuadro del GIF en PNG
  async function imagenManual(m) {
    const bmp = await createImageBitmap(await blobManual(m)), cv = document.createElement("canvas");
    cv.width = bmp.width; cv.height = bmp.height;
    const cx = cv.getContext("2d"); cx.fillStyle = "#fff"; cx.fillRect(0, 0, cv.width, cv.height); cx.drawImage(bmp, 0, 0); bmp.close();
    return new File([await canvasBlob(cv)], nombreMn(m, "png"), { type: "image/png" });
  }
  const archivoSegunFormato = (m) => (formatoMn() === "imagen" ? imagenManual(m) : archivoManual(m));
  // Ver en grande con zoom (mismo visor que las fotos; los GIF se ven animados)
  $("#mn-secciones").addEventListener("click", async (e) => {
    const card = e.target.closest("[data-mn]"); if (!card) return;
    if (e.target.closest(".mn-img")) { $("#f-grande").src = card.querySelector("img").src; $("#dlg-ver").showModal(); return; }
    const m = mnLista.find((x) => x.id === card.dataset.mn); if (!m) return;
    // Descargar: disponible para todos los usuarios
    if (e.target.closest("[data-mn-descargar]")) { try { const f = await archivoSegunFormato(m); descargar(f, f.name); } catch { toast("No se pudo leer el manual"); } return; }
    if (!esAdmin()) return;
    if (e.target.closest("[data-mn-quitar]")) {
      if (!confirm(`¿Eliminar el manual «${m.titulo}»?`)) return;
      await manualesDB.del(m.id);
      if (m.publicado) store.set(CLAVE_BORRADOS, [...new Set(store.get(CLAVE_BORRADOS, []).concat(m.id))]);
      return renderManuales();
    }
    const mv = e.target.closest("[data-mn-mover]");
    if (mv) { // intercambia el orden con el vecino de la misma sección
      const sec = mnLista.filter((x) => (x.seccion || "General") === (m.seccion || "General")), i = sec.indexOf(m), j = i + +mv.dataset.mnMover;
      if (j < 0 || j >= sec.length) return;
      sec.forEach((x, k) => (x.orden = k)); [sec[i].orden, sec[j].orden] = [sec[j].orden, sec[i].orden];
      for (const x of [sec[i], sec[j]]) { x.pendiente = true; await manualesDB.put(x); }
      return renderManuales();
    }
  });
  $("#mn-secciones").addEventListener("change", async (e) => {
    const el = e.target.closest("[data-mn-campo]"); if (!el || !esAdmin()) return;
    const m = mnLista.find((x) => x.id === el.closest("[data-mn]").dataset.mn); if (!m) return;
    m[el.dataset.mnCampo] = el.value.trim(); m.pendiente = true;
    await manualesDB.put(m); renderManuales();
  });
  async function subirManuales(files) {
    if (!esAdmin()) return;
    const imgs = [...files].filter((f) => /^image\//.test(f.type) || /\.gif$/i.test(f.name));
    if (!imgs.length) return toast("No se encontraron GIF ni imágenes en la selección");
    let orden = mnLista.length;
    for (const f of imgs.sort((a, b) => (a.webkitRelativePath || a.name).localeCompare(b.webkitRelativePath || b.name, "es", { numeric: true }))) {
      const partes = (f.webkitRelativePath || "").split("/"); // carpeta/subcarpeta/archivo.gif → la subcarpeta es la sección
      const seccion = partes.length > 2 ? partes[partes.length - 2] : partes.length === 2 ? partes[0] : "General";
      const ext = (f.name.match(/\.(\w+)$/) || [, "gif"])[1].toLowerCase();
      const titulo = f.name.replace(/\.\w+$/, "").replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
      // los GIF se guardan tal cual (para no perder la animación)
      await manualesDB.put({ id: nuevoId(), seccion, titulo, descripcion: "", orden: orden++, blob: f, ext, tipo: f.type || "image/gif", pendiente: true });
    }
    toast(`${imgs.length} manual(es) agregado(s). Pulsa «Publicar para todos» para que los vean los auditores.`);
    renderManuales();
  }
  $("#mn-subir-carpeta").addEventListener("click", () => $("#mn-carpeta").click());
  $("#mn-subir").addEventListener("click", () => $("#mn-archivos").click());
  ["#mn-carpeta", "#mn-archivos"].forEach((s) => $(s).addEventListener("change", (e) => { subirManuales(e.target.files); e.target.value = ""; }));
  // Publicar: el admin elige la carpeta de la app; se escriben manuales/*.gif y manuales/manuales.js
  $("#mn-publicar").addEventListener("click", async () => {
    if (!esAdmin()) return;
    if (!window.showDirectoryPicker) return toast("Este navegador no permite escribir en carpetas. Abre la app en Microsoft Edge o Google Chrome.");
    let dir;
    try { dir = await window.showDirectoryPicker({ id: "matriz-app", mode: "readwrite" }); } catch { return; } // canceló
    try { await dir.getFileHandle("index.html"); }
    catch { if (!confirm("Esta carpeta no tiene index.html: no parece ser la carpeta de la app. ¿Publicar aquí de todos modos?")) return; }
    toast("Publicando manuales…");
    try {
      const mdir = await dir.getDirectoryHandle("manuales", { create: true });
      const lista = ordenar(await manualesDB.all()), usados = new Set(["manuales.js"]);
      for (const m of lista) {
        const archivo = m.archivo || `manuales/${m.id}.${m.ext || "gif"}`, nombre = archivo.split("/").pop();
        usados.add(nombre); usados.add(`${nombre}.js`);
        if (m.blob) {
          const fh = await mdir.getFileHandle(nombre, { create: true }), w = await fh.createWritable(); await w.write(m.blob); await w.close();
          // copia en base64 para poder compartirlo desde file:// (ver blobManual)
          const fj = await mdir.getFileHandle(`${nombre}.js`, { create: true }), wj = await fj.createWritable();
          await wj.write(`(window.__mnDatos = window.__mnDatos || {})[${JSON.stringify(m.id)}] = "data:${m.tipo || "image/gif"};base64,${await blobB64(m.blob)}";
`); await wj.close();
        }
        Object.assign(m, { archivo, publicado: true, pendiente: false }); await manualesDB.put(m);
      }
      for await (const [nombre] of mdir.entries()) if (!usados.has(nombre)) await mdir.removeEntry(nombre).catch(() => {}); // GIF eliminados
      const meta = lista.map(({ id, seccion, titulo, descripcion, orden, archivo }) => ({ id, seccion, titulo, descripcion, orden, archivo }));
      const fh = await mdir.getFileHandle("manuales.js", { create: true }), w = await fh.createWritable();
      await w.write(`// Manuales de auditoría publicados desde la app (usuario admin). No editar a mano.\nwindow.MANUALES = ${JSON.stringify(meta, null, 1)};\n`); await w.close();
      window.MANUALES = meta; store.set(CLAVE_BORRADOS, []);
      toast(`Publicado: ${meta.length} manual(es) en la carpeta «manuales» de la app`);
    } catch (err) { toast("No se pudo publicar: revisa los permisos de la carpeta"); }
    renderManuales();
  });

  // ---------- Novedades de activos (usuario de Activos) ----------
  // Junta lo que registraron todos los usuarios (menos admin): equipos por crear, modificados, perdidos, otras novedades y bajas.
  // En este equipo se lee directo de cada usuario; desde otros equipos llega con el archivo de «Enviar a Activos».
  const esActivos = () => !!usuario && usuario.rol === "activos";
  const CLAVE_RECIBIDOS = "matriz:activos:recibidos"; // { usuario: { nombre, enviado, locales: { local: { activos, bajas } } } }
  const claveGestion = () => `matriz:${usuario.usuario}:gestion`; // { idFila: "Pendiente" | "En proceso" | "Atendido" }
  const ESTADOS_GESTION = ["Pendiente", "En proceso", "Atendido"];
  const localPorNombre = new Map(locales.map((l) => [l.local, l]));
  // Lo que le interesa a Activos de un local: activos con novedad, modificados o creados, y las bajas (con las fotos del equipo)
  function novedadesDeEstado(st) {
    const todos = st.activos || [];
    const activos = todos.filter((a) => enUso(a) && (conNovedadActivo(a) || (a.cambios && Object.keys(a.cambios).length) || a.creadoAuditor));
    const bajas = (st.bajas || []).map((b) => ({ ...b, fotos: (todos.find((a) => a.id === b.activoId) || {}).fotos || [] }));
    return { activos, bajas };
  }
  function novedadesDeUsuario(u) {
    const pref = `matriz:${u}:`, locs = {};
    Object.keys(localStorageSeguro()).forEach((k) => {
      if (!k.startsWith(pref) || !localPorNombre.has(k.slice(pref.length))) return;
      const d = novedadesDeEstado(store.get(k, {}));
      if (d.activos.length || d.bajas.length) locs[k.slice(pref.length)] = d;
    });
    return locs;
  }
  function tipoNovedadActivo(a) {
    if (a.creadoAuditor || a.novedad === "Equipo No creado") return "Crear equipo";
    if (a.novedad === "Activo perdido") return "Activo perdido";
    const campos = Object.keys(a.cambios || {}).map((k) => (CAMPOS_MOD[k] || k).toLowerCase());
    if (campos.length) return `Modificar ${unirY(campos)}`;
    return a.novedad || "Novedad";
  }
  // Una fila por novedad, de todos los auditores (lo de este equipo manda sobre lo recibido del mismo local)
  function filasNovedades() {
    const recibidos = store.get(CLAVE_RECIBIDOS, {}), filas = [];
    const usuarios = new Set([...USUARIOS.filter((u) => u.rol === "auditor").map((u) => u.usuario), ...Object.keys(recibidos)]);
    for (const u of usuarios) {
      const info = USUARIOS.find((x) => x.usuario === u) || {}, rec = recibidos[u] || {};
      if (info.rol === "admin" || info.rol === "activos") continue;
      const locs = { ...(rec.locales || {}), ...novedadesDeUsuario(u) }, auditor = info.nombre || rec.nombre || u;
      for (const [local, d] of Object.entries(locs)) {
        const l = localPorNombre.get(local) || {}, base = { usuario: u, auditor, local, cadena: l.cadena || "Sin marca", zona: l.zona || "" };
        const conBaja = new Set((d.bajas || []).map((b) => b.activoId));
        (d.activos || []).filter((a) => !conBaja.has(a.id)).forEach((a) => filas.push({ ...base, id: `${u}|${local}|${a.id}`, clase: "activo", tipo: tipoNovedadActivo(a),
          codigo: a.codigo, nombre: a.nombre, marca: a.marca, modelo: a.modelo, serie: a.serie, cambios: a.cambios || {}, obs: (a.perdido && a.perdido.obs) || a.obs || "",
          fecha: a.creadoAuditor || (a.perdido && a.perdido.fecha) || Object.values(a.cambios || {}).map((c) => c.fecha).sort().pop() || "", fotos: a.fotos || [] }));
        (d.bajas || []).forEach((b) => filas.push({ ...base, id: `${u}|${local}|baja:${b.id}`, clase: "baja", tipo: bajaPendiente(b) ? "Baja de equipo (pendiente)" : "Baja de equipo (aprobada)",
          codigo: b.codigo, nombre: b.nombre, marca: b.marca, modelo: b.modelo, serie: b.serie, cambios: {}, valor: b.valor, motivo: b.motivo,
          obs: [b.motivo, b.obs, b.gerente && `Certifica: ${b.gerente}`].filter(Boolean).join(" · "), fecha: b.fechaAprob || b.fecha || "", fotos: b.fotos || [] }));
      }
    }
    return filas;
  }
  const cmpTxt = (a, b) => String(a || "").localeCompare(String(b || ""), "es", { numeric: true, sensitivity: "base" });
  let naUrls = [];
  function renderNovedadesActivos() {
    if (!esActivos()) return;
    const todas = filasNovedades(), gestion = store.get(claveGestion(), {}), estadoDe = (f) => gestion[f.id] || "Pendiente";
    const recibidos = store.get(CLAVE_RECIBIDOS, {});
    // filtros (las listas se arman con lo que hay)
    const llenar = (sel, valores, todos) => { const v = sel.value; sel.innerHTML = `<option value="">${todos}</option>` + valores.map((x) => `<option>${esc(x)}</option>`).join(""); sel.value = valores.includes(v) ? v : ""; };
    llenar($("#na-tipo"), [...new Set(todas.map((f) => f.tipo))].sort(cmpTxt), "Todas las novedades");
    llenar($("#na-auditor"), [...new Set(todas.map((f) => f.auditor))].sort(cmpTxt), "Todos los auditores");
    const q = $("#na-buscar").value.trim().toLowerCase(), tipo = $("#na-tipo").value, aud = $("#na-auditor").value, est = $("#na-estado").value;
    const porMarcaEquipo = $("#na-orden").value === "marca", grupoDe = (f) => (porMarcaEquipo ? (f.marca || "Sin marca").toUpperCase() : f.cadena);
    const vista = todas.filter((f) => (!tipo || f.tipo === tipo) && (!aud || f.auditor === aud) && (!est || estadoDe(f) === est)
      && (!q || [f.codigo, f.nombre, f.marca, f.modelo, f.serie, f.local, f.auditor, f.tipo, f.obs].join(" ").toLowerCase().includes(q)))
      .sort((a, b) => cmpTxt(grupoDe(a), grupoDe(b)) || cmpTxt(a.local, b.local) || cmpTxt(a.tipo, b.tipo) || cmpTxt(a.nombre, b.nombre));
    const cuenta = (fn) => todas.filter(fn).length;
    $("#na-kpis").innerHTML = [[todas.length, "Novedades"], [cuenta((f) => f.tipo === "Crear equipo"), "Crear equipo"], [cuenta((f) => f.tipo.startsWith("Modificar")), "Modificar"],
      [cuenta((f) => f.tipo === "Activo perdido"), "Perdidos"], [cuenta((f) => f.clase === "baja"), "Bajas"], [cuenta((f) => estadoDe(f) !== "Atendido"), "Por atender"]]
      .map(([n, t]) => `<div><b>${n}</b><span>${t}</span></div>`).join("");
    const nRec = Object.keys(recibidos).length;
    $("#na-info").textContent = `${todas.length} novedad(es) de ${new Set(todas.map((f) => f.usuario)).size} auditor(es)`
      + (nRec ? ` · archivos recibidos de ${nRec} auditor(es), el último el ${Object.values(recibidos).map((r) => r.enviado || "").sort().pop().slice(0, 10)}` : " · carga los archivos que te envían los auditores");
    naUrls.forEach((u) => URL.revokeObjectURL(u)); naUrls = [];
    const grupos = vista.reduce((g, f) => ((g[grupoDe(f)] = g[grupoDe(f)] || []).push(f), g), {});
    const det = (f) => {
      const v = (k) => { const c = f.cambios[k]; return c ? `<span class="mod" title="Antes: ${esc(c.antes || "(vacío)")}">${esc(f[k] || "—")}</span> <small class="antes">antes: ${esc(c.antes || "—")}</small>` : esc(f[k] || "—"); };
      return `<b>${f.cambios.nombre ? v("nombre") : esc(f.nombre)}</b><small>${esc(f.codigo || "SIN CÓDIGO")} · Marca: ${v("marca")} · Modelo: ${v("modelo")} · Serie: ${v("serie")}${f.clase === "baja" && f.valor !== "" && f.valor != null ? ` · Valor en libros: ${fmtUSD(+f.valor)}` : ""}</small>${f.obs ? `<small>${esc(f.obs)}</small>` : ""}`;
    };
    const claseTipo = (t) => (t === "Crear equipo" ? "crear" : t === "Activo perdido" ? "perdido" : t.startsWith("Modificar") ? "modificar" : t.startsWith("Baja") ? "baja" : "otra");
    $("#na-lista").innerHTML = vista.length ? Object.entries(grupos).map(([g, fs]) => `<div class="na-grupo">
      <h3>${esc(g)} <small>${fs.length} novedad(es) · ${new Set(fs.map((f) => f.local)).size} local(es)</small></h3>
      <div class="tabla-scroll"><table class="na-tabla"><thead><tr><th>Novedad</th><th>Local</th><th>Equipo</th><th>Auditor</th><th>Fecha</th><th>Fotos</th><th>Gestión</th></tr></thead><tbody>
      ${fs.map((f) => `<tr data-na="${esc(f.id)}" class="${estadoDe(f) === "Atendido" ? "hecho" : ""}">
        <td><span class="na-tipo ${claseTipo(f.tipo)}">${esc(f.tipo)}</span></td>
        <td>${esc(f.local)}${porMarcaEquipo ? `<small>${esc(f.cadena)}</small>` : ""}</td><td class="na-eq">${det(f)}</td>
        <td>${esc(f.auditor)}</td><td class="nowrap">${esc(f.fecha)}</td>
        <td class="na-fotos">${f.fotos.slice(0, 4).map((id) => `<img data-na-foto="${esc(id)}" alt="Foto de ${esc(f.nombre)}">`).join("")}${f.fotos.length > 4 ? `<small>+${f.fotos.length - 4}</small>` : ""}${f.fotos.length ? "" : `<small class="muted">—</small>`}</td>
        <td><select data-na-estado aria-label="Gestión de ${esc(f.nombre)}">${ESTADOS_GESTION.map((e) => `<option${e === estadoDe(f) ? " selected" : ""}>${e}</option>`).join("")}</select></td></tr>`).join("")}
      </tbody></table></div></div>`).join("")
      : `<p class="muted">${todas.length ? "Ninguna novedad coincide con los filtros." : "Aún no hay novedades de activos. Aquí llegan las que registran los auditores en este equipo y las de los archivos «Enviar a Activos» que cargues."}</p>`;
    // fotos desde la base de fotos de este equipo
    $$("#na-lista img[data-na-foto]").forEach(async (img) => {
      const f = await fotosDB.get(img.dataset.naFoto).catch(() => null);
      if (!f || !f.blob) { img.replaceWith(Object.assign(document.createElement("small"), { className: "muted", textContent: "sin foto" })); return; }
      const u = URL.createObjectURL(f.blob); naUrls.push(u); img.src = u;
    });
  }
  ["#na-buscar", "#na-tipo", "#na-auditor", "#na-estado", "#na-orden"].forEach((s) => $(s).addEventListener(s === "#na-buscar" ? "input" : "change", renderNovedadesActivos));
  $("#na-lista").addEventListener("change", (e) => {
    const sel = e.target.closest("[data-na-estado]"); if (!sel) return;
    const g = store.get(claveGestion(), {}); g[sel.closest("[data-na]").dataset.na] = sel.value; store.set(claveGestion(), g);
    renderNovedadesActivos();
  });
  $("#na-lista").addEventListener("click", (e) => {
    const img = e.target.closest("img[data-na-foto]"); if (!img || !img.src) return;
    $("#f-grande").src = img.src; $("#dlg-ver").showModal();
  });
  // Excel (CSV con «;», como lo abre Excel en español)
  $("#na-excel").addEventListener("click", () => {
    const gestion = store.get(claveGestion(), {}), filas = filasNovedades().sort((a, b) => cmpTxt(a.cadena, b.cadena) || cmpTxt(a.local, b.local) || cmpTxt(a.tipo, b.tipo));
    if (!filas.length) return toast("No hay novedades para descargar");
    const c = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
    const cab = ["Marca local", "Zona", "Local", "Novedad", "Código", "Equipo", "Marca", "Modelo", "Serie", "Valor anterior", "Observación", "Auditor", "Fecha", "Fotos", "Gestión"];
    const csv = [cab, ...filas.map((f) => [f.cadena, f.zona, f.local, f.tipo, f.codigo, f.nombre, f.marca, f.modelo, f.serie,
      Object.entries(f.cambios).map(([k, x]) => `${CAMPOS_MOD[k] || k}: ${x.antes || "(vacío)"}`).join(" | "), f.obs, f.auditor, f.fecha, f.fotos.length, gestion[f.id] || "Pendiente"])]
      .map((r) => r.map(c).join(";")).join("\r\n");
    descargar(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" }), `Novedades de activos ${new Date().toISOString().slice(0, 10)}.csv`);
  });
  // Cargar lo que envían los auditores: el .json o el correo .eml que lo trae adjunto
  $("#na-cargar").addEventListener("click", () => $("#na-archivos").click());
  $("#na-archivos").addEventListener("change", async (e) => {
    const files = [...e.target.files]; e.target.value = "";
    const rec = store.get(CLAVE_RECIBIDOS, {}); let ok = 0, malos = 0, viejos = 0, nFotos = 0;
    for (const file of files) {
      try {
        let txt = await file.text();
        if (/\.eml$/i.test(file.name) || /^[\w-]+:/m.test(txt.slice(0, 200)) && !txt.trim().startsWith("{")) {
          const m = /Content-Type: application\/json[^]*?\r?\n\r?\n([A-Za-z0-9+/=\r\n]+)/i.exec(txt); if (!m) throw new Error("sin adjunto");
          txt = new TextDecoder().decode(Uint8Array.from(atob(m[1].replace(/\s+/g, "")), (ch) => ch.charCodeAt(0)));
        }
        const d = JSON.parse(txt);
        if (d.tipo !== "matriz-novedades-activos" || !d.usuario) throw new Error("no es un archivo de novedades");
        if (rec[d.usuario] && rec[d.usuario].enviado > d.enviado) { viejos++; continue; } // ya hay uno más reciente
        for (const [id, url] of Object.entries(d.fotos || {})) {
          try { await fotosDB.put({ id, blob: await dataUrlABlob(url), local: "recibido", activo: "recibido", fecha: d.enviado }); nFotos++; } catch { /* foto dañada */ }
        }
        rec[d.usuario] = { nombre: d.nombre, enviado: d.enviado, locales: d.locales || {} }; ok++;
      } catch { malos++; }
    }
    store.set(CLAVE_RECIBIDOS, rec);
    toast(`${ok} archivo(s) cargado(s)${nFotos ? ` con ${nFotos} foto(s)` : ""}${viejos ? ` · ${viejos} más antiguo(s) que lo ya cargado` : ""}${malos ? ` · ${malos} no válido(s)` : ""}`);
    renderNovedadesActivos();
  });
  // Auditores: todas sus novedades de activos y bajas, con fotos, en un correo con el archivo adjunto para Activos
  $$("[data-enviar-activos]").forEach((btn) => btn.addEventListener("click", async () => {
    const locs = novedadesDeUsuario(usuario.usuario), n = Object.values(locs).reduce((s, d) => s + d.activos.length + d.bajas.length, 0);
    if (!n) return toast("No tienes novedades de activos ni bajas registradas");
    toast("Preparando el archivo para Activos…");
    const fotos = {};
    for (const d of Object.values(locs)) for (const x of [...d.activos, ...d.bajas]) for (const id of x.fotos || []) {
      if (fotos[id]) continue; const f = await fotosDB.get(id).catch(() => null);
      if (f && f.blob) fotos[id] = `data:${f.blob.type || "image/jpeg"};base64,${await blobB64(f.blob)}`;
    }
    const enviado = new Date().toISOString(), hoy = enviado.slice(0, 10);
    const json = JSON.stringify({ tipo: "matriz-novedades-activos", v: 1, usuario: usuario.usuario, nombre: usuario.nombre, enviado, locales: locs, fotos });
    const b64 = (s) => { const b = new TextEncoder().encode(s); let t = ""; for (let i = 0; i < b.length; i += 8192) t += String.fromCharCode(...b.subarray(i, i + 8192)); return btoa(t); };
    const lineas = (s) => s.replace(/.{76}/g, "$&\r\n"), lim = "----=_activos_" + Date.now(), nom = `Novedades activos ${usuario.nombre} ${hoy}.json`.replace(/[^\x20-\x7e]/g, "_");
    const resumen = Object.entries(locs).map(([l, d]) => `<li><b>${esc(l)}</b>: ${[...d.activos.map((a) => `${esc(tipoNovedadActivo(a))} – ${esc(a.nombre)}`), ...d.bajas.map((b) => `Baja ${bajaPendiente(b) ? "pendiente" : "aprobada"} – ${esc(b.nombre)}`)].join("; ")}</li>`).join("");
    const html = `<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:11pt">Estimada Amanda,<br><br>Adjunto las novedades de activos y bajas registradas (${n}). Para verlas, carga el archivo adjunto en la pestaña <b>Novedades de activos</b>.<ul>${resumen}</ul>Saludos cordiales,<br>${esc(usuario.nombre)}<br>Auditoría · Grupo KFC Ecuador</body></html>`;
    const eml = [`Subject: =?UTF-8?B?${b64(`Novedades de activos – ${usuario.nombre} – ${hoy}`)}?=`, "X-Unsent: 1", "MIME-Version: 1.0", `Content-Type: multipart/mixed; boundary="${lim}"`, "",
      `--${lim}`, "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: base64", "", lineas(b64(html)),
      `--${lim}`, `Content-Type: application/json; name="${nom}"`, "Content-Transfer-Encoding: base64", `Content-Disposition: attachment; filename="${nom}"`, "", lineas(b64(json)),
      `--${lim}--`, ""].join("\r\n");
    descargar(new Blob([eml], { type: "message/rfc822" }), `Novedades activos ${usuario.nombre} ${hoy}.eml`);
    toast(`Correo listo con ${n} novedad(es)${Object.keys(fotos).length ? ` y ${Object.keys(fotos).length} foto(s)` : ""}: ábrelo y envíaselo a Amanda Taimal`);
  }));

  // ---------- Valor en libros (base contable de activos fijos) ----------
  // El admin sube la base (hoja «ACTIVOS FIJOS»: Cadena, Codigo, Local, Nombre, Marca, Modelo, Serie, valor en libros por mes,
  // fecha de compra, última baja, valor de adquisición) y elige el mes vigente. Se cruza con el listado del SIR de cada local
  // por código (y por serie si el código no está). En este equipo se guarda en IndexedDB; para los demás, «Publicar para todos»
  // la escribe en libros/libros.js dentro de la carpeta de la app.
  // Cada activo: { c: código, l: local, n, ma, mo, s, f: fecha compra, a: valor adquisición, v: [valor en libros por mes], b: última baja }
  const librosDB = (() => {
    let abierta;
    const abrir = () => abierta || (abierta = new Promise((ok, mal) => {
      const r = indexedDB.open("matriz-libros", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("base");
      r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error);
    }));
    const tx = async (modo, fn) => { const db = await abrir(); return new Promise((ok, mal) => { const t = db.transaction("base", modo), req = fn(t.objectStore("base")); t.oncomplete = () => ok(req.result); t.onerror = () => mal(t.error); }); };
    return { get: () => tx("readonly", (s) => s.get("base")), put: (v) => tx("readwrite", (s) => s.put(v, "base")) };
  })();
  let libros = null, librosCargando = null; // { meta, porCodigo, porLocal, porSerie }
  let librosIntentado = false;
  // Celda «Valor en libros» del listado de activos
  function celdaLibros(a) {
    if (!libros) return `<small class="muted">—</small>`;
    const x = libroDe(a, localActual.local);
    if (!x) return `<small class="muted">sin registro</small>`;
    return `<b>${fmtUSD(valorLibro(x) || 0)}</b>${x.b ? `<br><small class="tag bad">baja contable</small>` : ""}`;
  }
  const MESES_NOMBRE = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  const serieNorm = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const codLocal = (local) => ((/^([A-Z0-9]+)\s*-/i.exec(local || "") || [])[1] || "").toUpperCase();
  function indexarLibros(base) {
    const porCodigo = new Map(), porLocal = new Map(), porSerie = new Map();
    base.items.forEach((x) => {
      porCodigo.set(x.c, x);
      if (!porLocal.has(x.l)) porLocal.set(x.l, []); porLocal.get(x.l).push(x);
      const s = serieNorm(x.s); if (s.length >= 4 && !/^SN$|^SM$/.test(s)) porSerie.set(`${x.l}|${s}`, x);
    });
    return { meta: base.meta, porCodigo, porLocal, porSerie };
  }
  // Lo guardado en este equipo o lo publicado en la carpeta: lo más reciente
  function cargarLibros(forzar) {
    if (libros && !forzar) return Promise.resolve(libros);
    if (librosCargando && !forzar) return librosCargando;
    librosCargando = (async () => {
      const local = await librosDB.get().catch(() => null);
      const pub = await new Promise((ok) => {
        if (window.__libros) return ok(window.__libros);
        const s = document.createElement("script"); s.src = `libros/libros.js?t=${Date.now()}`;
        s.onload = () => ok(window.__libros || null); s.onerror = () => ok(null); document.head.appendChild(s);
      });
      const base = [local, pub].filter((b) => b && b.items).sort((a, b) => String(b.meta.cargado).localeCompare(String(a.meta.cargado)))[0];
      libros = base ? indexarLibros(base) : null; librosCargando = null;
      return libros;
    })();
    return librosCargando;
  }
  const mesLibros = () => (libros ? libros.meta.meses[libros.meta.mes] : "");
  // Registro contable de un activo del listado: por código; si no, por serie dentro del mismo local
  function libroDe(a, local) {
    if (!libros || !a) return null;
    return libros.porCodigo.get(String(a.codigo || "").trim().toUpperCase()) || libros.porSerie.get(`${codLocal(local)}|${serieNorm(a.serie)}`) || null;
  }
  const valorLibro = (x) => (x && x.v ? x.v[libros.meta.mes] ?? null : null);
  // ---- Subir la base (admin) ----
  async function subirLibros(file) {
    if (!file || !esAdmin()) return;
    toast("Leyendo la base de valor en libros… (puede tardar unos segundos)");
    let libro;
    try { const buf = await file.arrayBuffer(); if (new Uint8Array(buf)[0] !== 0x50) throw new Error(); libro = await leerXlsx(buf); }
    catch { return toast("No se pudo leer: debe ser el Excel .xlsx de la base de valor en libros"); }
    let filas = null, iCab = -1, hojaUsada = "";
    for (const n of libro.nombres) { // la primera hoja con encabezados Codigo + Serie + meses
      const fs = await libro.hoja(n).catch(() => []);
      const i = fs.slice(0, 15).findIndex((f) => f.some((c) => /^c[oó]digo$/i.test(String(c).trim())) && f.some((c) => /^serie$/i.test(String(c).trim())));
      if (i >= 0 && fs[i].some((c) => MESES_NOMBRE.includes(String(c).trim().toUpperCase()))) { filas = fs; iCab = i; hojaUsada = n; break; }
    }
    if (!filas) return toast("No encontré la hoja con Codigo, Serie y los meses del valor en libros");
    const cab = filas[iCab].map((c) => String(c).trim()), col = (re) => cab.findIndex((c) => re.test(c));
    const C = { codigo: col(/^c[oó]digo$/i), local: col(/^local$/i), nombre: col(/^nombre$/i), marca: col(/^marca$/i), modelo: col(/^modelo$/i), serie: col(/^serie$/i),
      compra: col(/fecha\s*compra/i), baja: col(/baja/i), adq: col(/adquisici/i), cc: col(/centro/i) };
    // primer bloque seguido de meses (a la derecha hay otro bloque de meses: es la depreciación)
    const iMes = cab.findIndex((c) => MESES_NOMBRE.includes(c.toUpperCase())), meses = [];
    for (let k = iMes; k < cab.length && MESES_NOMBRE.includes(cab[k].toUpperCase()); k++) meses.push(cab[k].toUpperCase());
    const anio = (/20\d\d/.exec(file.name) || [])[0] || "";
    const fecha = (v) => (typeof v === "number" && v > 20000 ? fechaConc(v) : String(v || "").trim());
    const r2 = (v) => (typeof v === "number" ? Math.round(v * 100) / 100 : Number.isFinite(numeroConc(v)) ? Math.round(numeroConc(v) * 100) / 100 : null);
    const items = [];
    for (const f of filas.slice(iCab + 1)) {
      const c = String(f[C.codigo] ?? "").trim().toUpperCase(); if (!c) continue;
      const l = String(f[C.local] ?? "").trim().toUpperCase() || codLocal(String(f[C.cc] ?? ""));
      items.push({ c, l, n: String(f[C.nombre] ?? "").trim(), ma: String(f[C.marca] ?? "").trim(), mo: String(f[C.modelo] ?? "").trim(), s: String(f[C.serie] ?? "").trim(),
        f: C.compra >= 0 ? fecha(f[C.compra]) : "", b: C.baja >= 0 ? fecha(f[C.baja]) : "", a: C.adq >= 0 ? r2(f[C.adq]) : null, v: meses.map((_, k) => r2(f[iMes + k])) });
    }
    if (!items.length) return toast("La hoja no tiene activos");
    const hoy = MESES_NOMBRE[new Date().getMonth()], mes = meses.includes(hoy) ? meses.indexOf(hoy) : meses.length - 1;
    const base = { meta: { archivo: file.name, hoja: hojaUsada, cargado: new Date().toISOString(), por: usuario.nombre, meses: meses.map((m) => `${m}${anio ? ` ${anio}` : ""}`), mes, total: items.length }, items };
    try { await librosDB.put(base); } catch { return toast("No se pudo guardar la base en este equipo"); }
    libros = indexarLibros(base);
    toast(`Base cargada: ${items.length.toLocaleString("es-EC")} activos · valor en libros a ${mesLibros()}. Pulsa «Publicar para todos» para que la vean los auditores.`);
    renderActivos();
  }
  $("#vl-subir").addEventListener("click", () => $("#vl-archivo").click());
  $("#vl-archivo").addEventListener("change", (e) => { subirLibros(e.target.files[0]); e.target.value = ""; });
  $("#vl-mes").addEventListener("change", async () => {
    if (!libros || !esAdmin()) return;
    const base = await librosDB.get().catch(() => null); if (!base) return toast("Vuelve a subir la base en este equipo para cambiar el mes");
    base.meta.mes = +$("#vl-mes").value; base.meta.cargado = new Date().toISOString(); await librosDB.put(base);
    libros = indexarLibros(base); toast(`Valor en libros a ${mesLibros()}. Publica de nuevo para los auditores.`); renderActivos();
  });
  $("#vl-publicar").addEventListener("click", async () => {
    if (!esAdmin()) return;
    const base = await librosDB.get().catch(() => null); if (!base) return toast("Primero sube la base en este equipo");
    if (!window.showDirectoryPicker) return toast("Este navegador no permite escribir en carpetas. Abre la app en Microsoft Edge o Google Chrome.");
    let dir;
    try { dir = await window.showDirectoryPicker({ id: "matriz-app", mode: "readwrite" }); } catch { return; }
    try { await dir.getFileHandle("index.html"); }
    catch { if (!confirm("Esta carpeta no tiene index.html: no parece ser la carpeta de la app. ¿Publicar aquí de todos modos?")) return; }
    try {
      toast("Publicando la base de valor en libros…");
      const ld = await dir.getDirectoryHandle("libros", { create: true }), fh = await ld.getFileHandle("libros.js", { create: true }), w = await fh.createWritable();
      await w.write(`// Base de valor en libros publicada desde la app (admin). No editar a mano.\nwindow.__libros = ${JSON.stringify(base)};\n`); await w.close();
      window.__libros = base; toast(`Publicado: los auditores verán el valor en libros a ${base.meta.meses[base.meta.mes]}`);
    } catch { toast("No se pudo publicar: revisa los permisos de la carpeta"); }
  });
  // Barra de valor en libros en Activos: subir (admin) y verificar (todos)
  function renderBarraLibros() {
    const admin = esAdmin();
    // El recuadro de valor en libros es solo del admin; el admin no carga listados del SIR ni respaldos de la app
    $("#vl-carga").hidden = !admin; $("#a-carga").hidden = admin; $("#j-carga").hidden = admin;
    $("#vl-admin").hidden = !admin;
    $("#vl-verificar").disabled = !libros || !localActual;
    if (!libros) { $("#vl-info").textContent = admin ? "Sube el Excel de la base de valor en libros (hoja ACTIVOS FIJOS) para cruzarlo con los activos de cada local." : "Aún no hay base de valor en libros publicada por el administrador."; return; }
    const m = libros.meta;
    $("#vl-info").textContent = `Valor en libros a ${mesLibros()} · ${m.archivo} · ${m.total.toLocaleString("es-EC")} activos · cargado el ${String(m.cargado).slice(0, 10)}${m.por ? ` por ${m.por}` : ""}`;
    if (admin) $("#vl-mes").innerHTML = m.meses.map((x, i) => `<option value="${i}"${i === m.mes ? " selected" : ""}>${esc(x)}</option>`).join("");
  }
  // ---- Verificar el valor en libros del local ----
  function datosVerificacion() {
    const st = activosLocal(), cod = codLocal(localActual.local), usados = new Set();
    const lista = st.activos.filter(enUso).map((a) => { const x = libroDe(a, localActual.local); if (x) usados.add(x.c); return { a, x }; });
    const conValor = lista.filter((r) => r.x), sinRegistro = lista.filter((r) => !r.x);
    const soloContable = (libros.porLocal.get(cod) || []).filter((x) => !usados.has(x.c) && !x.b); // en la contabilidad del local, sin baja, y no está en el listado
    const suma = (xs) => xs.reduce((s, x) => s + (valorLibro(x) || 0), 0);
    return { st, lista, conValor, sinRegistro, soloContable, total: suma(conValor.map((r) => r.x)), totalSolo: suma(soloContable), cero: conValor.filter((r) => !valorLibro(r.x)).length,
      bajaContable: conValor.filter((r) => r.x.b) };
  }
  async function verificarLibros() {
    if (!localActual) return toast("Selecciona un local primero");
    await cargarLibros(); if (!libros) return toast("Aún no hay base de valor en libros");
    const d = datosVerificacion(), fila = (a, x, extra = "") => `<tr><td>${esc(a ? a.codigo : x.c)}</td><td><b>${esc(a ? a.nombre : x.n)}</b><small>${esc([a ? a.marca : x.ma, a ? a.modelo : x.mo].filter(Boolean).join(" "))}${(a ? a.serie : x.s) ? ` · S: ${esc(a ? a.serie : x.s)}` : ""}</small></td>
      <td class="num">${x ? `<b>${fmtUSD(valorLibro(x) || 0)}</b>` : "—"}</td><td class="num">${x && x.a != null ? fmtUSD(x.a) : "—"}</td><td class="nowrap">${x && x.f ? esc(x.f) : "—"}</td><td>${extra}</td></tr>`;
    const tabla = (titulo, filas, nota = "") => `<h3 class="vl-h">${titulo} <small>${nota}</small></h3><div class="tabla-scroll"><table class="na-tabla vl-tabla"><thead><tr><th>Código</th><th>Activo</th><th class="num">Valor en libros</th><th class="num">Adquisición</th><th>Compra</th><th></th></tr></thead><tbody>${filas || `<tr><td colspan="6" class="muted">Ninguno.</td></tr>`}</tbody></table></div>`;
    $("#vl-cuerpo").innerHTML = `<div class="kpis mini">
        <div><b>${d.lista.length}</b><span>Activos en el listado del local</span></div><div><b>${d.conValor.length}</b><span>Con registro contable</span></div>
        <div><b>${fmtUSD(d.total)}</b><span>Valor en libros (listado)</span></div><div><b>${d.cero}</b><span>Con valor cero en libros</span></div>
        <div><b>${d.sinRegistro.length}</b><span>Sin registro contable</span></div><div><b>${d.soloContable.length}</b><span>En contabilidad y no en el listado</span></div></div>
      ${tabla("Activos del listado con su valor en libros", d.conValor.sort((p, q) => (valorLibro(q.x) || 0) - (valorLibro(p.x) || 0)).map((r) => fila(r.a, r.x, [r.x.b ? `<span class="tag bad">Baja contable ${esc(r.x.b)}</span>` : "", r.x.c !== String(r.a.codigo || "").toUpperCase() ? `<small>encontrado por serie (${esc(r.x.c)})</small>` : ""].join(" "))).join(""), `a ${esc(mesLibros())}`)}
      ${tabla("Sin registro en la base contable", d.sinRegistro.map((r) => fila(r.a, null, `<small>${esc(r.a.estado || "")}</small>`)).join(""), "revisar que el código o la serie estén bien en el SIR")}
      ${tabla("En la contabilidad del local pero no en el listado", d.soloContable.map((x) => fila(null, x, `<span class="tag warn">No está en el listado</span>`)).join(""), d.soloContable.length ? `${fmtUSD(d.totalSolo)} en libros · posibles faltantes o transferidos` : "")}`;
    $("#vl-titulo").textContent = `Valor en libros · ${localActual.local}`;
    $("#vl-sub").textContent = `Base ${libros.meta.archivo} · valor a ${mesLibros()}`;
    $("#dlg-vl").showModal();
  }
  $("#vl-verificar").addEventListener("click", verificarLibros);
  $("#vl-cerrar").addEventListener("click", () => $("#dlg-vl").close());
  $("#vl-excel").addEventListener("click", () => {
    if (!libros || !localActual) return;
    const d = datosVerificacion(), c = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
    const filas = [["Situación", "Código", "Activo", "Marca", "Modelo", "Serie", `Valor en libros ${mesLibros()}`, "Valor adquisición", "Fecha compra", "Baja contable"],
      ...d.conValor.map(({ a, x }) => ["Con registro", a.codigo, a.nombre, a.marca, a.modelo, a.serie, valorLibro(x), x.a, x.f, x.b]),
      ...d.sinRegistro.map(({ a }) => ["Sin registro contable", a.codigo, a.nombre, a.marca, a.modelo, a.serie, "", "", "", ""]),
      ...d.soloContable.map((x) => ["En contabilidad, no en el listado", x.c, x.n, x.ma, x.mo, x.s, valorLibro(x), x.a, x.f, x.b])];
    descargar(new Blob(["﻿" + filas.map((r) => r.map((v) => (typeof v === "number" ? String(v).replace(".", ",") : c(v))).join(";")).join("\r\n")], { type: "text/csv;charset=utf-8" }),
      `Valor en libros ${localActual.local} ${new Date().toISOString().slice(0, 10)}.csv`.replace(/[\\/:*?"<>|]/g, "_"));
  });

  // ---------- Conciliación bancaria ----------
  // El usuario de Conciliación sube el Excel; la app busca el local de cada fila y deja como novedad las filas con diferencia
  // o con observación. Cada auditor ve en «Conciliación bancaria» las novedades de sus locales.
  // Llega a los auditores: en este mismo equipo (localStorage), publicada en la carpeta de la app (conciliacion/conciliacion.js)
  // o con el correo que genera la app (trae el archivo adjunto para cargarlo).
  const esConciliacion = () => !!usuario && usuario.rol === "conciliacion";
  const puedeConciliar = () => esConciliacion(); // solo el usuario de Conciliación sube la conciliación
  // El Excel escribe los nombres a su manera («Walter Pin», «Janeth Insuasti»…): se comparan por las 3 primeras letras de nombre y apellido
  const claveNombre = (s) => { const p = normalTxt(String(s || "").replace(/\d+/g, "")).split(" ").filter(Boolean); return p.length ? `${p[0].slice(0, 3)}|${p[p.length - 1].slice(0, 3)}` : ""; };
  const AUDITORES_MATRIZ = [...new Set(locales.map((l) => l.auditor).filter(Boolean))];
  // Responsable de las novedades de un local: el auditor que trae el Excel (con el nombre de la matriz) o, si no trae, el de la matriz
  function responsableConc(local, xs) {
    const deMatriz = (localPorNombre.get(local) || {}).auditor || "", ex = (xs.find((x) => x.auditor) || {}).auditor;
    if (!ex || claveNombre(ex) === claveNombre(deMatriz)) return deMatriz || ex || "Sin auditor";
    return AUDITORES_MATRIZ.find((a) => claveNombre(a) === claveNombre(ex)) || ex;
  }
  const esDelAuditor = (local, xs) => !!usuario.auditor && ((localPorNombre.get(local) || {}).auditor === usuario.auditor || xs.some((x) => x.auditor && claveNombre(x.auditor) === claveNombre(usuario.auditor)));
  const CLAVE_CONC = "matriz:conciliacion:datos"; // { archivo, cargado, columnas, porLocal: { local: [novedad] }, sinLocal: [{ texto, n }], filas }
  const claveRevision = () => `matriz:${usuario.usuario}:conc-rev`; // { idNovedad: { estado, obs } }
  const CAMPOS_CONC = [
    ["local", "Local *", /local|tienda|restaurante|punto de venta|\bpdv\b|centro de costo|sucursal|establecimiento/i],
    ["valor", "Valor / diferencia *", /diferencia|dif\.|valor|monto|importe|saldo|total/i],
    ["fecha", "Fecha", /fecha/i],
    ["banco", "Banco / cuenta", /banco|cuenta|entidad/i],
    ["doc", "Documento / referencia", /refer|document|comprob|papeleta|n[°ºo]\.?\s*dep|dep[oó]sito|transacci/i],
    ["novedad", "Novedad / observación", /novedad|observ|concepto|detalle|descripci|motivo|estado|comentario/i],
  ];
  // Datos vigentes: lo publicado en la carpeta o lo guardado en este equipo, lo más reciente
  function datosConciliacion() {
    const local = store.get(CLAVE_CONC, null), pub = window.CONCILIACION || null;
    return [local, pub].filter(Boolean).sort((a, b) => String(b.cargado).localeCompare(String(a.cargado)))[0] || null;
  }
  // Lector de Excel sin librerías: .xlsx (zip), .xls exportado como HTML y .csv. Devuelve la primera hoja como filas de celdas.
  // libro = { nombres, hoja(nombre) } solo para .xlsx (para buscar la hoja «ERRORES»); filas = primera hoja
  async function leerHojaExcel(file) {
    const buf = await file.arrayBuffer(), u8 = new Uint8Array(buf);
    if (u8[0] === 0x50 && u8[1] === 0x4b) { const libro = await leerXlsx(buf); return { libro, filas: await libro.hoja(libro.nombres[0]) }; }
    return { libro: null, filas: await leerHojaTexto(file, u8) };
  }
  async function leerHojaTexto(file, u8) {
    if (u8[0] === 0xd0 && u8[1] === 0xcf) throw new Error("xls-binario");
    const t = await leerArchivoTexto(file);
    if (/<table/i.test(t)) {
      const doc = new DOMParser().parseFromString(t, "text/html");
      return [...doc.querySelectorAll("tr")].map((tr) => [...tr.children].map((c) => c.textContent.replace(/\s+/g, " ").trim()));
    }
    const sep = (t.split("\n")[0].match(/;/g) || []).length >= (t.split("\n")[0].match(/,/g) || []).length ? ";" : t.includes("\t") ? "\t" : ",";
    return t.split(/\r?\n/).map((l) => { const out = []; let cur = "", q = false;
      for (const ch of l) { if (ch === '"') q = !q; else if (ch === sep && !q) { out.push(cur.trim()); cur = ""; } else cur += ch; }
      out.push(cur.trim()); return out; });
  }
  async function leerXlsx(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(buf), dec = new TextDecoder();
    let e = u8.length - 22; while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error("zip dañado");
    const zip = {}; let p = dv.getUint32(e + 16, true);
    for (let i = 0, n = dv.getUint16(e + 10, true); i < n; i++) {
      const ln = dv.getUint16(p + 28, true), le = dv.getUint16(p + 30, true), lc = dv.getUint16(p + 32, true);
      zip[dec.decode(u8.subarray(p + 46, p + 46 + ln))] = { metodo: dv.getUint16(p + 10, true), tam: dv.getUint32(p + 20, true), off: dv.getUint32(p + 42, true) };
      p += 46 + ln + le + lc;
    }
    const leer = async (nombre) => {
      const f = zip[nombre]; if (!f) return "";
      const ini = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true), datos = u8.subarray(ini, ini + f.tam);
      return f.metodo === 0 ? dec.decode(datos) : new Response(new Blob([datos]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).text();
    };
    const xml = async (n) => new DOMParser().parseFromString((await leer(n)) || "<x/>", "application/xml");
    // hojas del libro en el orden de las pestañas
    const wb = await xml("xl/workbook.xml"), rels = [...(await xml("xl/_rels/workbook.xml.rels")).getElementsByTagName("Relationship")];
    const hojas = [...wb.getElementsByTagName("sheet")].map((h) => {
      const rid = h.getAttribute("r:id") || h.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"), rel = rels.find((r) => r.getAttribute("Id") === rid);
      return { nombre: h.getAttribute("name"), ruta: rel && rel.getAttribute("Target").replace(/^\/?(xl\/)?/, "xl/") };
    }).filter((h) => h.ruta);
    if (!hojas.length) hojas.push(...Object.keys(zip).filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k)).sort().map((ruta) => ({ nombre: ruta, ruta })));
    let ss = null;
    const hoja = async (nombre) => {
      const h = hojas.find((x) => x.nombre === nombre); if (!h) return [];
      ss = ss || [...(await xml("xl/sharedStrings.xml")).getElementsByTagName("si")].map((si) => [...si.getElementsByTagName("t")].filter((t) => t.parentNode.nodeName !== "rPh").map((t) => t.textContent).join(""));
      const txt = await leer(h.ruta);
      // hojas grandes (p. ej. la base de valor en libros, 50 MB): lectura directa del texto, sin armar el árbol XML
      return txt.length > 3e6 ? filasDeTextoXml(txt, ss) : filasDeHoja(new DOMParser().parseFromString(txt || "<x/>", "application/xml"), ss);
    };
    return { nombres: hojas.map((h) => h.nombre), hoja };
  }
  function filasDeTextoXml(xml, ss) {
    const filas = [], des = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
    const colIdx = (ref) => { let n = 0; for (const ch of ref) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
    const reFila = /<row\b([^>]*)>([\s\S]*?)<\/row>/g, reCelda = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = reFila.exec(xml))) {
      const fila = [], nr = +((/\br="(\d+)"/.exec(m[1]) || [])[1] || filas.length + 1);
      let c; reCelda.lastIndex = 0;
      while ((c = reCelda.exec(m[2]))) {
        const ref = (/\br="([A-Z]+)\d+"/.exec(c[1]) || [])[1]; if (!ref || !c[2]) continue;
        const t = (/\bt="(\w+)"/.exec(c[1]) || [])[1], v = /<v>([\s\S]*?)<\/v>/.exec(c[2]);
        const txt = t === "inlineStr" ? des((c[2].match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map((x) => x.replace(/<[^>]+>/g, "")).join("")) : v ? v[1] : "";
        fila[colIdx(ref)] = t === "s" ? ss[+txt] ?? "" : t === "inlineStr" ? txt : t === "str" || t === "e" ? des(txt) : t === "b" ? (txt === "1" ? "VERDADERO" : "FALSO") : txt === "" ? "" : +txt;
      }
      filas[nr - 1] = Array.from(fila, (x) => x ?? "");
    }
    return Array.from(filas, (f) => f || []);
  }
  function filasDeHoja(hoja, ss) {
    const filas = [];
    const colIdx = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, "")) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
    for (const row of hoja.getElementsByTagName("row")) {
      const fila = [];
      for (const c of row.getElementsByTagName("c")) {
        const t = c.getAttribute("t"), v = c.getElementsByTagName("v")[0], txt = v ? v.textContent : "";
        fila[colIdx(c.getAttribute("r") || "A")] = t === "s" ? ss[+txt] ?? "" : t === "inlineStr" ? [...c.getElementsByTagName("t")].map((x) => x.textContent).join("") : t === "b" ? (txt === "1" ? "VERDADERO" : "FALSO") : t === "str" || t === "e" ? txt : txt === "" ? "" : +txt;
      }
      filas[(+row.getAttribute("r") || filas.length + 1) - 1] = Array.from(fila, (x) => x ?? "");
    }
    return Array.from(filas, (f) => f || []);
  }
  // «$1.234,56», «1,234.56», «(123.45)», «-50» → número
  function numeroConc(v) {
    if (typeof v === "number") return v;
    let s = String(v || "").replace(/[$\s]|USD/gi, ""); if (!s) return NaN;
    const neg = /^\(.*\)$/.test(s) || /-$/.test(s); s = s.replace(/[()]/g, "").replace(/-$/, "");
    const uc = s.lastIndexOf(","), up = s.lastIndexOf(".");
    if (uc > -1 && up > -1) s = uc > up ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
    else if (uc > -1) s = /,\d{1,2}$/.test(s) ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
    const n = parseFloat(s); return neg ? -Math.abs(n) : n;
  }
  function fechaConc(v) {
    if (typeof v === "number" && v > 20000 && v < 80000) return new Date(Date.UTC(1899, 11, 30) + v * 864e5).toISOString().slice(0, 10);
    return String(v ?? "").trim();
  }
  // Local de la matriz a partir del texto de la celda: por código (K004, A001…) o por nombre
  const normalTxt = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
  const porCodigoLocal = new Map(locales.map((l) => [(/^([A-Z0-9]+)\s*-/i.exec(l.local) || [])[1]?.toUpperCase(), l]).filter(([k]) => k));
  const porNombreLocal = new Map(locales.map((l) => [normalTxt(l.local.replace(/^[A-Z0-9]+\s*-\s*/i, "")), l]));
  function localDeTexto(txt) {
    const t = normalTxt(txt); if (!t) return null;
    for (const m of t.matchAll(/\b([A-Z]{1,3}\d{2,4})\b/g)) if (porCodigoLocal.has(m[1])) return porCodigoLocal.get(m[1]);
    return porNombreLocal.get(t) || porNombreLocal.get(t.replace(/^[A-Z]{1,3}\d{2,4}\s+/, "")) || null;
  }
  // Fila de encabezados: la primera (de las 30 primeras) con 3+ celdas de texto y alguna palabra conocida
  function filaEncabezado(filas) {
    for (let i = 0; i < Math.min(30, filas.length); i++) {
      const txt = filas[i].filter((c) => typeof c === "string" && c.trim());
      if (txt.length >= 3 && txt.some((c) => CAMPOS_CONC.some(([, , re]) => re.test(c)))) return i;
    }
    return 0;
  }
  let cbPendiente = null; // { archivo, filas, cab, columnas }
  async function subirConciliacion(file) {
    if (!file || !puedeConciliar()) return;
    let filas, libro;
    try { ({ libro, filas } = await leerHojaExcel(file)); }
    catch (err) { return toast(err && err.message === "xls-binario" ? "Este .xls es del formato antiguo: ábrelo en Excel y guárdalo como .xlsx" : "No se pudo leer el archivo"); }
    // Formato de la conciliación mensual: las novedades están en la hoja «ERRORES»
    const hErr = libro && libro.nombres.find((n) => /^\s*errores\s*$/i.test(n));
    if (hErr) {
      toast("Analizando la hoja ERRORES…");
      let r;
      try { r = analizarHojaErrores(await libro.hoja(hErr)); } catch { return toast("No se pudo leer la hoja ERRORES"); }
      // detalle de cada local con novedad, sacado de las hojas de los bancos
      try {
        const porCod = await detallesBancos(libro, /^\s*(errores|base\s+dep)/i);
        r.detalles = {};
        Object.keys(r.porLocal).forEach((l) => { const c = (/^([A-Z0-9]+)\s*-/i.exec(l) || [])[1]; if (c && porCod[c.toUpperCase()]) r.detalles[l] = porCod[c.toUpperCase()]; });
      } catch { /* sin detalle: igual se notifican las novedades */ }
      return guardarConciliacion(r, file.name, `hoja ${hErr} · ${libro.nombres.length} hojas`);
    }
    filas = filas.filter((f) => f.some((c) => String(c).trim() !== ""));
    if (filas.length < 2) return toast("El archivo no tiene filas para analizar");
    const iCab = filaEncabezado(filas), cab = filas[iCab].map((c, i) => String(c || "").trim() || `Columna ${i + 1}`);
    const usadas = new Set(), columnas = {};
    for (const [k, , re] of CAMPOS_CONC) { // «diferencia» manda sobre «valor» si hay ambas
      const cands = cab.map((c, i) => [c, i]).filter(([c, i]) => re.test(c) && !usadas.has(i));
      const elegido = k === "valor" ? cands.find(([c]) => /diferencia|dif\./i.test(c)) || cands[0] : cands[0];
      columnas[k] = elegido ? elegido[1] : -1; if (elegido) usadas.add(elegido[1]);
    }
    if (columnas.local < 0) { // sin encabezado de local: la columna donde más celdas coinciden con un local
      const puntaje = cab.map((_, i) => filas.slice(iCab + 1, iCab + 60).filter((f) => localDeTexto(f[i])).length);
      const mejor = puntaje.indexOf(Math.max(...puntaje)); if (puntaje[mejor] > 0) columnas.local = mejor;
    }
    cbPendiente = { archivo: file.name, filas: filas.slice(iCab + 1), cab, columnas };
    $("#cb-mapeo-archivo").textContent = `${file.name} · ${cbPendiente.filas.length} filas`;
    $("#cb-campos").innerHTML = CAMPOS_CONC.map(([k, et]) => `<label>${et}<select data-cb-col="${k}"><option value="-1">— No está en el archivo —</option>${cab.map((c, i) => `<option value="${i}"${columnas[k] === i ? " selected" : ""}>${esc(c)}</option>`).join("")}</select></label>`).join("");
    $("#cb-previa").innerHTML = `<tr>${cab.map((c) => `<th>${esc(c)}</th>`).join("")}</tr>` + cbPendiente.filas.slice(0, 5).map((f) => `<tr>${cab.map((_, i) => `<td>${esc(f[i] ?? "")}</td>`).join("")}</tr>`).join("");
    $("#cb-mapeo").hidden = false; $("#cb-mapeo").scrollIntoView({ block: "nearest" });
  }
  $("#cb-subir").addEventListener("click", () => $("#cb-archivo").click());
  $("#cb-archivo").addEventListener("change", (e) => { subirConciliacion(e.target.files[0]); e.target.value = ""; });
  $("#cb-cancelar").addEventListener("click", () => { cbPendiente = null; $("#cb-mapeo").hidden = true; });
  // ---- Hoja «ERRORES» de la conciliación mensual ----
  // Cada bloque tiene su título en las primeras filas y debajo una fila por local: «K004 - CCI» y a la derecha el valor
  // y luego (según el bloque) acción, auditor, banco y tipo de depósito.
  //   NOVEDADES BANCARIAS (Quito y Guayaquil): local · diferencia · acción (corregir/desglosar/validar) · auditor · banco · depósito
  //   DEPOSITOS VENTANILLA: zona · local · valor · comentario · auditor
  //   DATAFONO / DEPOSITO CREADO VS. EFECTIVO RC / RETENCION FUENTE: zona · local · valor · auditor
  // En «Novedades bancarias» las diferencias menores a la diferencia mínima (centavos de redondeo) sin acción no se notifican.
  const BLOQUES_CONC = [
    [/novedades\s+bancarias/i, "Novedad bancaria"], [/dep[oó]sitos?\s+ventanilla/i, "Depósito ventanilla"], [/dat[aá]fono/i, "Datáfono"],
    [/creado.*efectivo|efectivo\s+rc/i, "Depósito creado vs. efectivo RC"], [/retenci[oó]n/i, "Retención en la fuente"],
  ];
  const RE_LOCAL_CELDA = /^\s*[A-Z]{1,4}\d{2,3}\s*-\s*\S/i;
  const RE_BANCO = /pichincha|produbanco|internacional|guayaquil|bolivariano|pac[ií]fico|austro|machala|rumi[ñn]ahui/i;
  const RE_TIPO_DEP = /^(cash|tevcol|ventanilla|transportadora|brinks|datafono)/i;
  const RE_PERSONA = /^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,3}$/;
  const difMinima = () => { const v = parseFloat(store.get("matriz:conciliacion:minimo", 1)); return Number.isFinite(v) && v >= 0 ? v : 1; };
  function analizarHojaErrores(filas) {
    const titulos = [];
    filas.slice(0, 10).forEach((f) => f.forEach((c, col) => { const b = BLOQUES_CONC.find(([re]) => re.test(String(c))); if (b && !titulos.some((t) => t.tipo === b[1])) titulos.push({ col, tipo: b[1] }); }));
    if (!titulos.length) throw new Error("sin bloques");
    titulos.sort((a, b) => a.col - b.col);
    const minimo = difMinima(), porLocal = {}, sinLocal = new Map(), conteo = {}; let n = 0, revisadas = 0;
    filas.forEach((f, r) => f.forEach((celda, col) => {
      if (!RE_LOCAL_CELDA.test(String(celda))) return;
      const t = [...titulos].reverse().find((x) => x.col <= col); if (!t || col - t.col > 6) return; // fuera de los bloques (p. ej. el listado Local/Auditor)
      const valor = numeroConc(f[col + 1]); if (!Number.isFinite(valor)) return;
      const sig = titulos[titulos.indexOf(t) + 1], extras = { banco: "", doc: "", auditor: "", novedad: [] };
      for (let k = col + 2; k < col + 7 && (!sig || k < sig.col); k++) {
        const v = String(f[k] ?? "").trim(); if (!v) continue; // la acción puede venir vacía
        if (RE_LOCAL_CELDA.test(v)) break;
        if (RE_BANCO.test(v) && !extras.banco) extras.banco = v;
        else if (RE_TIPO_DEP.test(v) && !extras.doc) extras.doc = v;
        else if (RE_PERSONA.test(v) && !extras.auditor) extras.auditor = v;
        else extras.novedad.push(v);
      }
      revisadas++;
      const v2 = Math.round(valor * 100) / 100, accion = extras.novedad.join(" · ");
      if (t.tipo === "Novedad bancaria" ? Math.abs(v2) < minimo && !accion : v2 === 0 && !accion) return;
      const l = localDeTexto(celda);
      if (!l) { const k = String(celda).trim(); sinLocal.set(k, (sinLocal.get(k) || 0) + 1); return; }
      (porLocal[l.local] = porLocal[l.local] || []).push({ id: `e${r}-${col}`, tipo: t.tipo, fecha: "", banco: extras.banco, doc: extras.doc, valor: v2, novedad: accion, auditor: extras.auditor });
      conteo[t.tipo] = (conteo[t.tipo] || 0) + 1; n++;
    }));
    return { porLocal, sinLocal: [...sinLocal].map(([texto, c]) => ({ texto, n: c })), filas: revisadas, n, conteo, minimo };
  }
  // ---- Detalle de las hojas de cada banco ----
  // Cada hoja trae el estado de cuenta por local: una fila con solo el código del local («M005»), luego los depósitos
  // (fecha, número, valor… y al final su número de secuencia) y cierra con MES / TRANSITO / TOTAL.
  // Lo que conciliación anota a la derecha de la secuencia («desglosar», «no se encuentra acreditado», el valor del sistema…)
  // es el detalle de la novedad. También se muestran los depósitos con la misma secuencia que uno anotado (p. ej. un desglose).
  const BANCO_HOJA = [[/pichi/i, "Pichincha"], [/prod/i, "Produbanco"], [/pacif/i, "Pacífico"], [/gye|guay/i, "Guayaquil"], [/boli/i, "Bolivariano"], [/inter/i, "Internacional"]];
  const RE_COD_SOLO = /^'?([A-Z]{1,4}\d{2,3})$/i;
  const esFechaCelda = (v) => (typeof v === "number" && v > 40000 && v < 60000) || /^\d{4}-\d{2}-\d{2}|^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(v));
  const txtFecha = (v) => { if (typeof v === "number") return fechaConc(v); const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(v)); return m ? `${m[3]}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}` : String(v).slice(0, 10); };
  async function detallesBancos(libro, excluir) {
    const porCodigo = {};
    for (const nombre of libro.nombres.filter((n) => !excluir.test(n))) {
      const banco = (BANCO_HOJA.find(([re]) => re.test(nombre)) || [])[1] || "";
      let filas; try { filas = await libro.hoja(nombre); } catch { continue; }
      // columna de la secuencia: la última celda más común de las filas de depósito
      const ultimas = {}, bloques = []; let cod = null;
      filas.forEach((f, r) => {
        const llenas = f.map((c, i) => [c, i]).filter(([c]) => String(c).trim() !== "");
        if (llenas.length === 1 && RE_COD_SOLO.test(String(llenas[0][0]).trim())) { cod = RE_COD_SOLO.exec(String(llenas[0][0]).trim())[1].toUpperCase(); return; }
        if (llenas.some(([c]) => /^total$/i.test(String(c).trim()))) { cod = null; return; }
        if (cod && llenas.length >= 4 && esFechaCelda(f[0])) { bloques.push({ cod, f, r }); const u = llenas[llenas.length - 1][1]; ultimas[u] = (ultimas[u] || 0) + 1; }
      });
      if (!bloques.length) continue;
      const base = +Object.entries(ultimas).sort((a, b) => b[1] - a[1])[0][0]; // índice de la secuencia
      const notaDe = (f) => {
        const cs = f.slice(base + 1).filter((c) => String(c).trim() !== ""), usd = (c) => fmtUSD(Math.round(c * 100) / 100);
        // dos números seguidos = diferencia y valor que tiene el sistema (p. ej. «-0,96 · 876,96»)
        if (typeof cs[0] === "number" && !esFechaCelda(cs[0]) && typeof cs[1] === "number" && !esFechaCelda(cs[1]))
          return [`Diferencia ${usd(cs[0])} · en el sistema ${usd(cs[1])}`, ...cs.slice(2).map(String)].join(" · ");
        return cs.map((c) => (esFechaCelda(c) ? `fecha ${txtFecha(c)}` : typeof c === "number" ? usd(c) : String(c).trim())).join(" · ");
      };
      const anotadas = bloques.filter((b) => notaDe(b.f)), claves = new Set(anotadas.map((b) => `${b.cod}|${b.f[base]}`));
      bloques.filter((b) => notaDe(b.f) || claves.has(`${b.cod}|${b.f[base]}`)).forEach((b) => {
        const f = b.f.slice(0, base), nums = f.map((c, i) => [c, i]).filter(([c, i]) => i > 0 && typeof c === "number" && Math.abs(c) < 1e6 && c !== 0);
        const doc = (f.map((c) => String(c).replace(/^'/, "")).filter((c) => /^\d{6,}$/.test(c)).sort((a, b) => b.length - a.length)[0]) || "";
        const concepto = f.filter((c) => typeof c === "string" && /[a-z]{4,}/i.test(c) && !/^(deposito|re-ef|dp|c)$/i.test(c.trim()) && !RE_COD_SOLO.test(c.trim())).join(" ");
        (porCodigo[b.cod] = porCodigo[b.cod] || []).push({ hoja: nombre, banco, fecha: txtFecha(b.f[0]), doc, valor: nums.length ? Math.round(nums[nums.length - 1][0] * 100) / 100 : null,
          sec: b.f[base], concepto, nota: notaDe(b.f) || `misma secuencia (${b.f[base]}) que el depósito anotado` });
      });
    }
    return porCodigo;
  }
  // Qué debe hacer el auditor con cada novedad (según la acción que puso conciliación y el tipo)
  function queHacerConc(x) {
    const a = String(x.novedad || "").toLowerCase(), v = x.valor == null ? "" : fmtUSD(Math.abs(x.valor));
    if (/desglos/.test(a)) return "Desglosar el depósito: en el banco el valor aparece en varias papeletas (o agrupado en una). Registra en el sistema cada depósito por separado, con su número y valor, como está en el estado de cuenta.";
    if (/corregir/.test(a)) return "Corregir en el sistema el depósito que no coincide con el estado de cuenta (número de papeleta, fecha o valor) y enviar el respaldo a conciliación.";
    if (/validar/.test(a)) return "Validar con el local los respaldos del depósito (papeleta, guía de la transportadora o comprobante) y enviarlos a conciliación.";
    if (/eliminar/.test(a)) return `Seguir la indicación de conciliación: ${x.novedad}.`;
    switch (x.tipo) {
      case "Depósito ventanilla": return `Depósito por ventanilla de ${v}${/comunic/.test(a) ? " (ya comunicado al local)" : ""}: confirmar con el local la papeleta y que el valor esté acreditado en el banco; si no, regularizar.`;
      case "Datáfono": return `Diferencia de ${v} en datáfono: revisar los cierres de lote del datáfono del local contra lo acreditado por el banco.`;
      case "Depósito creado vs. efectivo RC": return `El depósito creado en el sistema no coincide con el efectivo del Reporte de Caja (RC) por ${v}: revisar el cierre de caja del día y corregir el depósito.`;
      case "Retención en la fuente": return `Diferencia de ${v} en retenciones en la fuente: verificar los comprobantes de retención recibidos y su registro.`;
      case "Novedad bancaria": return x.valor < 0 ? `En el banco hay ${v} menos que en el sistema: confirmar si el depósito se hizo y está acreditado; si falta, regularizar el faltante con el local.`
        : x.valor > 0 ? `En el banco hay ${v} más que en el sistema: identificar el depósito que no está registrado (o registrado por menos) y regularizarlo.` : "Revisar el depósito con conciliación.";
      default: return "";
    }
  }
  // Tabla «Detalle en el estado de cuenta» de un local (correo = estilos en línea para Outlook)
  function htmlDetalleConc(lista, correo) {
    if (!lista || !lista.length) return "";
    const s = (css) => (correo ? ` style="${css}"` : ""), td = (t, x = "") => `<td${s(`padding:4px 8px;border-bottom:1px solid #dfe3e8;font-size:9.5pt;${x}`)}>${t}</td>`;
    const th = (t) => `<th${s("background:#fff3b0;padding:4px 8px;font-size:9pt;text-align:left")}>${t}</th>`;
    const hojas = [...new Set(lista.map((x) => x.hoja))].join(", ");
    return `<div class="cb-detalle"${s("margin:8px 0 4px")}><p${s("margin:6px 0 4px;font-size:10pt")}><b>Detalle en el estado de cuenta</b> <small class="muted">(hoja ${esc(hojas)})</small></p>
      ${correo ? "" : `<div class="tabla-scroll">`}<table class="na-tabla"${s("border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif")}><tr>${th("Banco")}${th("Fecha")}${th("N° depósito")}${th("Valor en banco")}${th("Sec.")}${th("Nota de conciliación")}</tr>
      ${lista.map((x) => `<tr>${td(esc(x.banco))}${td(esc(x.fecha), "white-space:nowrap")}${td(esc(x.doc || x.concepto || "—"))}${td(x.valor == null ? "—" : fmtUSD(x.valor), "text-align:right;white-space:nowrap")}${td(esc(x.sec))}${td(`<b>${esc(x.nota)}</b>`)}</tr>`).join("")}</table>${correo ? "" : "</div>"}</div>`;
  }
  function guardarConciliacion(r, archivo, detalle) {
    const datos = { archivo, detalle, cargado: new Date().toISOString(), por: usuario.nombre, filas: r.filas, porLocal: r.porLocal, sinLocal: r.sinLocal, conteo: r.conteo || null, minimo: r.minimo, detalles: r.detalles || {} };
    store.set(CLAVE_CONC, datos); cbPendiente = null; $("#cb-mapeo").hidden = true;
    toast(`${r.n} novedad(es) en ${Object.keys(r.porLocal).length} local(es)${r.sinLocal.length ? ` · ${r.sinLocal.length} local(es) del archivo no se encontraron en la matriz` : ""}`);
    renderConciliacion();
  }
  // Novedad = fila de un local con diferencia distinta de cero, o con observación que no diga que está conciliado
  const SIN_NOVEDAD = /^(ok|conciliad[oa]s?|cuadrad[oa]|sin novedad|correcto|cerrad[oa]|n\/?a|-|0)$/i;
  $("#cb-analizar").addEventListener("click", () => {
    if (!cbPendiente) return;
    const col = Object.fromEntries($$("[data-cb-col]").map((s) => [s.dataset.cbCol, +s.value]));
    if (col.local < 0) return toast("Elige la columna del local");
    if (col.valor < 0 && col.novedad < 0) return toast("Elige la columna del valor o de la novedad");
    const porLocal = {}, sinLocal = new Map(); let n = 0;
    cbPendiente.filas.forEach((f, i) => {
      const txtLocal = String(f[col.local] ?? "").trim(); if (!txtLocal || /^total/i.test(txtLocal)) return;
      const valor = col.valor >= 0 ? numeroConc(f[col.valor]) : NaN, nov = col.novedad >= 0 ? String(f[col.novedad] ?? "").trim() : "";
      const conDif = Number.isFinite(valor) && Math.abs(valor) >= 0.01, conNov = nov && !SIN_NOVEDAD.test(nov);
      if (!conDif && !conNov) return;
      const l = localDeTexto(txtLocal);
      if (!l) { sinLocal.set(txtLocal, (sinLocal.get(txtLocal) || 0) + 1); return; }
      (porLocal[l.local] = porLocal[l.local] || []).push({ id: `f${i}`, tipo: "Conciliación", fecha: col.fecha >= 0 ? fechaConc(f[col.fecha]) : "", banco: col.banco >= 0 ? String(f[col.banco] ?? "").trim() : "",
        doc: col.doc >= 0 ? String(f[col.doc] ?? "").trim() : "", valor: Number.isFinite(valor) ? Math.round(valor * 100) / 100 : null, novedad: nov, auditor: "" });
      n++;
    });
    guardarConciliacion({ porLocal, sinLocal: [...sinLocal].map(([texto, c]) => ({ texto, n: c })), filas: cbPendiente.filas.length, n }, cbPendiente.archivo, "columnas elegidas a mano");
  });
  // Revisión del auditor por novedad
  const idConc = (local, x) => `${local}|${x.id}|${x.doc}|${x.valor}`;
  const ESTADOS_CONC = ["Pendiente", "Revisado", "Justificado"];
  function renderConciliacion() {
    const d = datosConciliacion(), concilia = puedeConciliar();
    $("#cb-subir").hidden = !concilia; $("#cb-minimo-box").hidden = !concilia; $("#cb-publicar").hidden = !concilia || !d; $("#cb-correo-todos").hidden = !concilia || !d; $("#cb-cargar").hidden = concilia;
    if (!d) {
      $("#cb-info").textContent = concilia ? "Sube el Excel de conciliación bancaria: la app encuentra el local de cada fila y deja las novedades listas para los auditores." : "Aún no hay conciliación bancaria cargada.";
      ["#cb-kpis", "#cb-resumen", "#cb-lista"].forEach((s) => ($(s).innerHTML = "")); $("#cb-auditor").innerHTML = ""; return;
    }
    const rev = concilia ? {} : store.get(claveRevision(), {}), estadoDe = (local, x) => (rev[idConc(local, x)] || {}).estado || "Pendiente";
    // auditores: solo sus locales (y el local que tengan seleccionado)
    const mios = (l, xs) => concilia || !usuario.auditor || esDelAuditor(l, xs) || (localActual && localActual.local === l);
    const todos = Object.entries(d.porLocal).filter(([l, xs]) => mios(l, xs)).map(([l, xs]) => ({ local: l, info: localPorNombre.get(l) || {}, resp: responsableConc(l, xs), xs }));
    const auditores = [...new Set(todos.map((g) => g.resp))].sort(cmpTxt);
    const sel = $("#cb-auditor"), v = sel.value;
    sel.hidden = !concilia; sel.innerHTML = `<option value="">Todos los auditores</option>` + auditores.map((a) => `<option>${esc(a)}</option>`).join(""); sel.value = auditores.includes(v) ? v : "";
    const q = $("#cb-buscar").value.trim().toLowerCase(), est = $("#cb-estado").value, aud = sel.value;
    const tipoSel = $("#cb-tipo").value;
    const grupos = todos.filter((g) => !aud || g.resp === aud)
      .map((g) => ({ ...g, xs: g.xs.filter((x) => (!est || estadoDe(g.local, x) === est) && (!tipoSel || x.tipo === tipoSel) && (!q || [g.local, x.tipo, x.banco, x.doc, x.novedad, x.fecha, g.resp].join(" ").toLowerCase().includes(q))) }))
      .filter((g) => g.xs.length)
      .sort((a, b) => (localActual && b.local === localActual.local) - (localActual && a.local === localActual.local) || cmpTxt(a.local, b.local));
    const nTot = todos.reduce((s, g) => s + g.xs.length, 0), suma = todos.reduce((s, g) => s + g.xs.reduce((t, x) => t + (x.valor || 0), 0), 0);
    const nPend = concilia ? 0 : todos.reduce((s, g) => s + g.xs.filter((x) => estadoDe(g.local, x) === "Pendiente").length, 0);
    const tipos = [...new Set(todos.flatMap((g) => g.xs.map((x) => x.tipo || "Conciliación")))];
    const tsel = $("#cb-tipo"); tsel.innerHTML = `<option value="">Todos los tipos</option>` + tipos.map((t) => `<option${t === tipoSel ? " selected" : ""}>${esc(t)}</option>`).join("");
    tsel.hidden = tipos.length < 2;
    $("#cb-info").textContent = `${d.archivo} · cargado el ${d.cargado.slice(0, 10)}${d.por ? ` por ${d.por}` : ""}` + (concilia ? ` · ${d.detalle ? `${d.detalle} · ` : ""}${d.filas} filas revisadas${d.minimo != null ? ` · diferencia mínima ${fmtUSD(d.minimo)}` : ""}` : usuario.auditor ? " · tus locales" : "");
    const porTipo = tipos.map((t) => [todos.reduce((s, g) => s + g.xs.filter((x) => (x.tipo || "Conciliación") === t).length, 0), t]);
    $("#cb-kpis").innerHTML = [[nTot, "Novedades"], [todos.length, "Locales con novedad"], ...(tipos.length > 1 ? porTipo : [[fmtUSD(suma), "Diferencia total"]]), concilia ? [d.sinLocal.length, "Locales no encontrados"] : [nPend, "Por revisar"]]
      .map(([n, t]) => `<div><b>${n}</b><span>${esc(t)}</span></div>`).join("");
    // resumen por auditor con su correo (usuario de Conciliación)
    $("#cb-resumen").innerHTML = concilia ? `<div class="tabla-scroll"><table class="na-tabla cb-res"><thead><tr><th>Auditor</th><th>Locales</th><th>Novedades</th><th>Valor</th><th></th></tr></thead><tbody>
      ${auditores.map((a) => { const gs = todos.filter((g) => g.resp === a), nx = gs.reduce((s, g) => s + g.xs.length, 0);
        return `<tr><td><b>${esc(a)}</b></td><td>${gs.length}</td><td>${nx}</td><td class="nowrap">${fmtUSD(gs.reduce((s, g) => s + g.xs.reduce((t, x) => t + (x.valor || 0), 0), 0))}</td>
          <td><button type="button" class="ghost sm" data-cb-correo="${esc(a)}">✉ Correo</button></td></tr>`; }).join("")}</tbody></table></div>
      ${d.sinLocal.length ? `<p class="aviso-admin small">No se encontraron en la matriz: ${d.sinLocal.map((s) => `${esc(s.texto)} (${s.n})`).join(", ")}. Revisa que el código del local esté en el archivo (ej. K004).</p>` : ""}` : "";
    $("#cb-lista").innerHTML = grupos.length ? grupos.map((g) => `<div class="na-grupo${localActual && localActual.local === g.local ? " cb-actual" : ""}">
      <h3>${esc(g.local)} <small>${esc(g.resp)} · ${g.xs.length} novedad(es)${localActual && localActual.local === g.local ? " · local seleccionado" : ""}</small></h3>
      <div class="tabla-scroll"><table class="na-tabla"><thead><tr><th>Tipo</th><th>Banco</th><th>Depósito</th><th>Valor</th><th>Acción / observación</th>${concilia ? "" : "<th>Revisión</th>"}</tr></thead><tbody>
      ${g.xs.map((x) => { const r = rev[idConc(g.local, x)] || {};
        return `<tr data-cb="${esc(idConc(g.local, x))}" class="${estadoDe(g.local, x) !== "Pendiente" ? "hecho" : ""}"><td><span class="na-tipo ${claseTipoConc(x.tipo)}">${esc(x.tipo || "Conciliación")}</span>${x.fecha ? `<small>${esc(x.fecha)}</small>` : ""}</td><td>${esc(x.banco)}</td><td>${esc(x.doc)}</td>
        <td class="nowrap ${x.valor < 0 ? "neg" : ""}">${x.valor == null ? "—" : fmtUSD(x.valor)}</td><td class="cb-accion"><b>${esc(x.novedad ? x.novedad.charAt(0).toUpperCase() + x.novedad.slice(1) : x.tipo === "Novedad bancaria" ? "Diferencia en conciliación" : "—")}</b>${queHacerConc(x) ? `<div class="cb-hacer">👉 ${esc(queHacerConc(x))}</div>` : ""}</td>
        ${concilia ? "" : `<td><select data-cb-estado aria-label="Revisión">${ESTADOS_CONC.map((e) => `<option${e === estadoDe(g.local, x) ? " selected" : ""}>${e}</option>`).join("")}</select>
          <input data-cb-obs value="${esc(r.obs || "")}" placeholder="Observación del auditor"></td>`}</tr>`; }).join("")}
      </tbody></table></div>${htmlDetalleConc((d.detalles || {})[g.local])}</div>`).join("")
      : `<p class="muted">${nTot ? "Ninguna novedad coincide con los filtros." : concilia ? "El archivo no tiene novedades." : "Tus locales no tienen novedades bancarias en esta conciliación."}</p>`;
  }
  ["#cb-buscar", "#cb-estado", "#cb-auditor", "#cb-tipo"].forEach((s) => $(s).addEventListener(s === "#cb-buscar" ? "input" : "change", renderConciliacion));
  const claseTipoConc = (t) => ({ "Novedad bancaria": "perdido", "Depósito ventanilla": "baja", "Datáfono": "modificar", "Retención en la fuente": "crear" }[t] || "otra");
  // Diferencia mínima para «Novedades bancarias» (se aplica al subir el Excel)
  $("#cb-minimo").value = difMinima();
  $("#cb-minimo").addEventListener("change", () => { const v = parseFloat($("#cb-minimo").value); store.set("matriz:conciliacion:minimo", Number.isFinite(v) && v >= 0 ? v : 1); $("#cb-minimo").value = difMinima(); toast("Se aplica la próxima vez que subas el Excel"); });
  $("#cb-lista").addEventListener("change", (e) => {
    const tr = e.target.closest("[data-cb]"); if (!tr || puedeConciliar()) return;
    const r = store.get(claveRevision(), {}), x = r[tr.dataset.cb] = r[tr.dataset.cb] || {};
    if (e.target.matches("[data-cb-estado]")) x.estado = e.target.value;
    if (e.target.matches("[data-cb-obs]")) x.obs = e.target.value.trim();
    store.set(claveRevision(), r);
    if (e.target.matches("[data-cb-estado]")) renderConciliacion();
  });
  // Correo con las novedades (de un auditor o de todos) y el archivo adjunto para cargarlo en la app
  function correoConciliacion(auditor) {
    const d = datosConciliacion(); if (!d) return;
    const porLocal = Object.fromEntries(Object.entries(d.porLocal).filter(([l, xs]) => !auditor || responsableConc(l, xs) === auditor));
    if (!Object.keys(porLocal).length) return toast("No hay novedades para ese auditor");
    const b64 = (s) => { const b = new TextEncoder().encode(s); let t = ""; for (let i = 0; i < b.length; i += 8192) t += String.fromCharCode(...b.subarray(i, i + 8192)); return btoa(t); };
    const lineas = (s) => s.replace(/.{76}/g, "$&\r\n"), lim = "----=_conc_" + Date.now(), hoy = new Date().toISOString().slice(0, 10);
    const th = (t) => `<th style="background:#ffc629;padding:5px 8px;font-size:9pt;text-align:left">${t}</th>`, td = (t, x = "") => `<td style="padding:5px 8px;border-bottom:1px solid #dfe3e8;font-size:10pt;${x}">${t}</td>`;
    const tablas = Object.entries(porLocal).sort(([a], [b]) => cmpTxt(a, b)).map(([l, xs]) => `<p style="margin:14px 0 4px"><b>${esc(l)}</b>${auditor ? "" : ` · ${esc(responsableConc(l, xs))}`}</p>
      <table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif"><tr>${th("Tipo")}${th("Banco")}${th("Depósito")}${th("Valor")}${th("Acción / qué hacer")}</tr>
      ${xs.map((x) => `<tr>${td(esc([x.tipo || "Conciliación", x.fecha].filter(Boolean).join(" · ")))}${td(esc(x.banco))}${td(esc(x.doc))}${td(x.valor == null ? "—" : fmtUSD(x.valor), `text-align:right;white-space:nowrap;${x.valor < 0 ? "color:#b42318" : ""}`)}${td(`<b>${esc(x.novedad || (x.tipo === "Novedad bancaria" ? "Diferencia en conciliación" : "—"))}</b>${queHacerConc(x) ? `<br><span style="color:#3d3000">👉 ${esc(queHacerConc(x))}</span>` : ""}`, "max-width:420px")}</tr>`).join("")}</table>
      ${htmlDetalleConc((d.detalles || {})[l], true)}`).join("");
    const n = Object.values(porLocal).reduce((s, xs) => s + xs.length, 0);
    const html = `<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:11pt">${auditor ? `Estimado(a) ${esc(auditor)}` : "Estimados auditores"},<br><br>
      Se envían las novedades bancarias de la conciliación (${n} novedad(es) en ${Object.keys(porLocal).length} local(es)) para su revisión en los locales.
      También están en la pestaña <b>Conciliación bancaria</b> de la Matriz de Auditoría (si no aparecen, carga el archivo adjunto con «Cargar archivo recibido»).${tablas}
      <br>Saludos cordiales,<br>${esc(usuario.nombre)}<br>Auditoría · Grupo KFC Ecuador</body></html>`;
    const nom = `Conciliacion bancaria ${hoy}.json`, json = JSON.stringify({ tipo: "matriz-conciliacion", v: 1, ...d, porLocal });
    const asunto = `Novedades de conciliación bancaria – ${auditor || "Auditores"} – ${hoy}`;
    const eml = [`Subject: =?UTF-8?B?${b64(asunto)}?=`, "X-Unsent: 1", "MIME-Version: 1.0", `Content-Type: multipart/mixed; boundary="${lim}"`, "",
      `--${lim}`, "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: base64", "", lineas(b64(html)),
      `--${lim}`, `Content-Type: application/json; name="${nom}"`, "Content-Transfer-Encoding: base64", `Content-Disposition: attachment; filename="${nom}"`, "", lineas(b64(json)),
      `--${lim}--`, ""].join("\r\n");
    descargar(new Blob([eml], { type: "message/rfc822" }), `${asunto}.eml`.replace(/[\\/:*?"<>|]/g, "_"));
    toast("Borrador descargado: ábrelo en Outlook, pon el correo del auditor y envíalo");
  }
  $("#cb-resumen").addEventListener("click", (e) => { const b = e.target.closest("[data-cb-correo]"); if (b) correoConciliacion(b.dataset.cbCorreo); });
  $("#cb-correo-todos").addEventListener("click", () => correoConciliacion(""));
  // Publicar en la carpeta de la app: la ven todos los auditores que abren esa carpeta (compartida, OneDrive, red…)
  $("#cb-publicar").addEventListener("click", async () => {
    const d = datosConciliacion(); if (!d || !puedeConciliar()) return;
    if (!window.showDirectoryPicker) return toast("Este navegador no permite escribir en carpetas. Abre la app en Microsoft Edge o Google Chrome.");
    let dir;
    try { dir = await window.showDirectoryPicker({ id: "matriz-app", mode: "readwrite" }); } catch { return; }
    try { await dir.getFileHandle("index.html"); }
    catch { if (!confirm("Esta carpeta no tiene index.html: no parece ser la carpeta de la app. ¿Publicar aquí de todos modos?")) return; }
    try {
      const cdir = await dir.getDirectoryHandle("conciliacion", { create: true }), fh = await cdir.getFileHandle("conciliacion.js", { create: true }), w = await fh.createWritable();
      await w.write(`// Conciliación bancaria publicada desde la app (${usuario.nombre}). No editar a mano.\nwindow.CONCILIACION = ${JSON.stringify(d)};\n`); await w.close();
      window.CONCILIACION = d; toast("Publicado: los auditores la verán en la pestaña Conciliación bancaria");
    } catch { toast("No se pudo publicar: revisa los permisos de la carpeta"); }
  });
  // Auditores: cargar el correo / archivo recibido (solo si es más reciente que lo que ya tienen)
  $("#cb-cargar").addEventListener("click", () => $("#cb-recibido").click());
  $("#cb-recibido").addEventListener("change", async (e) => {
    const file = e.target.files[0]; e.target.value = ""; if (!file) return;
    try {
      let txt = await file.text();
      if (!txt.trim().startsWith("{")) {
        const m = /Content-Type: application\/json[^]*?\r?\n\r?\n([A-Za-z0-9+/=\r\n]+)/i.exec(txt); if (!m) throw new Error("sin adjunto");
        txt = new TextDecoder().decode(Uint8Array.from(atob(m[1].replace(/\s+/g, "")), (ch) => ch.charCodeAt(0)));
      }
      const d = JSON.parse(txt); if (d.tipo !== "matriz-conciliacion" || !d.porLocal) throw new Error("no es conciliación");
      const actual = datosConciliacion();
      if (actual && String(actual.cargado) > String(d.cargado)) return toast("Ya tienes una conciliación más reciente");
      delete d.tipo; delete d.v; store.set(CLAVE_CONC, d);
      toast(`Conciliación cargada: ${Object.values(d.porLocal).reduce((s, xs) => s + xs.length, 0)} novedad(es)`); renderConciliacion();
    } catch { toast("No reconozco el archivo: debe ser el correo o el .json de conciliación bancaria"); }
  });

  // ---------- Inicio ----------
  const previo = sess.get("matriz:usuario");
  const uPrevio = previo && USUARIOS.find((u) => u.usuario === previo);
  if (uPrevio) entrar(uPrevio);
})();
