import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { loadData, saveData } from "./db";

// ── Tokens de estilo ─────────────────────────────────────
const W = "#ffffff", BG = "#f5f5f5", BR = "#dddddd", T = "#111111", T2 = "#555555";
const card  = { background: W, border: `1px solid ${BR}`, borderRadius: 12, padding: "1rem" };
const card2 = { background: W, border: `1px solid ${BR}`, borderRadius: 12, padding: "0.75rem 1rem" };
const kpi   = { background: BG, border: `1px solid ${BR}`, borderRadius: 8, padding: "0.75rem", textAlign: "center" };
const TBL   = { width: "100%", borderCollapse: "collapse", fontSize: 13, background: W };
const TH    = { textAlign: "left", padding: "6px 8px", fontWeight: 500, color: T2, borderBottom: `1px solid ${BR}`, background: W };
const TD    = { padding: "6px 8px", borderBottom: `1px solid ${BR}`, color: T };

// ── Clave de almacenamiento persistente ──────────────────
const STORAGE_KEY = "flota_data_v1";

// ── Datos iniciales ──────────────────────────────────────
const INIT = {
  vehiculos:       [],
  choferes:        [],
  bancos:          [],
  cuentas:         [],
  inversionistas:  [],
  deudas:          [],
  pagos:           [],
  pagosInv:        [],
  gastosInv:       [],
  usuarios: [
    { id:1, nombre:"Administrador", email:"admin@flota.com", password:"admin123", rol:"administrador", activo:true },
  ],
};

// ── Constantes ───────────────────────────────────────────
const FORMAS_PAGO = [
  { value:"5_ls",      label:"$5 lunes–sábado" },
  { value:"30_ls",     label:"$30 lunes–sábado" },
  { value:"30_dom",    label:"$30 solo domingos" },
  { value:"5ls_30dom", label:"$5 L–S + $30 domingos" },
  { value:"5ls_20dom", label:"$5 L–S + $20 domingos" },
  { value:"manual",    label:"Personalizado (elegir fecha y monto)" },
];

// ── Funciones utilitarias ────────────────────────────────
const calcMontoDia = (d, fecha) => {
  const fp = typeof d === "string" ? d : d?.formaPago;   // compatibilidad: acepta deuda u objeto-formaPago
  const dia = new Date(fecha + "T12:00:00").getDay();
  if (fp === "5_ls")      return dia >= 1 && dia <= 6 ? 5 : 0;
  if (fp === "30_ls")     return dia >= 1 && dia <= 6 ? 30 : 0;
  if (fp === "30_dom")    return dia === 0 ? 30 : 0;
  if (fp === "5ls_30dom") return dia === 0 ? 30 : dia >= 1 && dia <= 6 ? 5 : 0;
  if (fp === "5ls_20dom") return dia === 0 ? 20 : dia >= 1 && dia <= 6 ? 5 : 0;
  if (fp === "manual") {
    if (typeof d !== "object") return 0;
    return d.fechaManual === fecha ? (Number(d.montoManual) || 0) : 0;
  }
  return 0;
};

// Fecha de hoy en horario de Ecuador (América/Guayaquil, UTC-5)
const hoy     = () => new Intl.DateTimeFormat("en-CA", { timeZone:"America/Guayaquil", year:"numeric", month:"2-digit", day:"2-digit" }).format(new Date());
const diasVig = (f) => { if (!f) return 999; const o = new Date(f + "T12:00:00Z"); if (isNaN(o.getTime())) return 999; const h = new Date(hoy() + "T12:00:00Z"); return Math.round((o - h) / 86400000); };
const vacio   = (obj, keys) => keys.filter(k => !obj[k] || !String(obj[k]).trim());
const imMens  = (m) => (Number(m) * 0.15) / 12;
const capAnual= (m) => Number(m) * 0.20;
const round2  = (n) => Math.round(n * 100) / 100;
// Parseo seguro de fechas: devuelve null si la fecha es inválida o vacía
const parseF  = (f) => { if (!f) return null; const d = new Date(f + "T12:00:00"); return isNaN(d.getTime()) ? null : d; };
// Formatea con los componentes LOCALES (no UTC) para que parseF→isoF sea idéntico en cualquier zona horaria
const isoF    = (d) => {
  if (!d || isNaN(d.getTime())) return "—";
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, "0"), day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};

// Enter pasa al siguiente campo del formulario (input/select/textarea visible)
const focusSiguiente = (e) => {
  if (e.key !== "Enter") return;
  const el = e.target;
  if (el.tagName === "TEXTAREA") return;
  e.preventDefault();
  const scope = el.closest("[data-form-scope]") || document;
  const campos = Array.from(scope.querySelectorAll("input, select, textarea"))
    .filter(f => !f.disabled && f.type !== "hidden" && f.getClientRects().length > 0);
  const i = campos.indexOf(el);
  if (i > -1 && i + 1 < campos.length) campos[i + 1].focus();
};

// Formatea entrada de hora a HH:MM insertando ":" automáticamente (teclado numérico en móvil)
const formatHora = (v) => {
  const d = String(v).replace(/\D/g, "").slice(0, 4);
  return d.length <= 2 ? d : d.slice(0, 2) + ":" + d.slice(2);
};
// Valida hora real en formato HH:MM (00:00 a 23:59)
const horaValida = (v) => /^([01]\d|2[0-3]):([0-5]\d)$/.test(v || "");

// ── Exportación a Excel ──────────────────────────────────
const descargarLibro = (sheets, filename) => {
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => {
    const filas = (s.filas && s.filas.length) ? s.filas : [{ "Sin registros": "—" }];
    const ws = XLSX.utils.json_to_sheet(filas);
    XLSX.utils.book_append_sheet(wb, ws, s.nombre.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
};

// ── Guardas de integridad referencial ────────────────────
const choferTieneDeudaActiva = (data, cid) => data.deudas.some(d => String(d.choferId) === String(cid) && d.activa);
const choferTienePagos       = (data, cid) => data.pagos.some(p => String(p.choferId) === String(cid));
const vehiculoTieneDeudas    = (data, vid) => data.deudas.some(d => String(d.vehiculoId) === String(vid));
const vehiculoDeudaActiva    = (data, vid) => data.deudas.find(d => String(d.vehiculoId) === String(vid) && d.activa);
const bancoEnUso             = (data, nombre) => data.pagos.some(p => p.banco === nombre);
const cuentaEnUso            = (data, numero) => data.pagos.some(p => p.cuentaDestino === numero);
const invEnUso               = (data, id) => data.gastosInv.some(g => g.invId === id) || data.pagosInv.some(p => p.invId === id);
const comprobanteDuplicado   = (data, numero, banco, fechaComp, hora) =>
  data.pagos.some(p => p.comprobante === numero && p.banco === banco && p.fechaComp === fechaComp && p.hora === hora);

// ── Cálculo de saldo por día (considera abonos parciales) ─
const montoRequerido = (data, cid, f) =>
  data.deudas.filter(d => String(d.choferId)===String(cid) && d.activa && d.fechaInicio<=f && d.fechaFin>=f)
    .reduce((s,d) => s + calcMontoDia(d, f), 0);

// Devuelve lo requerido, lo abonado y el saldo restante de un día concreto
const estadoDia = (data, cid, f) => {
  const req = montoRequerido(data, cid, f);
  const pgs = data.pagos.filter(p => String(p.choferId)===String(cid) && p.fecha === f);
  if (pgs.some(p => p.estado === "condonado")) return { req, pagado: req, restante: 0, condonado: true };
  const pagado = pgs.filter(p => p.estado === "pagado" || p.estado === "abono").reduce((s,p) => s + Number(p.monto||0), 0);
  return { req, pagado: round2(pagado), restante: round2(req - pagado), condonado: false };
};

// Lista de días con saldo pendiente (> 0) hasta una fecha
const diasPendientes = (data, cid, hasta) => {
  const deudas = data.deudas.filter(d => String(d.choferId)===String(cid) && d.activa);
  if (!deudas.length) return [];
  const fechasIni = deudas.map(d => d.fechaInicio).filter(Boolean).sort();
  const lim = parseF(hasta);
  let cur = fechasIni.length ? parseF(fechasIni[0]) : null;
  if (!cur || !lim) return [];
  const res = [];
  let guard = 0;
  while (cur <= lim && guard++ < 4000) {
    const f = isoF(cur);
    const deudasDia = deudas.filter(d => d.fechaInicio<=f && d.fechaFin>=f);
    const req = deudasDia.reduce((s,d) => s + calcMontoDia(d, f), 0);
    if (req > 0) {
      const ed = estadoDia(data, cid, f);
      if (!ed.condonado && ed.restante > 0) {
        const tipos = [...new Set(deudasDia.filter(d => calcMontoDia(d, f) > 0).map(d => d.tipo))];
        res.push({ fecha:f, monto: ed.restante, montoTotal: req, abonado: ed.pagado, tipos });
      }
    }
    cur.setDate(cur.getDate() + 1);
  }
  return res;
};

// ── Componentes UI base ──────────────────────────────────
const Err  = ({ msg }) => msg ? <div style={{ background:"#fde8e8", color:"#c0392b", border:"1px solid #f5c6cb", borderRadius:8, padding:"8px 12px", fontSize:13, marginBottom:10 }}>{msg}</div> : null;
const Info = ({ children }) => <div style={{ background:"#dbeafe", color:"#1e40af", border:"1px solid #bfdbfe", borderRadius:8, padding:"8px 12px", fontSize:13, marginBottom:10 }}>{children}</div>;
const Ok   = ({ children }) => <div style={{ background:"#dcfce7", color:"#15803d", border:"1px solid #bbf7d0", borderRadius:8, padding:"8px 12px", fontSize:13, marginBottom:10 }}>{children}</div>;
const Warn = ({ children }) => <div style={{ background:"#fef3cd", color:"#856404", border:"1px solid #ffc107", borderRadius:8, padding:"8px 12px", fontSize:13, marginBottom:10 }}>{children}</div>;
const Vacío= ({ texto }) => <div style={{ padding:"1.5rem", textAlign:"center", color:T2, ...card }}>{texto}</div>;

const VigBadge = ({ dias }) => {
  if (dias > 30) return null;
  const bg = dias <= 0 ? "#fde8e8" : "#fef3cd", tc = dias <= 0 ? "#c0392b" : "#856404";
  return <span style={{ background:bg, color:tc, borderRadius:4, padding:"1px 6px", fontSize:11, fontWeight:500, marginLeft:6 }}>{dias <= 0 ? "VENCIDO" : `${dias}d`}</span>;
};

const Tag = ({ children, color = "gray" }) => {
  const m = { gray:["#f0f0f0","#555"], blue:["#dbeafe","#1e40af"], green:["#dcfce7","#15803d"], red:["#fde8e8","#c0392b"], amber:["#fef3cd","#856404"], purple:["#ede9fe","#6d28d9"] };
  const [bg, tc] = m[color] || m.gray;
  return <span style={{ background:bg, color:tc, borderRadius:6, padding:"2px 8px", fontSize:12, fontWeight:500 }}>{children}</span>;
};

const Btn = ({ children, v = "default", ...p }) => {
  const s = { default:{ background:BG, border:`1px solid ${BR}`, color:T }, primary:{ background:"#4A6FA5", border:"none", color:W }, danger:{ background:"#fde8e8", border:"1px solid #f5c6cb", color:"#c0392b" }, success:{ background:"#dcfce7", border:"1px solid #bbf7d0", color:"#15803d" } };
  return <button {...p} style={{ ...s[v], borderRadius:8, padding:"6px 14px", cursor:"pointer", fontSize:13, fontWeight:500, ...(p.style||{}) }}>{children}</button>;
};

const Inp = ({ label, req, onKeyDown, inputMode, ...p }) => {
  // En móviles, abre el teclado numérico para montos (number) y teléfonos (tel)
  const im = inputMode || (p.type === "number" ? "decimal" : p.type === "tel" ? "numeric" : undefined);
  return (
    <div style={{ marginBottom:10 }}>
      {label && <label style={{ fontSize:13, color:T2, display:"block", marginBottom:3 }}>{label}{req && <span style={{ color:"#c0392b" }}> *</span>}</label>}
      <input {...p} inputMode={im} onKeyDown={onKeyDown || focusSiguiente} style={{ width:"100%", boxSizing:"border-box", background:W, color:T, border:`1px solid ${BR}`, borderRadius:8, padding:"6px 10px", ...(p.style||{}) }} />
    </div>
  );
};

const Sel = ({ label, opts, req, onKeyDown, ...p }) => (
  <div style={{ marginBottom:10 }}>
    {label && <label style={{ fontSize:13, color:T2, display:"block", marginBottom:3 }}>{label}{req && <span style={{ color:"#c0392b" }}> *</span>}</label>}
    <select {...p} onKeyDown={onKeyDown || focusSiguiente} style={{ width:"100%", boxSizing:"border-box", padding:"6px 8px", border:`1px solid ${BR}`, borderRadius:8, background:W, color:T }}>
      {opts.map(o => <option key={o.v ?? o} value={o.v ?? o}>{o.l ?? o}</option>)}
    </select>
  </div>
);

const RO = ({ label, value }) => (
  <div style={{ marginBottom:10 }}>
    <label style={{ fontSize:13, color:T2, display:"block", marginBottom:3 }}>{label}</label>
    <div style={{ background:BG, border:`1px solid ${BR}`, borderRadius:8, padding:"6px 10px", fontSize:14, color:T }}>{value || "—"}</div>
  </div>
);

// Campo de contraseña reutilizable (componente estable de nivel superior: no pierde el foco al escribir)
const PassField = ({ value, onChange, show, onToggle }) => (
  <div style={{ marginBottom:10 }}>
    <label style={{ fontSize:13, color:T2, display:"block", marginBottom:3 }}>Contraseña <span style={{ color:"#c0392b" }}>*</span></label>
    <div style={{ position:"relative" }}>
      <input type={show?"text":"password"} value={value||""} onChange={onChange} onKeyDown={focusSiguiente}
        style={{ width:"100%", boxSizing:"border-box", background:W, color:T, border:`1px solid ${BR}`, borderRadius:8, padding:"6px 60px 6px 10px" }} />
      <button type="button" onClick={onToggle} style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", background:"none", border:"none", cursor:"pointer", fontSize:12, color:T2 }}>
        {show ? "Ocultar" : "Ver"}
      </button>
    </div>
  </div>
);

const Modal = ({ title, children }) => (
  <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.4)", zIndex:100, display:"flex", alignItems:"center", justifyContent:"center" }}>
    <div data-form-scope style={{ ...card, width:"min(580px,95vw)", maxHeight:"88vh", overflowY:"auto", boxShadow:"0 4px 24px rgba(0,0,0,0.15)" }}>
      <h3 style={{ margin:"0 0 1rem", fontSize:16, fontWeight:500, color:T }}>{title}</h3>
      {children}
    </div>
  </div>
);

const Acciones = ({ children }) => <div style={{ display:"flex", gap:8, justifyContent:"flex-end", marginTop:4 }}>{children}</div>;

const Steps = ({ steps, cur }) => (
  <div style={{ display:"flex", marginBottom:20 }}>
    {steps.map((s, i) => (
      <div key={i} style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", position:"relative" }}>
        {i > 0 && <div style={{ position:"absolute", top:14, right:"50%", width:"100%", height:2, background: i <= cur ? "#4A6FA5" : BR, zIndex:0 }} />}
        <div style={{ width:28, height:28, borderRadius:"50%", background: i <= cur ? "#4A6FA5" : BG, border:`2px solid ${i <= cur ? "#4A6FA5" : BR}`, color: i <= cur ? W : T2, display:"flex", alignItems:"center", justifyContent:"center", fontSize:13, fontWeight:500, zIndex:1, position:"relative" }}>{i < cur ? "✓" : i+1}</div>
        <div style={{ fontSize:11, color: i === cur ? "#4A6FA5" : T2, marginTop:4, textAlign:"center", fontWeight: i === cur ? 500 : 400 }}>{s}</div>
      </div>
    ))}
  </div>
);

// ── Módulos del administrador ────────────────────────────

function Vehiculos({ data, setData }) {
  const [modal,   setModal]   = useState(false);
  const [rModal,  setRModal]  = useState(false);
  const [form,    setForm]    = useState({});
  const [editId,  setEditId]  = useState(null);
  const [rVeh,    setRVeh]    = useState(null);
  const [rChofer, setRChofer] = useState("");
  const [err,     setErr]     = useState("");

  const campos = [
    ["placa","Placa","text"],["marca","Marca","text"],["modelo","Modelo","text"],["anio","Año","text"],
    ["motor","N° Motor","text"],["chasis","N° Chasis","text"],["color","Color","text"],["factura","Factura","text"],
    ["matriculaFecha","Matrícula fecha","date"],["matriculaVigencia","Matrícula vigencia","date"],
    ["rtvFecha","RTV fecha","date"],
    ["seguroCompania","Compañía seguro","text"],["seguroPoliza","Póliza","text"],
    ["seguroInicio","Seguro inicio","date"],["seguroVigencia","Seguro vigencia","date"],
    ["gpsEmpresa","GPS empresa","text"],["gpsInicio","GPS inicio","date"],["gpsVigencia","GPS vigencia","date"],
  ];

  const abrir = (v = null) => { setForm(v || {}); setEditId(v?.id || null); setErr(""); setModal(true); };

  const guardar = () => {
    const miss = vacio(form, campos.map(c => c[0]));
    if (miss.length) { setErr("Requeridos: " + campos.filter(c => miss.includes(c[0])).map(c => c[1]).join(", ")); return; }
    if (editId) setData(d => ({ ...d, vehiculos: d.vehiculos.map(v => v.id === editId ? { ...form, id: editId } : v) }));
    else        setData(d => ({ ...d, vehiculos: [...d.vehiculos, { ...form, id: Date.now() }] }));
    setModal(false);
  };

  const eliminar = (v) => {
    if (vehiculoTieneDeudas(data, v.id)) { alert("No se puede eliminar: vehículo con deudas vinculadas."); return; }
    setData(d => ({ ...d, vehiculos: d.vehiculos.filter(x => x.id !== v.id) }));
  };

  const reasignar = () => {
    if (!rChofer) { alert("Selecciona un chofer."); return; }
    setRModal(false);
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
        <h3 style={{ margin:0, fontSize:16, fontWeight:500, color:T }}>Vehículos</h3>
        <Btn v="primary" onClick={() => abrir()}>+ Nuevo</Btn>
      </div>
      {data.vehiculos.length === 0 && <Vacío texto="Sin vehículos registrados." />}
      {data.vehiculos.map(v => {
        const da     = vehiculoDeudaActiva(data, v.id);
        const chofer = da ? data.choferes.find(c => c.id === Number(da.choferId)) : null;
        return (
          <div key={v.id} style={{ ...card, marginBottom:10 }}>
            <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
              <div>
                <b style={{ fontSize:15, color:T }}>{v.placa}</b> — {v.marca} {v.modelo} {v.anio}
                <div style={{ fontSize:13, color:T2, marginTop:4 }}>Motor: {v.motor} | Chasis: {v.chasis} | Color: {v.color}</div>
                {chofer  && <div style={{ fontSize:13, marginTop:4 }}>Chofer activo: <b>{chofer.nombres}</b> <Tag color="blue">Deuda activa</Tag></div>}
                {!da     && <div style={{ fontSize:13, color:T2, marginTop:4 }}>Sin deuda activa asignada</div>}
              </div>
              <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                <Btn onClick={() => abrir(v)}>Editar</Btn>
                {da && <Btn onClick={() => { setRVeh(v); setRChofer(""); setRModal(true); }}>Reasignar chofer</Btn>}
                <Btn v="danger" onClick={() => eliminar(v)}>Eliminar</Btn>
              </div>
            </div>
            <div style={{ marginTop:8, display:"flex", flexWrap:"wrap", gap:8, fontSize:13, color:T }}>
              <span>Matrícula: {v.matriculaVigencia}<VigBadge dias={diasVig(v.matriculaVigencia)} /></span>
              <span>· Seguro: {v.seguroVigencia}<VigBadge dias={diasVig(v.seguroVigencia)} /></span>
              <span>· GPS: {v.gpsVigencia}<VigBadge dias={diasVig(v.gpsVigencia)} /></span>
            </div>
          </div>
        );
      })}
      {modal && (
        <Modal title={editId ? "Editar vehículo" : "Nuevo vehículo"}>
          {campos.map(([k,l,t]) => <Inp key={k} label={l} req type={t} value={form[k]||""} onChange={e => setForm(f => ({...f,[k]:e.target.value}))} />)}
          <Err msg={err} />
          <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
        </Modal>
      )}
      {rModal && rVeh && (
        <Modal title={`Reasignar chofer — ${rVeh.placa}`}>
          <Info>Desactiva primero la deuda del chofer actual antes de reasignar.</Info>
          <Sel label="Nuevo chofer" req value={rChofer} onChange={e => setRChofer(e.target.value)} opts={[{v:"",l:"Seleccionar..."}, ...data.choferes.map(c => ({v:c.id, l:c.nombres}))]} />
          <Acciones><Btn onClick={() => setRModal(false)}>Cancelar</Btn><Btn v="primary" onClick={reasignar}>Reasignar</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

function Choferes({ data, setData }) {
  const [modal,  setModal]  = useState(false);
  const [form,   setForm]   = useState({});
  const [editId, setEditId] = useState(null);
  const [err,    setErr]    = useState("");

  const campos = [
    ["nombres","Nombres completos","text"],["cedulaNum","Cédula","text"],
    ["celular","Teléfono celular","tel"],["direccion","Dirección","text"],
    ["cedulaExp","Cédula expedición","date"],["cedulaVig","Cédula vigencia","date"],
    ["licenciaNum","N° Licencia","text"],["licenciaExp","Licencia expedición","date"],["licenciaVig","Licencia vigencia","date"],
    ["antecedentesExp","Antecedentes expedición","date"],["antecedentesVig","Antecedentes vigencia","date"],
  ];

  const abrir = (c) => { setForm({...c}); setEditId(c.id); setErr(""); setModal(true); };

  const guardar = () => {
    const miss = vacio(form, campos.map(c => c[0]));
    if (miss.length) { setErr("Requeridos: " + campos.filter(c => miss.includes(c[0])).map(c => c[1]).join(", ")); return; }
    setData(d => ({ ...d, choferes: d.choferes.map(c => c.id === editId ? {...form, id:editId} : c) }));
    setModal(false);
  };

  const eliminar = (c) => {
    if (choferTieneDeudaActiva(data, c.id)) { alert("No se puede eliminar: el chofer tiene deudas activas."); return; }
    if (choferTienePagos(data, c.id))       { alert("No se puede eliminar: el chofer tiene pagos registrados."); return; }
    setData(d => ({ ...d, choferes: d.choferes.filter(x => x.id !== c.id) }));
  };

  return (
    <div>
      <h3 style={{ margin:"0 0 12px", fontSize:16, fontWeight:500, color:T }}>Choferes</h3>
      {data.choferes.length === 0 && <Vacío texto="Sin choferes registrados. Créalos desde Usuarios." />}
      {data.choferes.map(c => (
        <div key={c.id} style={{ ...card, marginBottom:10 }}>
          <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
            <div>
              <b style={{ fontSize:15, color:T }}>{c.nombres}</b> {choferTieneDeudaActiva(data,c.id) && <Tag color="amber">Deudas activas</Tag>}
              <div style={{ fontSize:13, color:T2 }}>Cédula: {c.cedulaNum}{c.celular && ` · Cel: ${c.celular}`}</div>
              {c.direccion && <div style={{ fontSize:13, color:T2 }}>Dirección: {c.direccion}</div>}
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <Btn onClick={() => abrir(c)}>Editar</Btn>
              <Btn v="danger" onClick={() => eliminar(c)}>Eliminar</Btn>
            </div>
          </div>
          <div style={{ marginTop:8, display:"flex", flexWrap:"wrap", gap:8, fontSize:13, color:T }}>
            <span>Cédula: {c.cedulaVig}<VigBadge dias={diasVig(c.cedulaVig)} /></span>
            <span>· Licencia: {c.licenciaVig}<VigBadge dias={diasVig(c.licenciaVig)} /></span>
            <span>· Antecedentes: {c.antecedentesVig}<VigBadge dias={diasVig(c.antecedentesVig)} /></span>
          </div>
        </div>
      ))}
      {modal && (
        <Modal title="Editar chofer">
          {campos.map(([k,l,t]) => <Inp key={k} label={l} req type={t} value={form[k]||""} onChange={e => setForm(f => ({...f,[k]:e.target.value}))} />)}
          <Err msg={err} />
          <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

function Bancos({ data, setData }) {
  const [modal,  setModal]  = useState(false);
  const [form,   setForm]   = useState({});
  const [editId, setEditId] = useState(null);
  const [err,    setErr]    = useState("");

  const abrir   = (b = null) => { setForm(b || {}); setEditId(b?.id || null); setErr(""); setModal(true); };
  const eliminar = (b) => {
    if (bancoEnUso(data, b.nombre)) { alert("No se puede eliminar: banco en uso en pagos registrados."); return; }
    setData(d => ({ ...d, bancos: d.bancos.filter(x => x.id !== b.id) }));
  };
  const guardar = () => {
    if (!form.nombre?.trim()) { setErr("El nombre es requerido."); return; }
    if (editId) setData(d => ({ ...d, bancos: d.bancos.map(b => b.id === editId ? {...form, id:editId} : b) }));
    else        setData(d => ({ ...d, bancos: [...d.bancos, {...form, id:Date.now()}] }));
    setModal(false);
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
        <h3 style={{ margin:0, fontSize:16, fontWeight:500, color:T }}>Bancos</h3>
        <Btn v="primary" onClick={() => abrir()}>+ Nuevo</Btn>
      </div>
      {data.bancos.length === 0 && <Vacío texto="Sin bancos registrados." />}
      {data.bancos.map(b => (
        <div key={b.id} style={{ ...card2, marginBottom:8, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <span style={{ fontWeight:500, color:T }}>{b.nombre}</span>
          <div style={{ display:"flex", gap:8 }}>
            <Btn onClick={() => abrir(b)}>Editar</Btn>
            <Btn v="danger" onClick={() => eliminar(b)}>Eliminar</Btn>
          </div>
        </div>
      ))}
      {modal && (
        <Modal title={editId ? "Editar banco" : "Nuevo banco"}>
          <Inp label="Nombre del banco" req value={form.nombre||""} onChange={e => setForm(f => ({...f, nombre:e.target.value}))} />
          <Err msg={err} />
          <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

function Cuentas({ data, setData }) {
  const [modal,  setModal]  = useState(false);
  const [form,   setForm]   = useState({});
  const [editId, setEditId] = useState(null);
  const [err,    setErr]    = useState("");

  const abrir    = (c = null) => { setForm(c || {}); setEditId(c?.id || null); setErr(""); setModal(true); };
  const eliminar = (c) => {
    if (cuentaEnUso(data, c.numero)) { alert("No se puede eliminar: cuenta en uso en pagos registrados."); return; }
    setData(d => ({ ...d, cuentas: d.cuentas.filter(x => x.id !== c.id) }));
  };
  const guardar = () => {
    if (vacio(form, ["banco","numero","titular"]).length) { setErr("Todos los campos son requeridos."); return; }
    if (editId) setData(d => ({ ...d, cuentas: d.cuentas.map(c => c.id === editId ? {...form, id:editId} : c) }));
    else        setData(d => ({ ...d, cuentas: [...d.cuentas, {...form, id:Date.now()}] }));
    setModal(false);
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
        <h3 style={{ margin:0, fontSize:16, fontWeight:500, color:T }}>Cuentas destino de depósito</h3>
        <Btn v="primary" onClick={() => abrir()}>+ Nueva</Btn>
      </div>
      {data.cuentas.length === 0 && <Vacío texto="Sin cuentas registradas." />}
      {data.cuentas.map(c => (
        <div key={c.id} style={{ ...card2, marginBottom:8, display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
          <div><b style={{ fontWeight:500, color:T }}>{c.titular}</b><div style={{ fontSize:13, color:T2 }}>{c.banco} · {c.numero}</div></div>
          <div style={{ display:"flex", gap:8 }}>
            <Btn onClick={() => abrir(c)}>Editar</Btn>
            <Btn v="danger" onClick={() => eliminar(c)}>Eliminar</Btn>
          </div>
        </div>
      ))}
      {modal && (
        <Modal title={editId ? "Editar cuenta" : "Nueva cuenta"}>
          <Inp label="Banco" req value={form.banco||""} onChange={e => setForm(f => ({...f, banco:e.target.value}))} />
          <Inp label="Número de cuenta" req value={form.numero||""} onChange={e => setForm(f => ({...f, numero:e.target.value}))} />
          <Inp label="Nombre del titular" req value={form.titular||""} onChange={e => setForm(f => ({...f, titular:e.target.value}))} />
          <Err msg={err} />
          <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

function Inversionistas({ data, setData }) {
  const [modal,  setModal]  = useState(false);
  const [form,   setForm]   = useState({});
  const [editId, setEditId] = useState(null);
  const [err,    setErr]    = useState("");

  const campos = ["nombres","contacto","monto","fechaEntrega","diaPagoIntereses","diaPagoCapital","vehiculoId"];

  const abrir    = (i) => { setForm({...i}); setEditId(i.id); setErr(""); setModal(true); };
  const eliminar = (i) => {
    if (invEnUso(data, i.id)) { alert("No se puede eliminar: inversionista con gastos o pagos registrados."); return; }
    setData(d => ({ ...d, inversionistas: d.inversionistas.filter(x => x.id !== i.id) }));
  };
  const guardar = () => {
    if (vacio(form, campos).length) { setErr("Todos los campos son requeridos."); return; }
    setData(d => ({ ...d, inversionistas: d.inversionistas.map(i => i.id === editId ? {...form, id:editId, monto:Number(form.monto)} : i) }));
    setModal(false);
  };

  return (
    <div>
      <h3 style={{ margin:"0 0 12px", fontSize:16, fontWeight:500, color:T }}>Inversionistas</h3>
      {data.inversionistas.length === 0 && <Vacío texto="Sin inversionistas registrados. Créalos desde Usuarios." />}
      {data.inversionistas.map(i => {
        const veh = data.vehiculos.find(v => v.id === Number(i.vehiculoId));
        return (
          <div key={i.id} style={{ ...card, marginBottom:10 }}>
            <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
              <div>
                <b style={{ fontSize:15, color:T }}>{i.nombres}</b>
                <div style={{ fontSize:13, color:T2 }}>{i.contacto}</div>
              </div>
              <div style={{ display:"flex", gap:8 }}>
                <Btn onClick={() => abrir(i)}>Editar</Btn>
                <Btn v="danger" onClick={() => eliminar(i)}>Eliminar</Btn>
              </div>
            </div>
            <div style={{ marginTop:8, display:"flex", flexWrap:"wrap", gap:12, fontSize:13, color:T }}>
              <span>Capital: <b>${Number(i.monto).toLocaleString()}</b></span>
              <span>Entrega: {i.fechaEntrega}</span>
              <span>Interés mensual: <b>${imMens(i.monto).toFixed(2)}</b></span>
              <span>Capital anual: <b>${capAnual(i.monto).toFixed(2)}</b></span>
              {veh && <Tag color="blue">Vehículo: {veh.placa}</Tag>}
            </div>
          </div>
        );
      })}
      {modal && (
        <Modal title="Editar inversionista">
          <Inp label="Nombres completos" req value={form.nombres||""} onChange={e => setForm(f => ({...f, nombres:e.target.value}))} />
          <Inp label="Contacto" req value={form.contacto||""} onChange={e => setForm(f => ({...f, contacto:e.target.value}))} />
          <Inp label="Monto total ($)" req type="number" value={form.monto||""} onChange={e => setForm(f => ({...f, monto:e.target.value}))} />
          <Inp label="Fecha de entrega" req type="date" value={form.fechaEntrega||""} onChange={e => setForm(f => ({...f, fechaEntrega:e.target.value}))} />
          <Inp label="Día pago intereses" req type="number" min="1" max="31" value={form.diaPagoIntereses||""} onChange={e => setForm(f => ({...f, diaPagoIntereses:e.target.value}))} />
          <Inp label="Día pago capital" req type="number" min="1" max="31" value={form.diaPagoCapital||""} onChange={e => setForm(f => ({...f, diaPagoCapital:e.target.value}))} />
          <Sel label="Vehículo asignado" req value={form.vehiculoId||""} onChange={e => setForm(f => ({...f, vehiculoId:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, ...data.vehiculos.map(v => ({v:v.id, l:`${v.placa} — ${v.marca}`}))]} />
          {form.monto && <Info>Interés mensual: <b>${imMens(Number(form.monto)).toFixed(2)}</b> · Capital anual: <b>${capAnual(Number(form.monto)).toFixed(2)}</b></Info>}
          <Err msg={err} />
          <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

function Deudas({ data, setData }) {
  const [modal,     setModal]     = useState(false);
  const [condModal, setCondModal] = useState(false);
  const [form,      setForm]      = useState({});
  const [condDeuda, setCondDeuda] = useState(null);
  const [condNota,  setCondNota]  = useState("");
  const [editId,    setEditId]    = useState(null);
  const [err,       setErr]       = useState("");
  const [filtroCh,  setFiltroCh]  = useState("");
  const [expCh,     setExpCh]     = useState({});

  const abrir = (d = null) => { setForm(d || {activa:true}); setEditId(d?.id || null); setErr(""); setModal(true); };

  const guardar = () => {
    if (form.formaPago === "manual") {
      if (!form.fechaManual) { setErr("Selecciona la fecha exacta del pago."); return; }
      if (!form.montoManual || Number(form.montoManual) <= 0) { setErr("Ingresa un monto mayor a 0 para la fecha personalizada."); return; }
    }
    if (vacio(form, ["choferId","vehiculoId","tipo","formaPago","fechaInicio","fechaFin","descripcion"]).length) { setErr("Todos los campos son requeridos."); return; }
    if (form.fechaFin < form.fechaInicio) { setErr("La fecha fin no puede ser anterior a la fecha inicio."); return; }
    if (form.tipo === "Cuota") {
      const conflicto = data.deudas.find(d => String(d.vehiculoId) === String(form.vehiculoId) && d.tipo === "Cuota" && d.activa && d.id !== editId);
      if (conflicto) {
        const ch = data.choferes.find(c => c.id === Number(conflicto.choferId));
        setErr(`Este vehículo ya tiene cuota activa asignada a "${ch?.nombres || "—"}". Desactívala primero.`);
        return;
      }
    }
    const rec = { ...form, activa: form.activa !== false, id: editId || Date.now() };
    if (rec.formaPago === "manual") {
      rec.montoManual = Number(rec.montoManual);
      rec.fechaInicio = rec.fechaManual;
      rec.fechaFin = rec.fechaManual;
    }
    if (editId) setData(d => ({ ...d, deudas: d.deudas.map(x => x.id === editId ? rec : x) }));
    else        setData(d => ({ ...d, deudas: [...d.deudas, rec] }));
    setModal(false);
  };

  const toggle = (id, activa) => setData(d => ({ ...d, deudas: d.deudas.map(x => x.id === id ? {...x, activa:!activa} : x) }));

  const condonar = () => {
    if (!condNota.trim()) { alert("Ingresa el motivo de condonación."); return; }
    setData(d => ({
      ...d,
      pagos:  d.pagos.map(p => String(p.choferId) === String(condDeuda.choferId) && p.estado === "no_pagado" ? {...p, estado:"condonado", notaCondonacion:condNota} : p),
      deudas: d.deudas.map(x => x.id === condDeuda.id ? {...x, condonada:true, notaCondonacion:condNota} : x),
    }));
    setCondModal(false);
  };

  const nombreChofer = id => data.choferes.find(c => c.id === Number(id))?.nombres || "—";
  const placaVeh     = id => data.vehiculos.find(v => v.id === Number(id))?.placa   || "—";
  const labelFP      = v  => FORMAS_PAGO.find(f => f.value === v)?.label            || v;

  const renderDeuda = (d) => (
    <div key={d.id} style={{ ...card, marginBottom:10, opacity:d.activa?1:0.65, borderLeft:`4px solid ${d.activa?"#4A6FA5":d.condonada?"#6d28d9":"#aaa"}` }}>
      <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
        <div>
          <b style={{ color:T }}>{nombreChofer(d.choferId)}</b>
          <span style={{ fontSize:13, color:T2 }}> · {placaVeh(d.vehiculoId)}</span>
          <span style={{ marginLeft:8 }}><Tag color={d.tipo==="Cuota"?"blue":d.tipo==="Préstamo"?"amber":"red"}>{d.tipo}</Tag></span>
          {!d.activa && <span style={{ marginLeft:6 }}><Tag color={d.condonada?"purple":"gray"}>{d.condonada?"Condonada":"Inactiva"}</Tag></span>}
          <div style={{ fontSize:13, color:T2, marginTop:4 }}>{d.formaPago==="manual" ? `Personalizado: $${Number(d.montoManual||0).toFixed(2)} el ${d.fechaManual||d.fechaInicio}` : `${labelFP(d.formaPago)} · ${d.fechaInicio} → ${d.fechaFin}`}</div>
          <div style={{ fontSize:13, color:T, marginTop:2 }}>{d.descripcion}</div>
          {d.notaCondonacion && <div style={{ fontSize:12, color:"#6d28d9", marginTop:2 }}>Nota: {d.notaCondonacion}</div>}
        </div>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"flex-start" }}>
          {d.activa && <Btn onClick={() => abrir(d)}>Editar</Btn>}
          <Btn v={d.activa?"danger":"primary"} onClick={() => toggle(d.id, d.activa)}>{d.activa?"Desactivar":"Activar"}</Btn>
          {!d.activa && !d.condonada && <Btn v="success" onClick={() => { setCondDeuda(d); setCondNota(""); setCondModal(true); }}>Condonar</Btn>}
        </div>
      </div>
    </div>
  );

  const grupos = data.choferes
    .filter(c => !filtroCh || String(c.id) === String(filtroCh))
    .map(c => ({ chofer:c, deudas: data.deudas.filter(d => String(d.choferId) === String(c.id)) }))
    .filter(g => g.deudas.length > 0)
    .sort((a,b) => String(a.chofer.nombres||"").localeCompare(String(b.chofer.nombres||"")));

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
        <h3 style={{ margin:0, fontSize:16, fontWeight:500, color:T }}>Deudas</h3>
        <Btn v="primary" onClick={() => abrir()}>+ Nueva</Btn>
      </div>
      <Warn>Las deudas no se eliminan. Desactívalas para que no aparezcan en la agenda de cobros.</Warn>
      <Sel label="Filtrar por chofer" value={filtroCh} onChange={e => setFiltroCh(e.target.value)} opts={[{v:"",l:"Todos los choferes"}, ...data.choferes.map(c => ({v:c.id, l:c.nombres}))]} />
      {data.deudas.length === 0 && <Vacío texto="Sin deudas registradas." />}
      {data.deudas.length > 0 && grupos.length === 0 && <Vacío texto="Este chofer no tiene deudas registradas." />}
      {grupos.map(g => {
        const abierto = !!filtroCh || expCh[g.chofer.id];
        const activas = g.deudas.filter(d => d.activa).length;
        return (
          <div key={g.chofer.id} style={{ marginBottom:10 }}>
            <div onClick={() => setExpCh(e => ({...e, [g.chofer.id]: !e[g.chofer.id]}))} style={{ ...card2, cursor:"pointer", display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
              <div>
                <b style={{ color:T }}>{g.chofer.nombres}</b>
                <span style={{ fontSize:13, color:T2, marginLeft:8 }}>{g.deudas.length} deuda{g.deudas.length===1?"":"s"} · {activas} activa{activas===1?"":"s"}</span>
              </div>
              <span style={{ fontSize:13, color:T2 }}>{abierto ? "Ocultar ▲" : "Ver ▼"}</span>
            </div>
            {abierto && <div style={{ marginTop:8 }}>{g.deudas.map(renderDeuda)}</div>}
          </div>
        );
      })}
      {modal && (
        <Modal title={editId ? "Editar deuda" : "Nueva deuda"}>
          <Sel label="Chofer" req value={form.choferId||""} onChange={e => setForm(f => ({...f, choferId:e.target.value}))} opts={[{v:"",l:"Seleccionar chofer..."}, ...data.choferes.map(c => ({v:c.id, l:c.nombres}))]} />
          <Sel label="Vehículo" req value={form.vehiculoId||""} onChange={e => setForm(f => ({...f, vehiculoId:e.target.value}))} opts={[{v:"",l:"Seleccionar vehículo..."}, ...data.vehiculos.map(v => ({v:v.id, l:`${v.placa} — ${v.marca} ${v.modelo}`}))]} />
          <Sel label="Tipo" req value={form.tipo||""} onChange={e => setForm(f => ({...f, tipo:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, "Cuota", "Préstamo", "Multa"]} />
          <Sel label="Forma de pago" req value={form.formaPago||""} onChange={e => setForm(f => ({...f, formaPago:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, ...FORMAS_PAGO.map(fp => ({v:fp.value, l:fp.label}))]} />
          {form.formaPago === "manual" ? (
            <>
              <Inp label="Fecha exacta del pago" req type="date" value={form.fechaManual||""} onChange={e => setForm(f => ({...f, fechaManual:e.target.value, fechaInicio:e.target.value, fechaFin:e.target.value}))} />
              <Inp label="Monto a pagar ese día ($)" req type="number" min="0.01" step="0.01" value={form.montoManual||""} onChange={e => setForm(f => ({...f, montoManual:e.target.value}))} />
              {form.fechaManual && form.montoManual && <Info>Se cobrará <b>${Number(form.montoManual).toFixed(2)}</b> únicamente el día <b>{form.fechaManual}</b>.</Info>}
            </>
          ) : (
            <>
              <Inp label="Fecha inicio" req type="date" value={form.fechaInicio||""} onChange={e => setForm(f => ({...f, fechaInicio:e.target.value}))} />
              <Inp label="Fecha fin" req type="date" value={form.fechaFin||""} onChange={e => setForm(f => ({...f, fechaFin:e.target.value}))} />
            </>
          )}
          <Inp label="Descripción" req value={form.descripcion||""} onChange={e => setForm(f => ({...f, descripcion:e.target.value}))} />
          <Err msg={err} />
          <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
        </Modal>
      )}
      {condModal && condDeuda && (
        <Modal title="Condonar deuda pendiente">
          <Info>Se condonarán todos los pagos pendientes de <b>{nombreChofer(condDeuda.choferId)}</b> en esta deuda.</Info>
          <Inp label="Motivo de condonación" req value={condNota} onChange={e => setCondNota(e.target.value)} />
          <Acciones><Btn onClick={() => setCondModal(false)}>Cancelar</Btn><Btn v="success" onClick={condonar}>Confirmar</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

function Alertas({ data }) {
  const alertas = [];
  data.vehiculos.forEach(v => {
    [["matriculaVigencia","Matrícula"],["seguroVigencia","Seguro"],["gpsVigencia","GPS"]].forEach(([k,l]) => {
      const d = diasVig(v[k]); if (d <= 30) alertas.push({ tipo:"Vehículo", categoria:l, id:v.placa, label:l, vigencia:v[k], dias:d });
    });
  });
  data.choferes.forEach(c => {
    [["cedulaVig","Cédula"],["licenciaVig","Licencia"],["antecedentesVig","Antecedentes"]].forEach(([k,l]) => {
      const d = diasVig(c[k]); if (d <= 30) alertas.push({ tipo:"Chofer", categoria:l, id:c.nombres, label:l, vigencia:c[k], dias:d });
    });
  });

  // Orden de categorías y conteo
  const ordenCat = ["Matrícula","Seguro","GPS","Cédula","Licencia","Antecedentes"];
  const categorias = [...new Set(alertas.map(a => a.categoria))]
    .sort((a,b) => ordenCat.indexOf(a) - ordenCat.indexOf(b));

  return (
    <div>
      <h3 style={{ margin:"0 0 12px", fontSize:16, fontWeight:500, color:T }}>Panel de alertas de vencimiento</h3>
      {alertas.length === 0
        ? <Vacío texto="Sin alertas próximas. Todo está al día." />
        : categorias.map(cat => {
            const grupo = alertas.filter(a => a.categoria === cat).sort((a,b) => a.dias - b.dias);
            const vencidos = grupo.filter(a => a.dias <= 0).length;
            return (
              <div key={cat} style={{ marginBottom:18 }}>
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
                  <span style={{ fontSize:11, fontWeight:600, color:T2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{cat}</span>
                  <span style={{ fontSize:12, color:T2 }}>· {grupo.length} alerta{grupo.length===1?"":"s"}{vencidos>0 && <span style={{ color:"#c0392b" }}> · {vencidos} vencida{vencidos===1?"":"s"}</span>}</span>
                </div>
                {grupo.map((a,i) => (
                  <div key={i} style={{ display:"flex", alignItems:"center", gap:12, background:a.dias<=0?"#fde8e8":"#fef3cd", border:`1px solid ${a.dias<=0?"#f5c6cb":"#ffc107"}`, borderRadius:12, padding:"0.75rem 1rem", marginBottom:8 }}>
                    <div style={{ flex:1 }}>
                      <b style={{ color:a.dias<=0?"#c0392b":"#856404" }}>{a.tipo}: {a.id}</b>
                      <div style={{ fontSize:13, color:T }}>{a.label} vence: {a.vigencia}</div>
                    </div>
                    <Tag color={a.dias<=0?"red":"amber"}>{a.dias<=0?"VENCIDO":`En ${a.dias} días`}</Tag>
                  </div>
                ))}
              </div>
            );
          })
      }
    </div>
  );
}

function Inversiones({ data, setData }) {
  const [mG, setMG] = useState(false);
  const [mP, setMP] = useState(false);
  const [fG, setFG] = useState({});
  const [fP, setFP] = useState({});
  const [errG, setErrG] = useState("");
  const [errP, setErrP] = useState("");

  const guardarGasto = () => {
    if (vacio(fG, ["fecha","vehiculoId","monto","descripcion"]).length) { setErrG("Todos los campos son requeridos."); return; }
    setData(d => ({ ...d, gastosInv: [...d.gastosInv, {...fG, id:Date.now(), monto:Number(fG.monto)}] }));
    setMG(false);
  };

  const guardarPago = () => {
    if (vacio(fP, ["fecha","monto","tipo","metodoPago","referencia"]).length) { setErrP("Todos los campos son requeridos."); return; }
    setData(d => ({ ...d, pagosInv: [...d.pagosInv, {...fP, id:Date.now(), monto:Number(fP.monto), realizado:true}] }));
    setMP(false);
  };

  return (
    <div>
      <h3 style={{ margin:"0 0 12px", fontSize:16, fontWeight:500, color:T }}>Inversiones</h3>
      {data.inversionistas.length === 0 && <Vacío texto="Sin inversionistas registrados." />}
      {data.inversionistas.map(inv => {
        const gastos = data.gastosInv.filter(g => g.invId === inv.id);
        const pagos  = data.pagosInv.filter(p => p.invId === inv.id && p.realizado);
        const tG     = gastos.reduce((s,g) => s + g.monto, 0);
        const tP     = pagos.reduce((s,p) => s + p.monto, 0);
        return (
          <div key={inv.id} style={{ ...card, marginBottom:16 }}>
            <b style={{ fontSize:15, color:T }}>{inv.nombres}</b>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))", gap:8, margin:"12px 0" }}>
              {[["Capital recibido","$"+Number(inv.monto).toLocaleString()],["Total gastado","$"+tG.toLocaleString()],["Pagado al inv.","$"+tP.toLocaleString()],["Balance","$"+(inv.monto-tG-tP).toLocaleString()]].map(([l,v]) => (
                <div key={l} style={kpi}><div style={{ fontSize:12, color:T2 }}>{l}</div><div style={{ fontSize:18, fontWeight:500, color:T }}>{v}</div></div>
              ))}
            </div>
            <div style={{ fontSize:13, marginBottom:8, color:T }}>
              <b>Egresos:</b>
              {gastos.length === 0 && <span style={{ color:T2 }}> Sin egresos.</span>}
              {gastos.map(g => { const veh = data.vehiculos.find(v => v.id === Number(g.vehiculoId)); return <div key={g.id} style={{ color:T2 }}>{g.fecha} · {veh?.placa||"—"} · {g.descripcion} · <b style={{ color:T }}>${g.monto}</b></div>; })}
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <Btn onClick={() => { setFG({invId:inv.id}); setErrG(""); setMG(true); }}>+ Registrar egreso</Btn>
              <Btn v="primary" onClick={() => { setFP({invId:inv.id}); setErrP(""); setMP(true); }}>+ Registrar pago</Btn>
            </div>
          </div>
        );
      })}
      {mG && (
        <Modal title="Registrar egreso de inversión">
          <Inp label="Fecha" req type="date" value={fG.fecha||""} onChange={e => setFG(f => ({...f, fecha:e.target.value}))} />
          <Sel label="Vehículo afectado" req value={fG.vehiculoId||""} onChange={e => setFG(f => ({...f, vehiculoId:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, ...data.vehiculos.map(v => ({v:v.id, l:`${v.placa} — ${v.marca}`}))]} />
          <Inp label="Monto ($)" req type="number" value={fG.monto||""} onChange={e => setFG(f => ({...f, monto:e.target.value}))} />
          <Inp label="Descripción" req value={fG.descripcion||""} onChange={e => setFG(f => ({...f, descripcion:e.target.value}))} />
          <Err msg={errG} />
          <Acciones><Btn onClick={() => setMG(false)}>Cancelar</Btn><Btn v="primary" onClick={guardarGasto}>Guardar</Btn></Acciones>
        </Modal>
      )}
      {mP && (
        <Modal title="Registrar pago a inversionista">
          <Inp label="Fecha" req type="date" value={fP.fecha||""} onChange={e => setFP(f => ({...f, fecha:e.target.value}))} />
          <Inp label="Monto ($)" req type="number" value={fP.monto||""} onChange={e => setFP(f => ({...f, monto:e.target.value}))} />
          <Sel label="Tipo" req value={fP.tipo||""} onChange={e => setFP(f => ({...f, tipo:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, "Interés mensual", "Capital anual"]} />
          <Sel label="Método de pago" req value={fP.metodoPago||""} onChange={e => setFP(f => ({...f, metodoPago:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, "Transferencia", "Efectivo", "Pago digital"]} />
          <Inp label="Referencia / comprobante" req value={fP.referencia||""} onChange={e => setFP(f => ({...f, referencia:e.target.value}))} />
          <Err msg={errP} />
          <Acciones><Btn onClick={() => setMP(false)}>Cancelar</Btn><Btn v="primary" onClick={guardarPago}>Guardar</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

function Usuarios({ data, setData }) {
  const [modal,  setModal]  = useState(false);
  const [rol,    setRol]    = useState("");
  const [paso,   setPaso]   = useState(0);
  const [form,   setForm]   = useState({});
  const [fCh,    setFCh]    = useState({});
  const [fInv,   setFInv]   = useState({});
  const [err,    setErr]    = useState("");
  const [showP,  setShowP]  = useState(false);
  const [editId, setEditId] = useState(null);

  const camposCh = [
    ["nombres","Nombres completos","text"],["cedulaNum","Cédula","text"],
    ["celular","Teléfono celular","tel"],["direccion","Dirección","text"],
    ["cedulaExp","Cédula expedición","date"],["cedulaVig","Cédula vigencia","date"],
    ["licenciaNum","N° Licencia","text"],["licenciaExp","Licencia expedición","date"],["licenciaVig","Licencia vigencia","date"],
    ["antecedentesExp","Antecedentes expedición","date"],["antecedentesVig","Antecedentes vigencia","date"],
  ];
  const camposInv = ["nombres","contacto","monto","fechaEntrega","diaPagoIntereses","diaPagoCapital","vehiculoId"];

  const abrir = (u = null) => {
    if (u) {
      setEditId(u.id); setRol(u.rol); setForm(u);
      setPaso(["chofer","inversionista"].includes(u.rol) ? 1 : 0);
      setFCh(data.choferes.find(c => c.id === u.choferId) || {});
      setFInv(data.inversionistas.find(i => i.id === u.inversionistaId) || {});
    } else {
      setEditId(null); setRol(""); setForm({}); setFCh({}); setFInv({}); setPaso(0);
    }
    setErr(""); setShowP(false); setModal(true);
  };

  const selRol = (r) => { setRol(r); setForm({}); setFCh({}); setFInv({}); setPaso(0); setErr(""); };

  const siguiente = () => {
    setErr("");
    if (rol === "chofer"       && vacio(fCh,  camposCh.map(c=>c[0])).length) { setErr("Completa todos los campos del chofer."); return; }
    if (rol === "inversionista" && vacio(fInv, camposInv).length)             { setErr("Completa todos los campos del inversionista."); return; }
    setPaso(p => p + 1);
  };

  const guardar = () => {
    setErr("");
    if (vacio(form, ["email","password"]).length)                                          { setErr("Email y contraseña son requeridos."); return; }
    if (["administrador","digitador"].includes(rol) && !form.nombre?.trim())               { setErr("El nombre completo es requerido."); return; }
    if (rol === "chofer"       && !editId && !form.choferId       && !fCh.nombres)         { setErr("Completa los datos del chofer."); return; }
    if (rol === "inversionista" && !editId && !form.inversionistaId && !fInv.nombres)      { setErr("Completa los datos del inversionista."); return; }

    let choferId = form.choferId, invId = form.inversionistaId;

    if (rol === "chofer") {
      if (!editId || !form.choferId) {
        const nc = { ...fCh, id: Date.now() };
        setData(d => ({ ...d, choferes: [...d.choferes, nc] }));
        choferId = nc.id;
      } else {
        setData(d => ({ ...d, choferes: d.choferes.map(c => c.id === form.choferId ? {...fCh, id:form.choferId} : c) }));
      }
    }

    if (rol === "inversionista") {
      if (!editId || !form.inversionistaId) {
        const ni = { ...fInv, id:Date.now()+1, monto:Number(fInv.monto), diaPagoIntereses:Number(fInv.diaPagoIntereses), diaPagoCapital:Number(fInv.diaPagoCapital) };
        setData(d => ({ ...d, inversionistas: [...d.inversionistas, ni] }));
        invId = ni.id;
      } else {
        setData(d => ({ ...d, inversionistas: d.inversionistas.map(i => i.id === form.inversionistaId ? {...fInv, id:form.inversionistaId, monto:Number(fInv.monto)} : i) }));
      }
    }

    const nombre = rol === "chofer" ? fCh.nombres : rol === "inversionista" ? fInv.nombres : form.nombre;
    const u = { ...form, rol, activo:true, id:editId||Date.now()+2, nombre, choferId, inversionistaId:invId };
    if (editId) setData(d => ({ ...d, usuarios: d.usuarios.map(x => x.id === editId ? u : x) }));
    else        setData(d => ({ ...d, usuarios: [...d.usuarios, u] }));
    setModal(false);
  };

  const eliminar = (u) => {
    if (u.rol === "administrador") {
      if (data.usuarios.filter(x => x.rol === "administrador" && x.activo).length <= 1) { alert("Debe existir al menos un administrador."); return; }
    }
    setData(d => ({ ...d, usuarios: d.usuarios.filter(x => x.id !== u.id) }));
  };

  const rc = { administrador:"blue", digitador:"amber", chofer:"green", inversionista:"gray" };
  const onPass   = e => setForm(f => ({...f, password:e.target.value}));
  const togglePass = () => setShowP(s => !s);

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
        <h3 style={{ margin:0, fontSize:16, fontWeight:500, color:T }}>Usuarios</h3>
        <Btn v="primary" onClick={() => abrir()}>+ Nuevo usuario</Btn>
      </div>
      {data.usuarios.length === 0 && <Vacío texto="Sin usuarios registrados." />}
      {data.usuarios.map(u => (
        <div key={u.id} style={{ ...card2, marginBottom:8, display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
          <div><b style={{ fontWeight:500, color:T }}>{u.nombre}</b> <Tag color={rc[u.rol]||"gray"}>{u.rol}</Tag><div style={{ fontSize:13, color:T2 }}>{u.email}</div></div>
          <div style={{ display:"flex", gap:8 }}>
            <Btn onClick={() => abrir(u)}>Editar</Btn>
            <Btn v="danger" onClick={() => eliminar(u)}>Eliminar</Btn>
          </div>
        </div>
      ))}
      {modal && (
        <Modal title={editId ? "Editar usuario" : "Nuevo usuario"}>
          {!rol && (
            <>
              <p style={{ fontSize:14, color:T2, margin:"0 0 16px" }}>Selecciona el tipo de usuario:</p>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
                {[["administrador","🔧","Administrador","Acceso total"],["digitador","⌨️","Digitador","Registra pagos"],["chofer","🚗","Chofer","Conductor del vehículo"],["inversionista","💼","Inversionista","Accionista del negocio"]].map(([r,ico,titulo,desc]) => (
                  <button key={r} onClick={() => selRol(r)} style={{ background:W, border:`1px solid ${BR}`, borderRadius:10, padding:"1rem", cursor:"pointer", textAlign:"left" }}>
                    <div style={{ fontSize:22, marginBottom:6 }}>{ico}</div>
                    <div style={{ fontWeight:500, color:T, fontSize:14 }}>{titulo}</div>
                    <div style={{ fontSize:12, color:T2, marginTop:2 }}>{desc}</div>
                  </button>
                ))}
              </div>
              <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn></Acciones>
            </>
          )}
          {rol && ["administrador","digitador"].includes(rol) && (
            <>
              <Ok>Rol: <b>{rol}</b></Ok>
              <Inp label="Nombre completo" req value={form.nombre||""} onChange={e => setForm(f => ({...f, nombre:e.target.value}))} />
              <Inp label="Email" req type="email" value={form.email||""} onChange={e => setForm(f => ({...f, email:e.target.value}))} />
              <PassField value={form.password} onChange={onPass} show={showP} onToggle={togglePass} />
              <Err msg={err} />
              <Acciones><Btn onClick={() => setRol("")}>← Volver</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
            </>
          )}
          {rol === "chofer" && (
            <>
              <Steps steps={["Datos del chofer","Credenciales"]} cur={paso} />
              {paso === 0 && (
                <>
                  <Ok>Paso 1: Datos personales del chofer</Ok>
                  {camposCh.map(([k,l,t]) => <Inp key={k} label={l} req type={t} value={fCh[k]||""} onChange={e => setFCh(f => ({...f,[k]:e.target.value}))} />)}
                  <Err msg={err} />
                  <Acciones><Btn onClick={() => setRol("")}>← Volver</Btn><Btn v="primary" onClick={siguiente}>Siguiente →</Btn></Acciones>
                </>
              )}
              {paso === 1 && (
                <>
                  <Ok>Paso 2: Credenciales de acceso</Ok>
                  <RO label="Nombre (del maestro de choferes)" value={fCh.nombres} />
                  <Inp label="Email" req type="email" value={form.email||""} onChange={e => setForm(f => ({...f, email:e.target.value}))} />
                  <PassField value={form.password} onChange={onPass} show={showP} onToggle={togglePass} />
                  <Err msg={err} />
                  <Acciones><Btn onClick={() => setPaso(0)}>← Anterior</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
                </>
              )}
            </>
          )}
          {rol === "inversionista" && (
            <>
              <Steps steps={["Datos del inversionista","Credenciales"]} cur={paso} />
              {paso === 0 && (
                <>
                  <Ok>Paso 1: Datos del inversionista</Ok>
                  <Inp label="Nombres completos" req value={fInv.nombres||""} onChange={e => setFInv(f => ({...f, nombres:e.target.value}))} />
                  <Inp label="Contacto" req value={fInv.contacto||""} onChange={e => setFInv(f => ({...f, contacto:e.target.value}))} />
                  <Inp label="Monto total ($)" req type="number" value={fInv.monto||""} onChange={e => setFInv(f => ({...f, monto:e.target.value}))} />
                  <Inp label="Fecha de entrega" req type="date" value={fInv.fechaEntrega||""} onChange={e => setFInv(f => ({...f, fechaEntrega:e.target.value}))} />
                  <Inp label="Día pago intereses" req type="number" min="1" max="31" value={fInv.diaPagoIntereses||""} onChange={e => setFInv(f => ({...f, diaPagoIntereses:e.target.value}))} />
                  <Inp label="Día pago capital" req type="number" min="1" max="31" value={fInv.diaPagoCapital||""} onChange={e => setFInv(f => ({...f, diaPagoCapital:e.target.value}))} />
                  <Sel label="Vehículo asignado" req value={fInv.vehiculoId||""} onChange={e => setFInv(f => ({...f, vehiculoId:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, ...data.vehiculos.map(v => ({v:v.id, l:`${v.placa} — ${v.marca}`}))]} />
                  {fInv.monto && <Info>Interés mensual: <b>${imMens(Number(fInv.monto)).toFixed(2)}</b> · Capital anual: <b>${capAnual(Number(fInv.monto)).toFixed(2)}</b></Info>}
                  <Err msg={err} />
                  <Acciones><Btn onClick={() => setRol("")}>← Volver</Btn><Btn v="primary" onClick={siguiente}>Siguiente →</Btn></Acciones>
                </>
              )}
              {paso === 1 && (
                <>
                  <Ok>Paso 2: Credenciales de acceso</Ok>
                  <RO label="Nombre (del maestro de inversionistas)" value={fInv.nombres} />
                  <Inp label="Email" req type="email" value={form.email||""} onChange={e => setForm(f => ({...f, email:e.target.value}))} />
                  <PassField value={form.password} onChange={onPass} show={showP} onToggle={togglePass} />
                  <Err msg={err} />
                  <Acciones><Btn onClick={() => setPaso(0)}>← Anterior</Btn><Btn v="primary" onClick={guardar}>Guardar</Btn></Acciones>
                </>
              )}
            </>
          )}
        </Modal>
      )}
    </div>
  );
}

// ── Reportes Excel + Copias de seguridad ─────────────────
function Reportes({ data, setData }) {
  const [msg, setMsg] = useState("");
  const [confRest, setConfRest] = useState(null); // respaldo pendiente de confirmar
  const fileRef = useRef(null);

  const nombreChofer = id => data.choferes.find(c => c.id === Number(id))?.nombres || "—";
  const placaVeh     = id => data.vehiculos.find(v => v.id === Number(id))?.placa   || "—";
  const nombreInv    = id => data.inversionistas.find(i => i.id === id)?.nombres    || "—";
  const labelFP      = v  => FORMAS_PAGO.find(f => f.value === v)?.label            || v;

  const mVehiculos = () => data.vehiculos.map(v => ({
    Placa:v.placa, Marca:v.marca, Modelo:v.modelo, "Año":v.anio, "N° Motor":v.motor, "N° Chasis":v.chasis,
    Color:v.color, Factura:v.factura, "Matrícula fecha":v.matriculaFecha, "Matrícula vigencia":v.matriculaVigencia,
    "RTV fecha":v.rtvFecha, "Seguro compañía":v.seguroCompania, "Póliza":v.seguroPoliza,
    "Seguro inicio":v.seguroInicio, "Seguro vigencia":v.seguroVigencia,
    "GPS empresa":v.gpsEmpresa, "GPS inicio":v.gpsInicio, "GPS vigencia":v.gpsVigencia,
  }));

  const mChoferes = () => data.choferes.map(c => ({
    Nombres:c.nombres, "Cédula":c.cedulaNum, "Cédula expedición":c.cedulaExp, "Cédula vigencia":c.cedulaVig,
    "N° Licencia":c.licenciaNum, "Licencia expedición":c.licenciaExp, "Licencia vigencia":c.licenciaVig,
    "Antecedentes expedición":c.antecedentesExp, "Antecedentes vigencia":c.antecedentesVig,
  }));

  const mInversionistas = () => data.inversionistas.map(i => ({
    Nombres:i.nombres, Contacto:i.contacto, "Capital ($)":Number(i.monto), "Fecha entrega":i.fechaEntrega,
    "Día pago intereses":i.diaPagoIntereses, "Día pago capital":i.diaPagoCapital,
    "Interés mensual ($)":round2(imMens(i.monto)), "Capital anual ($)":round2(capAnual(i.monto)),
    "Vehículo":placaVeh(i.vehiculoId),
  }));

  const mDeudas = () => data.deudas.map(d => ({
    Chofer:nombreChofer(d.choferId), "Vehículo":placaVeh(d.vehiculoId), Tipo:d.tipo,
    "Forma de pago": d.formaPago==="manual"
      ? `Personalizado $${Number(d.montoManual||0).toFixed(2)} el ${d.fechaManual||d.fechaInicio}`
      : labelFP(d.formaPago),
    "Fecha inicio":d.fechaInicio, "Fecha fin":d.fechaFin, "Descripción":d.descripcion,
    Estado:d.activa?"Activa":(d.condonada?"Condonada":"Inactiva"), "Nota condonación":d.notaCondonacion||"",
  }));

  const mPagos = () => data.pagos.map(p => ({
    Chofer:nombreChofer(p.choferId), "Fecha cuota":p.fecha, "Fecha comprobante":p.fechaComp, Hora:p.hora,
    "Monto ($)":Number(p.monto), Estado:p.estado, Tipo:p.tipo, "Comprobante":p.comprobante,
    Banco:p.banco, "Cuenta destino":p.cuentaDestino, "Imputación":p.esImputacion?"Sí":"No",
    "Nota condonación":p.notaCondonacion||"",
  }));

  const mPagosInv = () => data.pagosInv.map(p => ({
    Inversionista:nombreInv(p.invId), Fecha:p.fecha, "Monto ($)":Number(p.monto),
    Tipo:p.tipo, "Método":p.metodoPago, Referencia:p.referencia,
  }));

  const mGastosInv = () => data.gastosInv.map(g => ({
    Inversionista:nombreInv(g.invId), Fecha:g.fecha, "Vehículo":placaVeh(g.vehiculoId),
    "Monto ($)":Number(g.monto), "Descripción":g.descripcion,
  }));

  const exportar = (nombreHoja, filas, archivo) => {
    try { descargarLibro([{ nombre:nombreHoja, filas }], `${archivo}_${hoy()}.xlsx`); setMsg(`Reporte generado: ${archivo}_${hoy()}.xlsx`); }
    catch (e) { setMsg("Error al generar: " + e.message); }
  };

  const exportarTodo = () => {
    try {
      descargarLibro([
        { nombre:"Vehículos",          filas:mVehiculos() },
        { nombre:"Choferes",           filas:mChoferes() },
        { nombre:"Inversionistas",     filas:mInversionistas() },
        { nombre:"Deudas",             filas:mDeudas() },
        { nombre:"Cobros a choferes",  filas:mPagos() },
        { nombre:"Pagos a inv.",       filas:mPagosInv() },
        { nombre:"Gastos de inv.",     filas:mGastosInv() },
      ], `reporte_completo_${hoy()}.xlsx`);
      setMsg(`Reporte completo generado: reporte_completo_${hoy()}.xlsx`);
    } catch (e) { setMsg("Error al generar: " + e.message); }
  };

  const descargarRespaldo = () => {
    try {
      const blob = new Blob([JSON.stringify(data, null, 2)], { type:"application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `respaldo_flota_${hoy()}.json`;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setMsg("Copia de seguridad descargada.");
    } catch (e) { setMsg("Error en respaldo: " + e.message); }
  };

  const restaurar = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const obj = JSON.parse(ev.target.result);
        if (!obj || typeof obj !== "object" || !Array.isArray(obj.usuarios)) {
          setMsg("Archivo inválido: no es un respaldo de flota.");
          return;
        }
        // Abrimos un modal propio (window.confirm puede estar bloqueado en algunos entornos)
        setConfRest({ obj, nombre: file.name });
      } catch (err) { setMsg("Archivo inválido: " + err.message); }
    };
    reader.onerror = () => setMsg("No se pudo leer el archivo.");
    reader.readAsText(file);
    e.target.value = "";
  };

  const confirmarRestaurar = () => {
    if (!confRest) return;
    const obj = { ...confRest.obj };
    // Garantizar que existan todas las colecciones esperadas
    Object.keys(INIT).forEach(k => { if (!obj[k]) obj[k] = INIT[k]; });
    if (!obj.usuarios || obj.usuarios.length === 0) obj.usuarios = INIT.usuarios;
    setData(obj);
    setConfRest(null);
    setMsg("Copia restaurada correctamente.");
  };

  const items = [
    ["Vehículos",          () => exportar("Vehículos",         mVehiculos(),      "vehiculos"),       data.vehiculos.length],
    ["Choferes",           () => exportar("Choferes",          mChoferes(),       "choferes"),        data.choferes.length],
    ["Inversionistas",     () => exportar("Inversionistas",    mInversionistas(), "inversionistas"),  data.inversionistas.length],
    ["Deudas",             () => exportar("Deudas",            mDeudas(),         "deudas"),          data.deudas.length],
    ["Cobros a choferes",  () => exportar("Cobros",            mPagos(),          "cobros"),          data.pagos.length],
    ["Pagos a inversionistas", () => exportar("Pagos inv",     mPagosInv(),       "pagos_inversionistas"), data.pagosInv.length],
    ["Gastos de inversión", () => exportar("Gastos inv",       mGastosInv(),      "gastos_inversion"), data.gastosInv.length],
  ];

  return (
    <div>
      <h3 style={{ margin:"0 0 12px", fontSize:16, fontWeight:500, color:T }}>Reportes en Excel</h3>
      {msg && <Ok>{msg}</Ok>}
      <Info>Descarga cada listado por separado o todo en un solo archivo con varias hojas.</Info>

      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))", gap:10, marginBottom:16 }}>
        {items.map(([titulo, fn, n]) => (
          <div key={titulo} style={{ ...card2, display:"flex", justifyContent:"space-between", alignItems:"center", gap:8 }}>
            <div>
              <div style={{ fontWeight:500, color:T, fontSize:14 }}>{titulo}</div>
              <div style={{ fontSize:12, color:T2 }}>{n} registro{n===1?"":"s"}</div>
            </div>
            <Btn v="primary" onClick={fn}>Excel</Btn>
          </div>
        ))}
      </div>

      <div style={{ ...card, background:BG, marginBottom:16 }}>
        <div style={{ fontWeight:500, color:T, marginBottom:8 }}>Reporte completo</div>
        <div style={{ fontSize:13, color:T2, marginBottom:10 }}>Genera un único archivo Excel con todas las hojas del sistema.</div>
        <Btn v="primary" onClick={exportarTodo}>Descargar reporte completo (todas las hojas)</Btn>
      </div>

      <h3 style={{ margin:"0 0 8px", fontSize:16, fontWeight:500, color:T }}>Copia de seguridad de datos</h3>
      <div style={{ ...card, background:BG }}>
        <div style={{ fontSize:13, color:T2, marginBottom:10 }}>Descarga un respaldo completo (formato JSON) para guardarlo o trasladarlo a otro equipo. Puedes restaurarlo cuando quieras.</div>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
          <Btn onClick={descargarRespaldo}>Descargar copia de seguridad (.json)</Btn>
          <Btn onClick={() => fileRef.current?.click()}>Restaurar desde copia</Btn>
          <input ref={fileRef} type="file" accept="application/json,.json" onChange={restaurar} style={{ display:"none" }} />
        </div>
      </div>

      {confRest && (
        <Modal title="Restaurar copia de seguridad">
          <Warn>Esto reemplazará <b>TODOS</b> los datos actuales por los del archivo <b>{confRest.nombre}</b>. Esta acción no se puede deshacer.</Warn>
          <div style={{ fontSize:13, color:T2, marginBottom:10 }}>
            El respaldo contiene: {confRest.obj.vehiculos?.length||0} vehículos · {confRest.obj.choferes?.length||0} choferes · {confRest.obj.inversionistas?.length||0} inversionistas · {confRest.obj.pagos?.length||0} pagos.
          </div>
          <Acciones>
            <Btn onClick={() => setConfRest(null)}>Cancelar</Btn>
            <Btn v="primary" onClick={confirmarRestaurar}>Restaurar definitivamente</Btn>
          </Acciones>
        </Modal>
      )}
    </div>
  );
}

// ── Editar pagos registrados (solo administrador) ────────
function EditarPagos({ data, setData }) {
  const [filtro, setFiltro] = useState("");
  const [filtroTipo, setFiltroTipo] = useState("");
  const [modal,  setModal]  = useState(false);
  const [delPago, setDelPago] = useState(null);
  const [form,   setForm]   = useState({});
  const [err,    setErr]    = useState("");

  const nombreChofer = id => data.choferes.find(c => c.id === Number(id))?.nombres || "—";
  const elMap = { pagado:"Pagado", abono:"Abono", no_pagado:"No pagado", condonado:"Condonado" };
  const ecMap = { pagado:"green", abono:"amber", no_pagado:"red", condonado:"purple" };

  const pagos = data.pagos
    .filter(p => !filtro || String(p.choferId) === String(filtro))
    .filter(p => !filtroTipo || (filtroTipo === "imput" ? p.esImputacion : !p.esImputacion))
    .slice()
    .sort((a,b) => String(b.fecha||"").localeCompare(String(a.fecha||"")));

  const abrir = (p) => { setForm({ ...p }); setErr(""); setModal(true); };

  const guardar = () => {
    setErr("");
    if (!form.fecha) { setErr("La fecha de la cuota es requerida."); return; }
    if (!form.monto || Number(form.monto) <= 0) { setErr("El monto debe ser mayor a 0."); return; }
    setData(d => ({ ...d, pagos: d.pagos.map(x => x.id === form.id ? { ...form, monto:Number(form.monto) } : x) }));
    setModal(false);
  };

  const eliminar = (p) => setDelPago(p);
  const confirmarEliminar = () => {
    setData(d => ({ ...d, pagos: d.pagos.filter(x => x.id !== delPago.id) }));
    setDelPago(null);
  };

  const bancoOpts = [{v:"",l:"Seleccionar..."}, ...data.bancos.map(b => ({v:b.nombre, l:b.nombre}))];
  if (form.banco && !data.bancos.some(b => b.nombre === form.banco)) bancoOpts.push({ v:form.banco, l:`${form.banco} (actual)` });
  const cuentaOpts = [{v:"",l:"Seleccionar..."}, ...data.cuentas.map(c => ({v:c.numero, l:`${c.titular} · ${c.numero}`}))];
  if (form.cuentaDestino && !data.cuentas.some(c => c.numero === form.cuentaDestino)) cuentaOpts.push({ v:form.cuentaDestino, l:`${form.cuentaDestino} (actual)` });

  return (
    <div>
      <h3 style={{ margin:"0 0 12px", fontSize:16, fontWeight:500, color:T }}>Editar pagos registrados</h3>
      <Info>Sección exclusiva del administrador para corregir o eliminar los pagos registrados por el digitador.</Info>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0 12px" }}>
        <Sel label="Filtrar por chofer" value={filtro} onChange={e => setFiltro(e.target.value)} opts={[{v:"",l:"Todos los choferes"}, ...data.choferes.map(c => ({v:c.id, l:c.nombres}))]} />
        <Sel label="Filtrar por tipo de registro" value={filtroTipo} onChange={e => setFiltroTipo(e.target.value)} opts={[{v:"",l:"Todos"},{v:"normal",l:"Pago normal"},{v:"imput",l:"Imputación a futuro"}]} />
      </div>
      {pagos.length === 0
        ? <Vacío texto="Sin pagos registrados." />
        : <div style={{ overflowX:"auto", maxHeight:420, overflowY:"auto", background:W, border:`1px solid ${BR}`, borderRadius:8 }}>
            <table style={TBL}>
              <thead><tr>{["Fecha cuota","Chofer","Monto","Estado","Registro","Comprobante","Banco","Acciones"].map(h => <th key={h} style={{ ...TH, position:"sticky", top:0 }}>{h}</th>)}</tr></thead>
              <tbody>{pagos.map(p => (
                <tr key={p.id}>
                  <td style={TD}>{p.fecha}</td>
                  <td style={TD}>{nombreChofer(p.choferId)}</td>
                  <td style={TD}>${p.monto}</td>
                  <td style={TD}><Tag color={ecMap[p.estado]||"gray"}>{elMap[p.estado]||p.estado}</Tag></td>
                  <td style={TD}><Tag color={p.esImputacion?"amber":"blue"}>{p.esImputacion?"Imputación":"Pago normal"}</Tag></td>
                  <td style={TD}>{p.comprobante}</td>
                  <td style={TD}>{p.banco}</td>
                  <td style={TD}>
                    <div style={{ display:"flex", gap:6 }}>
                      <Btn onClick={() => abrir(p)}>Editar</Btn>
                      <Btn v="danger" onClick={() => eliminar(p)}>Eliminar</Btn>
                    </div>
                  </td>
                </tr>
              ))}</tbody>
            </table>
          </div>
      }
      {modal && (
        <Modal title="Editar pago registrado">
          <RO label="Chofer" value={nombreChofer(form.choferId)} />
          <Inp label="Fecha de la cuota" req type="date" value={form.fecha||""} onChange={e => setForm(f => ({...f, fecha:e.target.value}))} />
          <Inp label="Monto ($)" req type="number" min="0.01" step="0.01" value={form.monto||""} onChange={e => setForm(f => ({...f, monto:e.target.value}))} />
          <Sel label="Estado" req value={form.estado||""} onChange={e => setForm(f => ({...f, estado:e.target.value}))} opts={[{v:"pagado",l:"Pagado"},{v:"abono",l:"Abono"},{v:"no_pagado",l:"No pagado"},{v:"condonado",l:"Condonado"}]} />
          <Inp label="Tipo" value={form.tipo||""} onChange={e => setForm(f => ({...f, tipo:e.target.value}))} />
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0 12px" }}>
            <Inp label="N° Comprobante" value={form.comprobante||""} onChange={e => setForm(f => ({...f, comprobante:e.target.value}))} />
            <Inp label="Hora (HH:MM)" inputMode="numeric" placeholder="14:30" maxLength={5} value={form.hora||""} onChange={e => setForm(f => ({...f, hora:formatHora(e.target.value)}))} />
          </div>
          <Inp label="Fecha del comprobante" type="date" value={form.fechaComp||""} onChange={e => setForm(f => ({...f, fechaComp:e.target.value}))} />
          <Sel label="Banco" value={form.banco||""} onChange={e => setForm(f => ({...f, banco:e.target.value}))} opts={bancoOpts} />
          <Sel label="Cuenta destino" value={form.cuentaDestino||""} onChange={e => setForm(f => ({...f, cuentaDestino:e.target.value}))} opts={cuentaOpts} />
          <Err msg={err} />
          <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn><Btn v="primary" onClick={guardar}>Guardar cambios</Btn></Acciones>
        </Modal>
      )}
      {delPago && (
        <Modal title="Eliminar registro de pago">
          <Info>Vas a eliminar el pago de <b>{nombreChofer(delPago.choferId)}</b> del día <b>{delPago.fecha}</b> por <b>${delPago.monto}</b> ({delPago.esImputacion ? "imputación a futuro" : "pago normal"}). Esta acción no se puede deshacer.</Info>
          <Acciones><Btn onClick={() => setDelPago(null)}>Cancelar</Btn><Btn v="danger" onClick={confirmarEliminar}>Eliminar definitivamente</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

// ── Panel administrador ───────────────────────────────────
function PanelAdmin({ data, setData }) {
  const [tab, setTab] = useState("alertas");

  const panelVis = [{ id:"alertas",l:"Alertas" },{ id:"choferes",l:"Choferes" },{ id:"inversionistas",l:"Inversionistas" },{ id:"inversiones",l:"Inversiones" }];
  const panelCre = [{ id:"vehiculos",l:"Vehículos" },{ id:"bancos",l:"Bancos" },{ id:"cuentas",l:"Cuentas" },{ id:"deudas",l:"Deudas" },{ id:"usuarios",l:"Usuarios" },{ id:"editarPagos",l:"Editar pagos" }];
  const panelRep = [{ id:"reportes",l:"Reportes Excel" }];

  const TabBtn = ({ id, l }) => (
    <button onClick={() => setTab(id)} style={{ padding:"6px 14px", borderRadius:8, border:`1px solid ${BR}`, cursor:"pointer", fontWeight:tab===id?500:400, background:tab===id?"#4A6FA5":BG, color:tab===id?W:T, fontSize:13 }}>{l}</button>
  );

  return (
    <div>
      <h2 style={{ margin:"0 0 16px", fontSize:18, fontWeight:500, color:T }}>Panel de administrador</h2>
      {data.usuarios.some(u => u.email === "admin@flota.com" && u.password === "admin123" && u.activo) && (
        <Warn>Riesgo de seguridad: el usuario <b>admin@flota.com</b> sigue usando la contraseña de fábrica <b>admin123</b>. Cualquiera que la conozca puede entrar. Cámbiala o elimina ese usuario desde la pestaña <b>Usuarios</b>.</Warn>
      )}
      <div style={{ marginBottom:16 }}>
        <div style={{ fontSize:11, fontWeight:600, color:T2, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:6 }}>Visualización</div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>{panelVis.map(t => <TabBtn key={t.id} {...t} />)}</div>
      </div>
      <div style={{ marginBottom:16 }}>
        <div style={{ fontSize:11, fontWeight:600, color:T2, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:6 }}>Creación y gestión</div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>{panelCre.map(t => <TabBtn key={t.id} {...t} />)}</div>
      </div>
      <div style={{ marginBottom:20 }}>
        <div style={{ fontSize:11, fontWeight:600, color:T2, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:6 }}>Reportes y respaldos</div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>{panelRep.map(t => <TabBtn key={t.id} {...t} />)}</div>
      </div>
      {tab === "alertas"        && <Alertas        data={data} />}
      {tab === "choferes"       && <Choferes       data={data} setData={setData} />}
      {tab === "inversionistas" && <Inversionistas data={data} setData={setData} />}
      {tab === "inversiones"    && <Inversiones    data={data} setData={setData} />}
      {tab === "vehiculos"      && <Vehiculos      data={data} setData={setData} />}
      {tab === "bancos"         && <Bancos         data={data} setData={setData} />}
      {tab === "cuentas"        && <Cuentas        data={data} setData={setData} />}
      {tab === "deudas"         && <Deudas         data={data} setData={setData} />}
      {tab === "usuarios"       && <Usuarios       data={data} setData={setData} />}
      {tab === "editarPagos"    && <EditarPagos    data={data} setData={setData} />}
      {tab === "reportes"       && <Reportes       data={data} setData={setData} />}
    </div>
  );
}

// ── Digitador ────────────────────────────────────────────
function Digitador({ data, setData }) {
  const [fecha,  setFecha]  = useState(hoy());
  const [modal,  setModal]  = useState(false);
  const [selC,   setSelC]   = useState(null);
  const [err,    setErr]    = useState("");
  const [exp,    setExp]    = useState({});
  const [comp,   setComp]   = useState({ numero:"", hora:"", banco:"", cuentaDestino:"", totalComp:"" });
  const [selDias,setSelDias]= useState({});

  const pagoDe     = (cid, f) => data.pagos.find(p => String(p.choferId) === String(cid) && p.fecha === f);
  const pendientes = (cid, hasta) => diasPendientes(data, cid, hasta);

  const resumen = (cid) => {
    const cont = { Cuota:0, Multa:0, Préstamo:0 };
    diasPendientes(data, cid, fecha).forEach(p => { p.tipos.forEach(t => { cont[t] = (cont[t]||0) + 1; }); });
    return cont;
  };

  const abrirModal = (c) => {
    setSelC(c);
    setComp({ numero:"", hora:"", banco:"", cuentaDestino:"", totalComp:"" });
    setSelDias({}); setErr(""); setModal(true);
  };

  const totalComp = round2(Number(comp.totalComp || 0));
  const totalAsig = round2(Object.values(selDias).reduce((s,d) => s + Number(d.monto||0), 0));
  const saldo     = round2(totalComp - totalAsig);
  const diasSel   = Object.keys(selDias);
  const pendHoy   = selC ? pendientes(selC.id, fecha) : [];
  const hayAtras  = pendHoy.some(p => !selDias[p.fecha]);
  const futuros   = selC ? (() => {
    // Hasta el fin real de las deudas activas del chofer (antes estaba limitado a 180 días y 90 filas,
    // lo que ocultaba días registrables cuando el rango de la deuda era largo).
    const finMax = data.deudas
      .filter(d => String(d.choferId)===String(selC.id) && d.activa)
      .map(d => d.fechaFin).filter(Boolean).sort().pop();
    if (!finMax || finMax <= fecha) return [];
    return diasPendientes(data, selC.id, finMax).filter(p => p.fecha > fecha);
  })() : [];
  const combinado = [...pendHoy, ...futuros];
  const futSelez  = futuros.some(p => selDias[p.fecha]);

  const toggleDia = (f, montoBase) => {
    setSelDias(prev => {
      if (prev[f]) { const next = {...prev}; delete next[f]; return next; }
      const yaAsig    = Object.values(prev).reduce((s,d) => s + Number(d.monto||0), 0);
      const disponible= Math.max(0, totalComp - yaAsig);
      const mInit     = round2(Math.min(montoBase, disponible));
      return { ...prev, [f]:{ monto:mInit, estado: mInit >= montoBase ? "pagado" : "abono" } };
    });
  };

  const updateDia = (f, key, val) => {
    if (key === "monto") {
      const otrosAsig  = Object.entries(selDias).filter(([k]) => k!==f).reduce((s,[,d]) => s + Number(d.monto||0), 0);
      const maxPerm    = round2(Math.max(0, totalComp - otrosAsig));
      const limitado   = round2(Math.min(Number(val), maxPerm));
      const debeBase   = combinado.find(p => p.fecha===f)?.monto || 0;
      setSelDias(prev => ({ ...prev, [f]:{ ...prev[f], monto:limitado, estado: limitado >= debeBase ? "pagado" : "abono" } }));
    } else {
      setSelDias(prev => ({ ...prev, [f]:{ ...prev[f], [key]:val } }));
    }
  };

  const guardar = () => {
    setErr("");
    if (!totalComp || totalComp <= 0)                                              { setErr("Ingresa el monto total del comprobante."); return; }
    if (vacio(comp, ["numero","hora","banco","cuentaDestino"]).length)             { setErr("Completa todos los campos del comprobante."); return; }
    if (!horaValida(comp.hora))                                                    { setErr("Ingresa una hora válida en formato HH:MM (ej. 14:30)."); return; }
    if (diasSel.length === 0)                                                      { setErr("Selecciona al menos un día pendiente."); return; }
    if (diasSel.some(f => !selDias[f].monto || Number(selDias[f].monto) <= 0))     { setErr("Ingresa el monto para cada día seleccionado."); return; }
    if (saldo < 0)                                                                 { setErr(`El total asignado supera el comprobante en ${Math.abs(saldo).toFixed(2)}. Ajusta los montos.`); return; }
    if (saldo > 0 && hayAtras)                                                     { setErr(`Hay días atrasados sin seleccionar. Asigna el saldo (${saldo.toFixed(2)}) a los días pendientes.`); return; }
    if (saldo > 0 && !hayAtras && futuros.length > 0)                              { setErr(`Queda un saldo a favor de ${saldo.toFixed(2)}. Imputa el saldo a una o más fechas futuras con saldo pendiente.`); return; }
    if (comprobanteDuplicado(data, comp.numero, comp.banco, fecha, comp.hora))     { setErr("Comprobante duplicado: mismo número, banco, fecha y hora ya registrados."); return; }

    const nuevos = diasSel.map(f => ({
      id: Date.now() + Math.random(),
      choferId: selC.id, fecha:f, fechaComp:fecha, hora:comp.hora,
      monto: Number(selDias[f].monto), estado: selDias[f].estado,
      tipo: (combinado.find(p => p.fecha===f)?.tipos || []).join(" + "),
      comprobante: comp.numero, banco: comp.banco, cuentaDestino: comp.cuentaDestino,
      esImputacion: f > fecha,
    }));

    setData(d => ({ ...d, pagos: [...d.pagos, ...nuevos] }));
    setModal(false);
  };

  const ecMap = { pagado:"green", abono:"amber", no_pagado:"red", condonado:"purple" };
  const elMap = { pagado:"Pagado", abono:"Abono", no_pagado:"No pagado", condonado:"Condonado" };
  const tcMap = { Cuota:"blue", Multa:"red", Préstamo:"amber" };

  // Fila reutilizable: seleccionar un día pendiente y asignarle un monto
  const FilaDia = (p) => {
    const sel       = selDias[p.fecha];
    const otrosAsig = Object.entries(selDias).filter(([k])=>k!==p.fecha).reduce((s,[,d])=>s+Number(d.monto||0),0);
    const dispLeft  = round2(Math.max(0, totalComp - otrosAsig));
    const sinSaldo  = !sel && dispLeft <= 0;
    return (
      <tr key={p.fecha} style={{ background: sel ? "#f0f7ff" : sinSaldo ? "#fafafa" : W }}>
        <td style={{ ...TD, width:32 }}>
          <input type="checkbox" checked={!!sel} disabled={sinSaldo} onChange={() => toggleDia(p.fecha, p.monto)} style={{ cursor: sinSaldo?"not-allowed":"pointer", opacity: sinSaldo?0.4:1 }} />
        </td>
        <td style={TD}>{p.fecha}</td>
        <td style={TD}><div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>{p.tipos.map(t => <Tag key={t} color={tcMap[t]||"gray"}>{t}</Tag>)}</div></td>
        <td style={TD}>${p.monto}</td>
        <td style={{ ...TD, width:90 }}>
          {sel
            ? <input type="number" min="0.01" step="0.01" max={round2(dispLeft + Number(sel.monto||0))} value={sel.monto}
                onChange={e => updateDia(p.fecha, "monto", e.target.value)}
                style={{ width:"100%", background:W, color:T, border:`1px solid ${BR}`, borderRadius:6, padding:"3px 6px", fontSize:13 }} />
            : <span style={{ color:T2 }}>—</span>}
        </td>
        <td style={{ ...TD, width:100 }}>
          {sel ? <Tag color={sel.estado==="pagado"?"green":"amber"}>{sel.estado==="pagado"?"Pago total":"Abono"}</Tag> : <span style={{ color:T2 }}>—</span>}
        </td>
      </tr>
    );
  };

  const exportarAgenda = () => {
    const filas = [];
    data.choferes.forEach(c => {
      pendientes(c.id, fecha).forEach(p => {
        const reg = pagoDe(c.id, p.fecha);
        filas.push({
          Chofer:c.nombres, Fecha:p.fecha, Tipos:p.tipos.join(" + "), "Saldo ($)":p.monto,
          Estado: !reg ? "Sin registrar" : (p.abonado>0 ? "Abono parcial" : (elMap[reg.estado] || reg.estado)),
        });
      });
    });
    descargarLibro([{ nombre:"Agenda de cobros", filas }], `agenda_cobros_${fecha}.xlsx`);
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, marginBottom:16 }}>
        <h2 style={{ margin:0, fontSize:18, fontWeight:500, color:T }}>Agenda de cobros</h2>
        <Btn v="primary" onClick={exportarAgenda}>Descargar agenda (Excel)</Btn>
      </div>
      <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:16 }}>
        <label style={{ fontSize:14, color:T2 }}>Hasta la fecha:</label>
        <input type="date" value={fecha} onChange={e => setFecha(e.target.value)} style={{ background:W, color:T, border:`1px solid ${BR}`, borderRadius:8, padding:"6px 10px" }} />
      </div>

      {data.choferes.length === 0 && <Vacío texto="No hay choferes registrados." />}

      {data.choferes.map(c => {
        const pend  = pendientes(c.id, fecha);
        const res   = resumen(c.id);
        const total = pend.reduce((s,p) => s + p.monto, 0);
        const vehs  = [...new Set(data.deudas.filter(d => String(d.choferId)===String(c.id) && d.activa).map(d=>d.vehiculoId))].map(vid => data.vehiculos.find(v=>v.id===Number(vid))?.placa).filter(Boolean);
        return (
          <div key={c.id} style={{ ...card, marginBottom:12 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
              <div>
                <b style={{ fontSize:15, color:T }}>{c.nombres}</b>
                {vehs.length > 0 && <span style={{ fontSize:13, color:T2, marginLeft:8 }}>· {vehs.join(", ")}</span>}
                <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginTop:6 }}>
                  {res.Cuota    > 0 && <Tag color="blue">Cuotas: {res.Cuota}</Tag>}
                  {res.Multa    > 0 && <Tag color="red">Multas: {res.Multa}</Tag>}
                  {res.Préstamo > 0 && <Tag color="amber">Préstamos: {res.Préstamo}</Tag>}
                  {pend.length === 0 && <Tag color="green">Al día</Tag>}
                </div>
              </div>
              <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
                {pend.length > 0 && <div style={{ textAlign:"right" }}><div style={{ fontSize:18, fontWeight:500, color:"#c0392b" }}>${total.toFixed(2)}</div><div style={{ fontSize:12, color:T2 }}>total pendiente</div></div>}
                {pend.length > 0 && <Btn onClick={() => setExp(e => ({...e,[c.id]:!e[c.id]}))}>{exp[c.id]?"Ocultar":"Ver detalle"}</Btn>}
                <Btn v="primary" onClick={() => abrirModal(c)}>Registrar pago</Btn>
              </div>
            </div>
            {exp[c.id] && pend.length > 0 && (
              <div style={{ marginTop:12, borderTop:`1px solid ${BR}`, paddingTop:10 }}>
                <div style={{ fontSize:13, fontWeight:500, color:T2, marginBottom:6 }}>Días pendientes hasta {fecha}</div>
                <div style={{ overflowX:"auto", maxHeight:240, overflowY:"auto", border:`1px solid ${BR}`, borderRadius:8 }}>
                  <table style={TBL}>
                    <thead><tr>{["Fecha","Tipos","Saldo","Estado"].map(h => <th key={h} style={{ ...TH, position:"sticky", top:0 }}>{h}</th>)}</tr></thead>
                    <tbody>{pend.map(p => {
                      const reg = pagoDe(c.id, p.fecha);
                      return (
                        <tr key={p.fecha}>
                          <td style={TD}>{p.fecha}</td>
                          <td style={TD}><div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>{p.tipos.map(t => <Tag key={t} color={tcMap[t]||"gray"}>{t}</Tag>)}</div></td>
                          <td style={TD}>${p.monto}{p.abonado > 0 && <span style={{ color:T2, fontSize:11 }}> (abonado ${p.abonado} de ${p.montoTotal})</span>}</td>
                          <td style={TD}><Tag color={p.abonado>0?"amber":!reg?"gray":ecMap[reg.estado]||"gray"}>{p.abonado>0?"Abono parcial":!reg?"Sin registrar":elMap[reg.estado]||reg.estado}</Tag></td>
                        </tr>
                      );
                    })}</tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        );
      })}

      {modal && selC && (
        <Modal title={`Registrar pago — ${selC.nombres}`}>
          {/* Comprobante */}
          <div style={{ ...card, background:BG, marginBottom:14 }}>
            <div style={{ fontSize:13, fontWeight:500, color:T, marginBottom:10 }}>Datos del comprobante</div>
            <Inp label="Monto total del comprobante ($)" req type="number" min="0.01" step="0.01"
              value={comp.totalComp} onChange={e => setComp(c => ({...c, totalComp:e.target.value}))} />
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0 12px" }}>
              <Inp label="N° Comprobante" req inputMode="numeric" value={comp.numero} onChange={e => setComp(c => ({...c, numero:e.target.value}))} />
              <Inp label="Hora (HH:MM)" req inputMode="numeric" placeholder="14:30" maxLength={5} value={comp.hora} onChange={e => setComp(c => ({...c, hora:formatHora(e.target.value)}))} />
            </div>
            <Sel label="Banco del comprobante" req value={comp.banco} onChange={e => setComp(c => ({...c, banco:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, ...data.bancos.map(b => ({v:b.nombre, l:b.nombre}))]} />
            <Sel label="Cuenta destino" req value={comp.cuentaDestino} onChange={e => setComp(c => ({...c, cuentaDestino:e.target.value}))} opts={[{v:"",l:"Seleccionar..."}, ...data.cuentas.map(c => ({v:c.numero, l:`${c.titular} · ${c.numero}`}))]} />
          </div>

          {/* Días pendientes hasta el corte */}
          <div style={{ fontSize:13, fontWeight:500, color:T, marginBottom:8 }}>Días pendientes hasta {fecha} — selecciona los que cubre este comprobante</div>
          {pendHoy.length === 0
            ? <Vacío texto="Sin días pendientes hasta esta fecha." />
            : (
              <div style={{ maxHeight:260, overflowY:"auto", border:`1px solid ${BR}`, borderRadius:8, marginBottom:12 }}>
                <table style={TBL}>
                  <thead><tr>{["","Fecha","Tipos","Debe","Asigna","Estado"].map(h => <th key={h} style={{ ...TH, position:"sticky", top:0 }}>{h}</th>)}</tr></thead>
                  <tbody>{pendHoy.map(FilaDia)}</tbody>
                </table>
              </div>
            )
          }

          {/* Imputación del saldo a favor a fechas futuras con saldo pendiente */}
          {!hayAtras && totalComp > 0 && futuros.length > 0 && (saldo > 0 || futSelez) && (
            <div style={{ marginBottom:12 }}>
              <div style={{ fontSize:13, fontWeight:500, color:"#856404", marginBottom:6 }}>Saldo a favor — elige a qué fecha(s) futura(s) con saldo pendiente imputarlo</div>
              <div style={{ maxHeight:240, overflowY:"auto", border:"1px solid #ffc107", borderRadius:8 }}>
                <table style={TBL}>
                  <thead><tr>{["","Fecha","Tipos","Saldo pend.","Imputa","Estado"].map(h => <th key={h} style={{ ...TH, position:"sticky", top:0 }}>{h}</th>)}</tr></thead>
                  <tbody>{futuros.map(FilaDia)}</tbody>
                </table>
              </div>
            </div>
          )}

          {/* Panel de saldo */}
          {totalComp > 0 && (
            <div style={{ ...card, background:BG, marginBottom:12 }}>
              <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8, fontSize:13 }}>
                <span>Comprobante: <b>${totalComp.toFixed(2)}</b></span>
                <span>Asignado: <b>${totalAsig.toFixed(2)}</b></span>
                <span>Saldo: <b style={{ color: saldo<0?"#c0392b":saldo>0?"#856404":"#15803d" }}>${saldo.toFixed(2)}</b></span>
              </div>
              {saldo > 0 && hayAtras && (
                <div style={{ marginTop:10, padding:"8px 12px", background:"#fde8e8", border:"1px solid #f5c6cb", borderRadius:8 }}>
                  <div style={{ fontSize:13, color:"#c0392b" }}>Hay días atrasados sin seleccionar. Asigna el saldo (${saldo.toFixed(2)}) a los días pendientes antes de imputar a fechas futuras.</div>
                </div>
              )}
              {saldo > 0 && !hayAtras && futuros.length > 0 && (
                <div style={{ marginTop:10, padding:"8px 12px", background:"#fef3cd", border:"1px solid #ffc107", borderRadius:8 }}>
                  <div style={{ fontSize:13, color:"#856404" }}>Saldo a favor de ${saldo.toFixed(2)}. Selecciónalo en la tabla de fechas futuras con saldo pendiente (puedes repartirlo entre varias).</div>
                </div>
              )}
              {saldo > 0 && !hayAtras && futuros.length === 0 && (
                <div style={{ marginTop:10, padding:"8px 12px", background:"#fde8e8", border:"1px solid #f5c6cb", borderRadius:8 }}>
                  <div style={{ fontSize:13, color:"#c0392b" }}>Saldo a favor de ${saldo.toFixed(2)} sin fechas pendientes a las cuales imputar. Reduce el monto del comprobante.</div>
                </div>
              )}
              {saldo < 0 && (
                <div style={{ marginTop:10, padding:"8px 12px", background:"#fde8e8", border:"1px solid #f5c6cb", borderRadius:8 }}>
                  <div style={{ fontSize:13, color:"#c0392b" }}>El total asignado supera el comprobante en ${Math.abs(saldo).toFixed(2)}. Ajusta los montos.</div>
                </div>
              )}
              {saldo === 0 && totalAsig > 0 && (
                <div style={{ marginTop:10, padding:"8px 12px", background:"#dcfce7", border:"1px solid #bbf7d0", borderRadius:8 }}>
                  <div style={{ fontSize:13, color:"#15803d" }}>✓ Comprobante completamente distribuido.</div>
                </div>
              )}
            </div>
          )}

          <Err msg={err} />
          <Acciones><Btn onClick={() => setModal(false)}>Cancelar</Btn><Btn v="primary" onClick={guardar}>Guardar registro</Btn></Acciones>
        </Modal>
      )}
    </div>
  );
}

// ── Vista Chofer ─────────────────────────────────────────
function VistaChofer({ data, choferId }) {
  const [tab, setTab] = useState("vencidos");
  const chofer = data.choferes.find(c => c.id === choferId);
  if (!chofer) return <Vacío texto="Chofer no encontrado." />;

  const deudas = data.deudas.filter(d => String(d.choferId) === String(choferId));
  const ayer   = new Date(hoy() + "T12:00:00Z"); ayer.setUTCDate(ayer.getUTCDate()-1);
  const ayerS  = ayer.toISOString().slice(0,10);

  // Puntualidad: por cada día con cuota requerida hasta ayer, contar si quedó cubierto
  const deudasAct = deudas.filter(d => d.activa);
  const fechasIni = deudasAct.map(d => d.fechaInicio).filter(Boolean).sort();
  const ay = parseF(ayerS);
  let ok = 0, tot = 0;
  let cur = fechasIni.length ? parseF(fechasIni[0]) : null;
  let guard = 0;
  while (cur && ay && cur <= ay && guard++ < 4000) {
    const f = isoF(cur);
    const ed = estadoDia(data, choferId, f);
    if (ed.req > 0) { tot++; if (ed.condonado || ed.restante <= 0) ok++; }
    cur.setDate(cur.getDate() + 1);
  }

  const kpiPct = tot > 0 ? Math.round((ok/tot)*100) : 100;
  const real   = data.pagos.filter(p => String(p.choferId)===String(choferId) && p.estado!=="no_pagado");

  // Pendientes hasta hoy (indicador) y hasta el fin de las deudas (diario completo)
  const finMax   = (deudasAct.map(d => d.fechaFin).filter(Boolean).sort().pop()) || hoy();
  const hastaFin = finMax > hoy() ? finMax : hoy();
  const pendHoy  = diasPendientes(data, choferId, hoy());
  const pend     = diasPendientes(data, choferId, hastaFin).map(p => ({ fecha:p.fecha, monto:p.monto, tipos:p.tipos, montoTotal:p.montoTotal, abonado:p.abonado }));
  // Vencidos: días pendientes hasta hoy. Por vencer: días pendientes con fecha posterior a hoy (sin repetir).
  const vencidos  = pendHoy;
  const porVencer = pend.filter(p => p.fecha > hoy());

  // Agrupar pagos por comprobante para ver cómo se aplicó cada uno
  const compMap = {};
  real.filter(p => p.comprobante).forEach(p => {
    const key = `${p.comprobante}||${p.banco}||${p.fechaComp}||${p.hora}`;
    if (!compMap[key]) compMap[key] = { comprobante:p.comprobante, banco:p.banco, fechaComp:p.fechaComp||"", hora:p.hora||"", cuenta:p.cuentaDestino||"", total:0, items:[] };
    compMap[key].total += Number(p.monto||0);
    compMap[key].items.push(p);
  });
  const comprobantes = Object.values(compMap).sort((a,b) => String(b.fechaComp).localeCompare(String(a.fechaComp)));

  const exportar = () => {
    const realF = real.map(p => ({ Fecha:p.fecha, Hora:p.hora, Comprobante:p.comprobante, Banco:p.banco, "Monto ($)":Number(p.monto), Tipo:p.tipo, Estado:p.estado, Registro:p.esImputacion?"Imputación a futuro":"Pago normal" }));
    const pendF = pend.map(p => ({ Fecha:p.fecha, "Situación": p.fecha<=hoy()?"Vencida":"Por vencer", Tipos:p.tipos.join(" + "), "Abono parcial":p.abonado>0?"Sí":"No" }));
    const compF = [];
    comprobantes.forEach(c => c.items.slice().sort((a,b)=>String(a.fecha).localeCompare(String(b.fecha))).forEach(p => compF.push({
      Comprobante:c.comprobante, Banco:c.banco, "Fecha comprobante":c.fechaComp, Hora:c.hora,
      "Total comprobante ($)":round2(c.total), "Día aplicado":p.fecha, "Monto aplicado ($)":Number(p.monto),
      Tipo:p.tipo, Estado:p.estado, Registro:p.esImputacion?"Imputación a futuro":"Pago normal",
    })));
    descargarLibro([
      { nombre:"Pagos realizados",        filas:realF },
      { nombre:"Pagos pendientes",        filas:pendF },
      { nombre:"Aplicación comprobantes", filas:compF },
    ], `estado_${chofer.nombres.replace(/\s+/g,"_")}_${hoy()}.xlsx`);
  };

  const tabs = [
    { id:"vencidos",     l:"Pagos vencidos" },
    { id:"pendientes",   l:"Pagos pendientes" },
    { id:"comprobantes", l:"Por comprobantes" },
    { id:"realizados",   l:"Pagos realizados" },
  ];
  const TabBtn = ({ id, l }) => (
    <button onClick={() => setTab(id)} style={{ padding:"6px 14px", borderRadius:8, border:`1px solid ${BR}`, cursor:"pointer", fontWeight:tab===id?500:400, background:tab===id?"#3A8A6E":BG, color:tab===id?W:T, fontSize:13 }}>{l}</button>
  );

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, marginBottom:16 }}>
        <h2 style={{ margin:0, fontSize:18, fontWeight:500, color:T }}>Mi cuenta — {chofer.nombres}</h2>
        <Btn v="primary" onClick={exportar}>Descargar mi estado (Excel)</Btn>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:10, marginBottom:16 }}>
        {[["Pagos realizados",real.length,"#15803d"],["Cuotas vencidas (a hoy)",pendHoy.length,pendHoy.length===0?"#15803d":"#c0392b"],["Cuotas pendientes (total)",pend.length,pend.length===0?"#15803d":"#856404"],["Puntualidad",kpiPct+"%",kpiPct>=80?"#15803d":"#c0392b"]].map(([l,v,c]) => (
          <div key={l} style={kpi}><div style={{ fontSize:12, color:T2 }}>{l}</div><div style={{ fontSize:26, fontWeight:500, color:c }}>{v}</div></div>
        ))}
      </div>
      <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginBottom:16 }}>{tabs.map(t => <TabBtn key={t.id} {...t} />)}</div>

      {tab === "realizados" && (
        real.length === 0
          ? <p style={{ color:T2, fontSize:14 }}>Sin pagos registrados.</p>
          : <div style={{ overflowX:"auto", maxHeight:480, overflowY:"auto", background:W, border:`1px solid ${BR}`, borderRadius:8 }}>
              <table style={TBL}>
                <thead><tr>{["Fecha","Hora","Comprobante","Banco","Monto","Tipo","Estado"].map(h => <th key={h} style={{ ...TH, position:"sticky", top:0 }}>{h}</th>)}</tr></thead>
                <tbody>{real.map(p => (
                  <tr key={p.id}>
                    <td style={TD}>{p.fecha}</td><td style={TD}>{p.hora}</td><td style={TD}>{p.comprobante}</td>
                    <td style={TD}>{p.banco}</td><td style={TD}>${p.monto}</td>
                    <td style={TD}>{p.tipo}</td>
                    <td style={TD}><Tag color={p.estado==="pagado"?"green":p.estado==="abono"?"amber":"purple"}>{p.estado==="pagado"?"Pagado":p.estado==="abono"?"Abono":"Condonado"}</Tag></td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
      )}

      {tab === "comprobantes" && (
        <>
          <p style={{ fontSize:12, color:T2, margin:"0 0 8px" }}>Cómo se aplicó cada comprobante a los días de pago.</p>
          {comprobantes.length === 0
            ? <p style={{ color:T2, fontSize:14 }}>Sin comprobantes registrados.</p>
            : comprobantes.map((c,i) => (
                <div key={i} style={{ ...card, marginBottom:10 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
                    <div>
                      <b style={{ color:T }}>Comprobante {c.comprobante}</b>
                      <div style={{ fontSize:13, color:T2 }}>{c.banco}{c.fechaComp && ` · ${c.fechaComp}`}{c.hora && ` ${c.hora}`}</div>
                    </div>
                    <div style={{ textAlign:"right" }}>
                      <div style={{ fontSize:16, fontWeight:500, color:T }}>${round2(c.total).toFixed(2)}</div>
                      <div style={{ fontSize:12, color:T2 }}>total aplicado</div>
                    </div>
                  </div>
                  <div style={{ marginTop:8, overflowX:"auto", border:`1px solid ${BR}`, borderRadius:8 }}>
                    <table style={TBL}>
                      <thead><tr>{["Día aplicado","Monto","Tipo","Estado","Registro"].map(h => <th key={h} style={TH}>{h}</th>)}</tr></thead>
                      <tbody>{c.items.slice().sort((a,b)=>String(a.fecha).localeCompare(String(b.fecha))).map(p => (
                        <tr key={p.id}>
                          <td style={TD}>{p.fecha}</td>
                          <td style={TD}>${p.monto}</td>
                          <td style={TD}>{p.tipo}</td>
                          <td style={TD}><Tag color={p.estado==="pagado"?"green":p.estado==="abono"?"amber":"purple"}>{p.estado==="pagado"?"Pagado":p.estado==="abono"?"Abono":"Condonado"}</Tag></td>
                          <td style={TD}><Tag color={p.esImputacion?"amber":"blue"}>{p.esImputacion?"Imputación":"Normal"}</Tag></td>
                        </tr>
                      ))}</tbody>
                    </table>
                  </div>
                </div>
              ))
          }
        </>
      )}

      {tab === "vencidos" && (
        <>
          <div style={{ display:"flex", flexWrap:"wrap", gap:8, marginBottom:10 }}>
            <Tag color={vencidos.length>0?"red":"green"}>Vencidos a hoy: {vencidos.length}</Tag>
          </div>
          <p style={{ fontSize:12, color:T2, margin:"0 0 8px" }}>Días que ya debían estar pagados a la fecha de hoy (cuotas, préstamos y multas).</p>
          {vencidos.length === 0
            ? <p style={{ color:T2, fontSize:14 }}>No tienes pagos vencidos. Estás al día.</p>
            : <div style={{ overflowX:"auto", maxHeight:480, overflowY:"auto", background:W, border:`1px solid ${BR}`, borderRadius:8 }}>
                <table style={TBL}>
                  <thead><tr>{["Fecha","Tipos","Detalle"].map(h => <th key={h} style={{ ...TH, position:"sticky", top:0 }}>{h}</th>)}</tr></thead>
                  <tbody>{vencidos.map((p,i) => (
                    <tr key={i}>
                      <td style={TD}>{p.fecha}</td>
                      <td style={TD}><div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>{p.tipos.map(t => <Tag key={t} color={t==="Cuota"?"blue":t==="Multa"?"red":"amber"}>{t}</Tag>)}</div></td>
                      <td style={TD}>{p.abonado > 0 ? <span style={{ fontSize:12, color:T2 }}>abono parcial registrado</span> : <span style={{ color:T2 }}>—</span>}</td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
          }
        </>
      )}

      {tab === "pendientes" && (
        <>
          <div style={{ display:"flex", flexWrap:"wrap", gap:8, marginBottom:10 }}>
            <Tag color="amber">Pendientes por vencer: {porVencer.length}</Tag>
          </div>
          <p style={{ fontSize:12, color:T2, margin:"0 0 8px" }}>Días futuros aún por vencer (cuotas, préstamos y multas).</p>
          {porVencer.length === 0
            ? <p style={{ color:T2, fontSize:14 }}>No hay pagos por vencer.</p>
            : <div style={{ overflowX:"auto", maxHeight:480, overflowY:"auto", background:W, border:`1px solid ${BR}`, borderRadius:8 }}>
                <table style={TBL}>
                  <thead><tr>{["Fecha","Tipos","Detalle"].map(h => <th key={h} style={{ ...TH, position:"sticky", top:0 }}>{h}</th>)}</tr></thead>
                  <tbody>{porVencer.map((p,i) => (
                    <tr key={i}>
                      <td style={TD}>{p.fecha}</td>
                      <td style={TD}><div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>{p.tipos.map(t => <Tag key={t} color={t==="Cuota"?"blue":t==="Multa"?"red":"amber"}>{t}</Tag>)}</div></td>
                      <td style={TD}>{p.abonado > 0 ? <span style={{ fontSize:12, color:T2 }}>abono parcial registrado</span> : <span style={{ color:T2 }}>—</span>}</td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
          }
        </>
      )}
    </div>
  );
}

// ── Vista Inversionista ───────────────────────────────────
function VistaInv({ data, invId }) {
  const inv = data.inversionistas.find(i => i.id === invId);
  if (!inv) return <Vacío texto="Inversionista no encontrado." />;

  const im = imMens(Number(inv.monto) || 0), ca = capAnual(Number(inv.monto) || 0);
  const base = parseF(inv.fechaEntrega);
  const diaI = Number(inv.diaPagoIntereses);
  const diaC = Number(inv.diaPagoCapital);

  const cuotas = (!base || !diaI) ? [] : Array.from({ length:60 }, (_,i) => {
    const d = new Date(base); d.setMonth(d.getMonth()+i+1); d.setDate(diaI);
    const f = isoF(d);
    return { n:i+1, fecha:f, monto:im, ok:!!data.pagosInv.find(x => x.invId===invId && x.tipo==="Interés mensual" && x.fecha===f && x.realizado) };
  });

  const caps = (!base || !diaC) ? [] : Array.from({ length:5 }, (_,i) => {
    const d = new Date(base); d.setFullYear(d.getFullYear()+i+1); d.setDate(diaC);
    const f = isoF(d);
    return { y:i+1, fecha:f, monto:ca, ok:!!data.pagosInv.find(x => x.invId===invId && x.tipo==="Capital anual" && x.fecha===f && x.realizado) };
  });

  const tiP = cuotas.filter(c=>c.ok).reduce((s,c)=>s+c.monto,0);
  const caP = caps.filter(c=>c.ok).reduce((s,c)=>s+c.monto,0);

  const exportar = () => {
    const cuotasF = cuotas.map(c => ({ "N°":c.n, "Fecha programada":c.fecha, "Monto ($)":round2(c.monto), Estado:c.ok?"Pagado":(c.fecha<=hoy()?"Pendiente":"Próximo") }));
    const capsF   = caps.map(p => ({ "Año":p.y, "Fecha programada":p.fecha, "Monto ($)":round2(p.monto), Estado:p.ok?"Pagado":(p.fecha<=hoy()?"Pendiente":"Próximo") }));
    descargarLibro([{ nombre:"Intereses mensuales", filas:cuotasF }, { nombre:"Capital anual", filas:capsF }], `inversion_${inv.nombres.replace(/\s+/g,"_")}_${hoy()}.xlsx`);
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, marginBottom:16 }}>
        <h2 style={{ margin:0, fontSize:18, fontWeight:500, color:T }}>Mi inversión — {inv.nombres}</h2>
        <Btn v="primary" onClick={exportar}>Descargar mi inversión (Excel)</Btn>
      </div>
      {!base && <Warn>La fecha de entrega o los días de pago no están bien configurados. Pide al administrador que corrija los datos para ver el cronograma.</Warn>}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:10, marginBottom:20 }}>
        {[["Capital total","$"+Number(inv.monto).toLocaleString()],["Fecha entrega",inv.fechaEntrega],["Interés mensual","$"+im.toFixed(2)],["Capital anual (20%)","$"+ca.toFixed(2)],["Intereses cobrados","$"+tiP.toFixed(2)],["Capital cobrado","$"+caP.toFixed(2)]].map(([l,v]) => (
          <div key={l} style={kpi}><div style={{ fontSize:12, color:T2 }}>{l}</div><div style={{ fontSize:15, fontWeight:500, color:T, marginTop:4 }}>{v}</div></div>
        ))}
      </div>
      <h3 style={{ fontSize:15, fontWeight:500, margin:"0 0 8px", color:T }}>Cuotas de interés mensual (60 · ${im.toFixed(2)} c/u)</h3>
      <div style={{ overflowX:"auto", maxHeight:280, overflowY:"auto", background:W, border:`1px solid ${BR}`, borderRadius:8 }}>
        <table style={TBL}>
          <thead><tr>{["#","Fecha programada","Monto","Estado"].map(h => <th key={h} style={{ ...TH, position:"sticky", top:0 }}>{h}</th>)}</tr></thead>
          <tbody>{cuotas.map(c => (
            <tr key={c.n}><td style={TD}>{c.n}</td><td style={TD}>{c.fecha}</td><td style={TD}>${c.monto.toFixed(2)}</td>
              <td style={TD}>{c.ok ? <Tag color="green">Pagado</Tag> : <Tag color={c.fecha<=hoy()?"red":"gray"}>{c.fecha<=hoy()?"Pendiente":"Próximo"}</Tag>}</td>
            </tr>
          ))}</tbody>
        </table>
      </div>
      <h3 style={{ fontSize:15, fontWeight:500, margin:"16px 0 8px", color:T }}>Devoluciones de capital anual (${ca.toFixed(2)}/año)</h3>
      <div style={{ overflowX:"auto", background:W, border:`1px solid ${BR}`, borderRadius:8 }}>
        <table style={TBL}>
          <thead><tr>{["Año","Fecha programada","Monto","Estado"].map(h => <th key={h} style={TH}>{h}</th>)}</tr></thead>
          <tbody>{caps.map(p => (
            <tr key={p.y}><td style={TD}>Año {p.y}</td><td style={TD}>{p.fecha}</td><td style={TD}>${p.monto.toFixed(2)}</td>
              <td style={TD}>{p.ok ? <Tag color="green">Pagado</Tag> : <Tag color={p.fecha<=hoy()?"red":"gray"}>{p.fecha<=hoy()?"Pendiente":"Próximo"}</Tag>}</td>
            </tr>
          ))}</tbody>
        </table>
      </div>
    </div>
  );
}

// ── Login ─────────────────────────────────────────────────
function Login({ data, onLogin }) {
  const [email, setEmail] = useState("");
  const [pass,  setPass]  = useState("");
  const [err,   setErr]   = useState("");
  const [showP, setShowP] = useState(false);

  const login = () => {
    if (!email.trim() || !pass) { setErr("Ingresa correo y contraseña."); return; }
    const u = data.usuarios.find(u => u.email === email.trim() && u.password === pass && u.activo);
    if (u) { setErr(""); onLogin(u); } else setErr("Correo o contraseña incorrectos.");
  };

  return (
    <div style={{ minHeight:480, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:"2rem 1rem", background:BG, borderRadius:16 }}>
      <div style={{ width:"100%", maxWidth:380 }}>
        <div style={{ textAlign:"center", marginBottom:"2rem" }}>
          <div style={{ width:56, height:56, borderRadius:16, background:"#4A6FA5", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 12px" }}><span style={{ fontSize:26, color:W }}>🚌</span></div>
          <h1 style={{ fontSize:20, fontWeight:500, margin:"0 0 4px", color:T }}>Sistema de flota</h1>
          <p style={{ fontSize:13, color:T2, margin:0 }}>Ingresa tus credenciales para continuar</p>
        </div>
        <div data-form-scope style={{ ...card, boxShadow:"0 2px 12px rgba(0,0,0,0.08)" }}>
          <Inp label="Correo electrónico" req type="email" placeholder="usuario@flota.com" value={email} onChange={e => { setEmail(e.target.value); setErr(""); }} />
          <div style={{ marginBottom:10 }}>
            <label style={{ fontSize:13, color:T2, display:"block", marginBottom:3 }}>Contraseña <span style={{ color:"#c0392b" }}>*</span></label>
            <div style={{ position:"relative" }}>
              <input type={showP?"text":"password"} placeholder="••••••••" value={pass}
                onChange={e => { setPass(e.target.value); setErr(""); }} onKeyDown={e => e.key==="Enter" && login()}
                style={{ width:"100%", boxSizing:"border-box", background:W, color:T, border:`1px solid ${BR}`, borderRadius:8, padding:"6px 60px 6px 10px" }} />
              <button type="button" onClick={() => setShowP(s=>!s)} style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", background:"none", border:"none", cursor:"pointer", fontSize:12, color:T2 }}>
                {showP?"Ocultar":"Ver"}
              </button>
            </div>
          </div>
          <Err msg={err} />
          <Btn v="primary" onClick={login} style={{ width:"100%", padding:"10px", fontSize:14 }}>Ingresar</Btn>
        </div>
      </div>
    </div>
  );
}

// ── App ───────────────────────────────────────────────────
export default function App() {
  const [data,    setData]    = useState(INIT);
  const [usuario, setUsuario] = useState(null);
  const [cargando, setCargando] = useState(true);
  const [estado,  setEstado]  = useState("listo"); // listo | guardando | guardado | error
  const listo = useRef(false);
  const rc = { administrador:"#4A6FA5", digitador:"#7B6EA5", chofer:"#3A8A6E", inversionista:"#A56B3A" };

  // Cargar datos guardados al iniciar
  useEffect(() => {
    let activo = true;
    (async () => {
      try {
        const parsed = await loadData();
        if (activo && parsed) {
          if (!parsed.usuarios || parsed.usuarios.length === 0) parsed.usuarios = INIT.usuarios;
          Object.keys(INIT).forEach(k => { if (!parsed[k]) parsed[k] = INIT[k]; });
          setData(parsed);
        }
      } catch (e) {
        console.error("Error al cargar desde Supabase:", e);
      } finally {
        if (activo) { setCargando(false); listo.current = true; }
      }
    })();
    return () => { activo = false; };
  }, []);

  // Guardar automáticamente ante cualquier cambio
  useEffect(() => {
    if (!listo.current) return;
    let activo = true;
    (async () => {
      try {
        setEstado("guardando");
        await saveData(data);
        if (activo) { setEstado("guardado"); setTimeout(() => activo && setEstado("listo"), 1500); }
      } catch (e) {
        if (activo) setEstado("error");
      }
    })();
    return () => { activo = false; };
  }, [data]);

  const estadoTxt = { listo:"", guardando:"Guardando…", guardado:"✓ Guardado", error:"⚠ Error al guardar" };
  const estadoCol = { listo:T2, guardando:"#856404", guardado:"#15803d", error:"#c0392b" };

  if (cargando) {
    return (
      <div style={{ fontFamily:"var(--font-sans)", maxWidth:720, margin:"0 auto", padding:"1rem", background:BG, minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center" }}>
        <div style={{ textAlign:"center", color:T2 }}>
          <div style={{ width:48, height:48, borderRadius:14, background:"#4A6FA5", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 14px" }}><span style={{ fontSize:24 }}>🚌</span></div>
          <div style={{ fontSize:14 }}>Cargando datos guardados…</div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ fontFamily:"var(--font-sans)", color:T, maxWidth:720, margin:"0 auto", padding:"1rem", background:BG, minHeight:"100vh" }}>
      {!usuario
        ? <Login data={data} onLogin={setUsuario} />
        : <>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, ...card }}>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <div style={{ width:36, height:36, borderRadius:"50%", background:rc[usuario.rol]||"#4A6FA5", display:"flex", alignItems:"center", justifyContent:"center", color:W, fontWeight:500, fontSize:14 }}>
                {usuario.nombre.split(" ").map(n=>n[0]).join("").slice(0,2).toUpperCase()}
              </div>
              <div>
                <div style={{ fontWeight:500, fontSize:14, color:T }}>{usuario.nombre}</div>
                <div style={{ fontSize:12, color:T2 }}>{usuario.rol}</div>
              </div>
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:12 }}>
              {estadoTxt[estado] && <span style={{ fontSize:12, color:estadoCol[estado] }}>{estadoTxt[estado]}</span>}
              <Btn onClick={() => setUsuario(null)}>Cerrar sesión</Btn>
            </div>
          </div>
          <div style={{ ...card, padding:"1.5rem" }}>
            {usuario.rol === "administrador" && <PanelAdmin    data={data} setData={setData} />}
            {usuario.rol === "digitador"     && <Digitador     data={data} setData={setData} />}
            {usuario.rol === "chofer"        && <VistaChofer   data={data} choferId={usuario.choferId} />}
            {usuario.rol === "inversionista" && <VistaInv      data={data} invId={usuario.inversionistaId} />}
          </div>
        </>
      }
    </div>
  );
}
