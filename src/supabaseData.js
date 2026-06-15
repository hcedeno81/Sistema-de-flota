// src/supabaseData.js
// Carga toda la información desde Supabase y la mantiene sincronizada.
import { useState, useEffect, useRef } from "react";
import { supabase } from "./supabaseClient";

// Clave en la app  ->  nombre de la tabla en Supabase
const TABLAS = {
  vehiculos:      "vehiculos",
  choferes:       "choferes",
  bancos:         "bancos",
  cuentas:        "cuentas",
  inversionistas: "inversionistas",
  deudas:         "deudas",
  pagos:          "pagos",
  pagosInv:       "pagos_inv",
  gastosInv:      "gastos_inv",
  usuarios:       "usuarios",
};

export const VACIO = {
  vehiculos: [], choferes: [], bancos: [], cuentas: [], inversionistas: [],
  deudas: [], pagos: [], pagosInv: [], gastosInv: [], usuarios: [],
};

const NUM_KEYS = ["monto", "diaPagoIntereses", "diaPagoCapital"];
const REF_KEYS = ["id", "vehiculoId", "choferId", "invId", "inversionistaId"];

// Al leer: quita la columna interna y convierte los montos a número
const desdeDB = (row) => {
  const { created_at, ...r } = row;
  NUM_KEYS.forEach(k => { if (k in r && r[k] !== null && r[k] !== "") r[k] = Number(r[k]); });
  return r;
};

// Al escribir: asegura que los identificadores viajen como texto
const haciaDB = (row) => {
  const r = { ...row };
  REF_KEYS.forEach(k => { if (k in r && r[k] !== null && r[k] !== undefined && r[k] !== "") r[k] = String(r[k]); });
  return r;
};

// Trae todas las filas de una tabla (paginando de 1000 en 1000)
async function traerTabla(tabla) {
  const todo = [];
  let desde = 0;
  const tam = 1000;
  for (;;) {
    const { data: rows, error } = await supabase.from(tabla).select("*").range(desde, desde + tam - 1);
    if (error) throw error;
    todo.push(...(rows || []));
    if (!rows || rows.length < tam) break;
    desde += tam;
  }
  return todo;
}

async function cargarTodo() {
  const data = { ...VACIO };
  for (const [clave, tabla] of Object.entries(TABLAS)) {
    const rows = await traerTabla(tabla);
    data[clave] = rows.map(desdeDB);
  }
  return data;
}

// Sincroniza una tabla: inserta/actualiza lo nuevo o cambiado y borra lo eliminado
async function sincronizarTabla(tabla, prev, next) {
  const prevPorId = new Map((prev || []).map(r => [String(r.id), r]));
  const nextIds   = new Set((next || []).map(r => String(r.id)));

  const cambiados = (next || []).filter(r => {
    const ant = prevPorId.get(String(r.id));
    return !ant || JSON.stringify(ant) !== JSON.stringify(r);
  });
  if (cambiados.length) {
    const { error } = await supabase.from(tabla).upsert(cambiados.map(haciaDB));
    if (error) throw error;
  }

  const borrar = (prev || []).filter(r => !nextIds.has(String(r.id))).map(r => String(r.id));
  if (borrar.length) {
    const { error } = await supabase.from(tabla).delete().in("id", borrar);
    if (error) throw error;
  }
}

async function sincronizar(prev, next) {
  for (const [clave, tabla] of Object.entries(TABLAS)) {
    await sincronizarTabla(tabla, prev[clave], next[clave]);
  }
}

// Hook que reemplaza el almacenamiento anterior: data, setData, cargando y estado
export function useFlotaData() {
  const [data, setData]       = useState(VACIO);
  const [cargando, setCargando] = useState(true);
  const [estado, setEstado]   = useState("listo");
  const prevRef = useRef(VACIO);
  const listo   = useRef(false);

  // Cargar todo al iniciar
  useEffect(() => {
    let activo = true;
    (async () => {
      try {
        const d = await cargarTodo();
        if (activo) { setData(d); prevRef.current = d; }
      } catch (e) {
        console.error("Error cargando desde Supabase:", e);
        if (activo) setEstado("error");
      } finally {
        if (activo) { setCargando(false); listo.current = true; }
      }
    })();
    return () => { activo = false; };
  }, []);

  // Guardar (sincronizar) ante cada cambio
  useEffect(() => {
    if (!listo.current) return;
    let activo = true;
    (async () => {
      try {
        setEstado("guardando");
        await sincronizar(prevRef.current, data);
        prevRef.current = data;
        if (activo) { setEstado("guardado"); setTimeout(() => activo && setEstado("listo"), 1500); }
      } catch (e) {
        console.error("Error guardando en Supabase:", e);
        if (activo) setEstado("error");
      }
    })();
    return () => { activo = false; };
  }, [data]);

  return { data, setData, cargando, estado };
}