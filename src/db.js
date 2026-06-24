import { supabase } from "./supabaseClient";

const TABLA   = "flota_registros";
const BLOB_ID = "flota_data_v1"; // documento viejo (app_state), solo para migración única

// Las colecciones de la app. Cada registro vive en su propia fila.
const COLECCIONES = [
  "vehiculos", "choferes", "bancos", "cuentas", "inversionistas",
  "deudas", "pagos", "pagosInv", "gastosInv", "usuarios",
];

// Lee todas las filas ACTIVAS y reconstruye { coleccion: [registros] } + un espejo { coleccion: { id: registro } }
async function leerRegistros() {
  const datos = {}, mapa = {};
  COLECCIONES.forEach(c => { datos[c] = []; mapa[c] = {}; });
  const { data, error } = await supabase
    .from(TABLA)
    .select("coleccion, registro_id, datos, activo")
    .eq("activo", true);
  if (error) throw error;
  (data || []).forEach(fila => {
    if (!datos[fila.coleccion]) { datos[fila.coleccion] = []; mapa[fila.coleccion] = {}; }
    datos[fila.coleccion].push(fila.datos);
    mapa[fila.coleccion][String(fila.registro_id)] = fila.datos;
  });
  return { datos, mapa, total: (data || []).length };
}

// Migra el blob viejo (app_state) a filas individuales. Se ejecuta UNA sola vez (cuando la tabla nueva está vacía).
async function migrarBlob() {
  const { data, error } = await supabase
    .from("app_state").select("contenido").eq("id", BLOB_ID).maybeSingle();
  if (error) throw error;
  const blob = data?.contenido;
  if (!blob || typeof blob !== "object") return false;
  const filas = [];
  COLECCIONES.forEach(col => {
    (blob[col] || []).forEach(r => {
      if (r && r.id != null) filas.push({ coleccion: col, registro_id: String(r.id), datos: r, activo: true });
    });
  });
  if (!filas.length) return false;
  // Insertar en lotes para no exceder límites
  for (let i = 0; i < filas.length; i += 500) {
    const { error: e2 } = await supabase
      .from(TABLA)
      .upsert(filas.slice(i, i + 500), { onConflict: "coleccion,registro_id" });
    if (e2) throw e2;
  }
  return true;
}

// Carga inicial: si la tabla nueva está vacía, intenta migrar el blob viejo y vuelve a leer.
export async function loadData() {
  let r = await leerRegistros();
  if (r.total === 0) {
    const migrado = await migrarBlob();
    if (migrado) r = await leerRegistros();
  }
  return r; // { datos, mapa, total }
}

// Guarda (inserta o actualiza) UN solo registro. No toca ninguna otra fila.
// onConflict apunta a la clave primaria compuesta (coleccion, registro_id).
// .select() OBLIGA a Supabase a devolver la fila escrita: si no devuelve nada,
// el guardado NO se confirmó y lanzamos error (no decimos "guardado" en falso).
export async function guardarRegistro(coleccion, registro) {
  const { data, error } = await supabase
    .from(TABLA)
    .upsert({
      coleccion,
      registro_id: String(registro.id),
      datos: registro,
      activo: true,
      updated_at: new Date().toISOString(),
    }, { onConflict: "coleccion,registro_id" })
    .select();
  if (error) throw error;
  if (!data || data.length === 0)
    throw new Error(`Guardado no confirmado para ${coleccion}/${registro.id}`);
  return data[0].datos;
}

// Borrado SUAVE: marca la fila como inactiva en vez de eliminarla. Nada se pierde de verdad.
// .select() confirma que efectivamente se tocó una fila.
export async function borrarRegistro(coleccion, id) {
  const { data, error } = await supabase
    .from(TABLA)
    .update({ activo: false, updated_at: new Date().toISOString() })
    .eq("coleccion", coleccion)
    .eq("registro_id", String(id))
    .select();
  if (error) throw error;
  if (!data || data.length === 0)
    throw new Error(`Borrado no confirmado para ${coleccion}/${id}`);
}

// Suscripción en tiempo real a cambios de registros individuales.
// Llama onCambio({ col, id, datos, eliminado }) por cada inserción/edición/borrado de OTRO usuario.
export function subscribeRegistros(onCambio) {
  const ch = supabase
    .channel("flota_registros_changes")
    .on("postgres_changes", { event: "*", schema: "public", table: TABLA }, (payload) => {
      const fila = payload.new || payload.old;
      if (!fila || !fila.coleccion) return;
      onCambio({
        col: fila.coleccion,
        id: fila.registro_id,
        datos: fila.datos,
        eliminado: payload.eventType === "DELETE" || fila.activo === false,
      });
    })
    .subscribe();
  return () => supabase.removeChannel(ch);
}
