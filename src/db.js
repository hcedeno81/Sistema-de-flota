import { supabase } from "./supabaseClient";

const ROW_ID = "flota_data_v1";

// Lee el documento completo (o null si aún no hay datos guardados)
export async function loadData() {
  const { data, error } = await supabase
    .from("app_state")
    .select("contenido")
    .eq("id", ROW_ID)
    .maybeSingle();
  if (error) throw error;
  return data?.contenido ?? null;
}

// Guarda (inserta o actualiza) el documento completo
export async function saveData(obj) {
  const { error } = await supabase
    .from("app_state")
    .upsert({ id: ROW_ID, contenido: obj, updated_at: new Date().toISOString() });
  if (error) throw error;
}

// Se suscribe a los cambios de la fila. Llama a onChange(contenido) cada vez
// que OTRO usuario guarda. Devuelve una función para cancelar la suscripción.
export function subscribeData(onChange) {
  const channel = supabase
    .channel("app_state_changes")
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "app_state", filter: `id=eq.${ROW_ID}` },
      (payload) => {
        const nuevo = payload.new?.contenido;
        if (nuevo) onChange(nuevo);
      }
    )
    .subscribe();
  return () => supabase.removeChannel(channel);
}
