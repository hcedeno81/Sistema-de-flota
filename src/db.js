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
