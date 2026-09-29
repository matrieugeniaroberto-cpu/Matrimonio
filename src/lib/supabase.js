import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // eslint-disable-next-line no-console
  console.warn(
    "Faltan VITE_SUPABASE_URL y/o VITE_SUPABASE_ANON_KEY. Revisa tu archivo .env (local) o las variables de entorno en Vercel."
  );
}

export const supabase = createClient(url, anonKey);

// Todo el estado de la app se guarda como un único registro JSON en esta tabla.
export const DATA_TABLE = "wedding_data";
export const DATA_ROW_ID = "main";
