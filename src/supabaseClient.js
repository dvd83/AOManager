import { createClient } from "@supabase/supabase-js";

// Clé "publishable" (anonyme) — sûre à exposer côté client par conception Supabase.
// La sécurité réelle est assurée par les règles Row Level Security côté base de données,
// pas par le secret de cette clé.
const SUPABASE_URL = "https://ztzzxwnqikbtrenomdhb.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp0enp4d25xaWtidHJlbm9tZGhiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk5OTMzODgsImV4cCI6MjEwNTU2OTM4OH0.zCUFanh1piV57nXcCuZPv_juPawboNl8BmrZocs6gzU";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
