// Gestion des utilisateurs, réservée aux admins. Utilise la clé "service role"
// (injectée automatiquement par Supabase dans les Edge Functions — jamais exposée
// au navigateur) pour lister/révoquer/supprimer des comptes via l'API Admin Auth.
// L'appelant doit être authentifié ET avoir role='admin' dans public.profiles ;
// vérifié ici avant toute opération, indépendamment de ce que dit le frontend.
import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "jsr:@supabase/supabase-js@2";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { ...corsHeaders, "Content-Type": "application/json" } });
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders });

  const authHeader = req.headers.get("Authorization") || "";
  const callerClient = createClient(SUPABASE_URL, SERVICE_ROLE_KEY, { global: { headers: { Authorization: authHeader } } });
  const { data: { user: caller } } = await callerClient.auth.getUser();
  if (!caller) return json({ error: "Non authentifié." }, 401);

  const admin = createClient(SUPABASE_URL, SERVICE_ROLE_KEY);
  const { data: profile } = await admin.from("profiles").select("role").eq("id", caller.id).single();
  if (profile?.role !== "admin") return json({ error: "Réservé aux administrateurs." }, 403);

  let body: { action?: string; userId?: string };
  try { body = await req.json(); } catch { return json({ error: "Corps de requête JSON invalide." }, 400); }

  if (body.action === "list") {
    const { data: profiles, error: pErr } = await admin.from("profiles").select("id, email, role, created_at").order("created_at");
    if (pErr) return json({ error: pErr.message }, 500);
    const { data: authList, error: aErr } = await admin.auth.admin.listUsers({ perPage: 200 });
    if (aErr) return json({ error: aErr.message }, 500);
    const banned = new Map(authList.users.map(u => [u.id, !!u.banned_until && new Date(u.banned_until) > new Date()]));
    const { data: counts } = await admin.from("tenders").select("user_id");
    const countByUser = new Map<string, number>();
    (counts || []).forEach((r: { user_id: string }) => countByUser.set(r.user_id, (countByUser.get(r.user_id) || 0) + 1));
    return json({ users: (profiles || []).map(p => ({ ...p, revoked: banned.get(p.id) || false, tenderCount: countByUser.get(p.id) || 0 })) });
  }

  if (!body.userId) return json({ error: "userId requis." }, 400);
  if (body.userId === caller.id) return json({ error: "Impossible d'appliquer cette action à votre propre compte." }, 400);

  if (body.action === "revoke") {
    const { error } = await admin.auth.admin.updateUserById(body.userId, { ban_duration: "876000h" });
    if (error) return json({ error: error.message }, 500);
    return json({ ok: true });
  }
  if (body.action === "restore") {
    const { error } = await admin.auth.admin.updateUserById(body.userId, { ban_duration: "none" });
    if (error) return json({ error: error.message }, 500);
    return json({ ok: true });
  }
  if (body.action === "delete") {
    const { error } = await admin.auth.admin.deleteUser(body.userId);
    if (error) return json({ error: error.message }, 500);
    return json({ ok: true });
  }

  return json({ error: "Action inconnue." }, 400);
});
