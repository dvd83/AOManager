// Proxy serveur vers l'API Anthropic — la clé API ne quitte jamais le serveur.
// Le frontend appelle cette fonction via supabase.functions.invoke("ai-assist", ...),
// authentifié par le JWT Supabase de l'utilisateur connecté (verify_jwt activé).
import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY");
const MODEL = "claude-sonnet-5";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders });

  if (!ANTHROPIC_API_KEY) {
    return new Response(JSON.stringify({ error: "ANTHROPIC_API_KEY n'est pas configurée côté serveur." }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  let body: { prompt?: string; maxTokens?: number };
  try {
    body = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: "Corps de requête JSON invalide." }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const prompt = (body.prompt || "").trim();
  if (!prompt) {
    return new Response(JSON.stringify({ error: "Le champ 'prompt' est requis." }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
  const maxTokens = Math.min(Math.max(Number(body.maxTokens) || 1000, 1), 4096);

  const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({ model: MODEL, max_tokens: maxTokens, messages: [{ role: "user", content: prompt }] }),
  });

  const data = await anthropicRes.json().catch(() => null);
  if (!anthropicRes.ok || !data) {
    return new Response(JSON.stringify({ error: data?.error?.message || `Erreur API Anthropic (${anthropicRes.status}).` }), {
      status: anthropicRes.status || 502,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const text = (data.content || []).filter((b: { type: string }) => b.type === "text").map((b: { text: string }) => b.text).join("\n");

  return new Response(JSON.stringify({ text }), {
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});
