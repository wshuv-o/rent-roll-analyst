import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type, x-supabase-client-platform, x-supabase-client-platform-version, x-supabase-client-runtime, x-supabase-client-runtime-version",
};

const SYSTEM_PROMPT = `You are a rent roll spreadsheet parser. Given an anonymized HTML sample of an Excel rent roll, output ONLY a JSON instruction object. Be concise — no lengthy explanations.

Brief analysis then JSON. Use these sections:

[THINKING]
One-liner per observation: identify header row(s), metadata rows, data start row.

[GROUPING]  
One-liner: how new tenants start (e.g. "suite_id column non-empty"), what continuation rows look like, what to skip.

[PARSING INSTRUCTION]
\`\`\`json
{
  "header_rows": [],
  "data_starts_at_row": null,
  "column_map": {
    "suite_id": "",
    "tenant_name": "",
    "lease_start": "",
    "lease_end": "",
    "gla_sqft": "",
    "monthly_base_rent": "",
    "base_rent_psf": "",
    "recurring_charge_code": "",
    "recurring_charge_amount": "",
    "recurring_charge_psf": "",
    "future_rent_date": "",
    "future_rent_amount": "",
    "future_rent_psf": ""
  },
  "new_tenant_rule": "",
  "skip_row_patterns": [],
  "addon_space_patterns": [],
  "confidence": "high | medium | low",
  "notes": ""
}
\`\`\`

Note: Some columns may be grouped under a shared parent header (e.g. "Base Rent" spanning Monthly and PSF sub-columns). Account for merged/grouped headers when mapping columns.

[FLAGS]
Only list genuine ambiguities. If none, say "None."

IMPORTANT: Keep text minimal. The JSON is what matters.`;

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const { sampleHtml, contextNote } = await req.json();
    const LOVABLE_API_KEY = Deno.env.get("LOVABLE_API_KEY");
    if (!LOVABLE_API_KEY) throw new Error("LOVABLE_API_KEY is not configured");

    const userMessage = `${contextNote}\n\n${sampleHtml}`;

    const response = await fetch("https://ai.gateway.lovable.dev/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${LOVABLE_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "google/gemini-2.5-flash",
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: userMessage },
        ],
        stream: true,
      }),
    });

    if (!response.ok) {
      if (response.status === 429) {
        return new Response(JSON.stringify({ error: "Rate limits exceeded" }), {
          status: 429, headers: { ...corsHeaders, "Content-Type": "application/json" },
        });
      }
      if (response.status === 402) {
        return new Response(JSON.stringify({ error: "Payment required" }), {
          status: 402, headers: { ...corsHeaders, "Content-Type": "application/json" },
        });
      }
      const t = await response.text();
      console.error("AI gateway error:", response.status, t);
      return new Response(JSON.stringify({ error: "AI gateway error" }), {
        status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    return new Response(response.body, {
      headers: { ...corsHeaders, "Content-Type": "text/event-stream" },
    });
  } catch (e) {
    console.error("analyze-rent-roll error:", e);
    return new Response(JSON.stringify({ error: e instanceof Error ? e.message : "Unknown error" }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
