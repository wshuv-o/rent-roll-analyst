import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type, x-supabase-client-platform, x-supabase-client-platform-version, x-supabase-client-runtime, x-supabase-client-runtime-version",
};

const SYSTEM_PROMPT = `You are an expert commercial real estate data analyst specializing in parsing rent roll spreadsheets.

You will receive a sample of rows from an Excel rent roll file. The data has been anonymized — tenant names, suite IDs, and amounts have been replaced with placeholder IDs. A mapping exists on the client side to restore real values after you return your output.

Work through the following steps:

[THINKING]
Step 1 — Understand the Layout
Look at the rows. Identify:
- Which row is the actual header row. Note: headers sometimes span two consecutive rows that must be merged into one label.
- Which columns map to which fields — build an explicit column map (e.g. Column C = Tenant Name, Column F = Lease Start)
- What the metadata rows are at the top (report title, date, building name)
Narrate what you see. Think out loud.

[GROUPING]
Step 2 — Identify the Tenant Block Pattern
Explain the rule for where a new tenant starts and where continuation rows appear. Call out summary rows (Total SF, Total PSF) and add-on space rows.
Narrate this with row-level specifics.

[PARSING INSTRUCTION]
Step 3 — Produce a Parsing Instruction Object
Return a JSON object with this exact schema:

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

[FLAGS]
Step 4 — Flag Any Issues
List anything ambiguous, inconsistent, or that an analyst should verify.

Structure your response in labeled sections: [THINKING], [GROUPING], [PARSING INSTRUCTION], [FLAGS]`;

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const { sampleHtml, contextNote } = await req.json();
    const LOVABLE_API_KEY = Deno.env.get("LOVABLE_API_KEY");
    if (!LOVABLE_API_KEY) throw new Error("LOVABLE_API_KEY is not configured");

    const userMessage = `Here is a sample from an Excel rent roll file:\n\n${contextNote}\n\n${sampleHtml}`;

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
        return new Response(JSON.stringify({ error: "Rate limits exceeded, please try again later." }), {
          status: 429,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        });
      }
      if (response.status === 402) {
        return new Response(JSON.stringify({ error: "Payment required, please add funds to your Lovable AI workspace." }), {
          status: 402,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        });
      }
      const t = await response.text();
      console.error("AI gateway error:", response.status, t);
      return new Response(JSON.stringify({ error: "AI gateway error" }), {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    return new Response(response.body, {
      headers: { ...corsHeaders, "Content-Type": "text/event-stream" },
    });
  } catch (e) {
    console.error("analyze-rent-roll error:", e);
    return new Response(JSON.stringify({ error: e instanceof Error ? e.message : "Unknown error" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
