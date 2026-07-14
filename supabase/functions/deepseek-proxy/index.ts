const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

Deno.serve(async (request) => {
  if (request.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    if (request.method !== "POST") {
      return jsonResponse({ error: "Method not allowed" }, 405);
    }

    const body = await request.json();
    const apiKey = Deno.env.get("DEEPSEEK_API_KEY") || String(body.apiKey || "").trim();
    if (!apiKey) {
      return jsonResponse({ error: "Missing DeepSeek API key" }, 400);
    }

    const prompt = String(body.prompt || "").trim();
    const systemPrompt = String(body.systemPrompt || "").trim();
    if (!prompt || !systemPrompt) {
      return jsonResponse({ error: "Missing prompt or systemPrompt" }, 400);
    }

    const messages = [
      { role: "system", content: systemPrompt },
      { role: "user", content: prompt },
    ];
    const deepseekPayload: Record<string, unknown> = {
      model: "deepseek-chat",
      temperature: Number(body.temperature ?? 0.2),
      messages,
    };

    if (body.responseFormat === "json_object") {
      deepseekPayload.response_format = { type: "json_object" };
    }

    const deepseekResponse = await fetch("https://api.deepseek.com/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(deepseekPayload),
    });

    const responseText = await deepseekResponse.text();
    if (!deepseekResponse.ok) {
      return jsonResponse(
        {
          error: responseText || `DeepSeek request failed: ${deepseekResponse.status}`,
        },
        deepseekResponse.status,
      );
    }

    const deepseekData = JSON.parse(responseText);
    const content = deepseekData?.choices?.[0]?.message?.content;
    if (typeof content !== "string") {
      return jsonResponse({ error: "Unexpected DeepSeek response" }, 502);
    }

    return jsonResponse({ content });
  } catch (error) {
    return jsonResponse(
      {
        error: error instanceof Error ? error.message : "Unknown Edge Function error",
      },
      500,
    );
  }
});

function jsonResponse(payload: unknown, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json",
    },
  });
}
