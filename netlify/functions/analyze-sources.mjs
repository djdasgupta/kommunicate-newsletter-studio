const MAX_SOURCES = 12;
const MAX_TOTAL_CHARS = 90000;

const json = (statusCode, body) => ({
  statusCode,
  headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' },
  body: JSON.stringify(body)
});

function sanitizeSources(input) {
  if (!Array.isArray(input)) return [];
  let remaining = MAX_TOTAL_CHARS;
  return input.slice(0, MAX_SOURCES).map((source, index) => {
    const raw = String(source?.text || '').trim();
    const text = raw.slice(0, Math.max(0, remaining));
    remaining -= text.length;
    return {
      id: String(source?.id || `source_${index + 1}`).slice(0, 120),
      name: String(source?.name || `Source ${index + 1}`).slice(0, 240),
      type: String(source?.type || 'DOCUMENT').slice(0, 40),
      text
    };
  }).filter(source => source.text.length >= 40);
}

const suggestionSchema = {
  type: 'object',
  additionalProperties: false,
  properties: {
    suggestions: {
      type: 'array',
      maxItems: 10,
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          source_id: { type: 'string' },
          source_name: { type: 'string' },
          source_excerpt: { type: 'string' },
          confidence: { type: 'number', minimum: 0, maximum: 1 },
          rationale: { type: 'string' },
          fields: {
            type: 'object',
            additionalProperties: false,
            properties: {
              short_headline: { type: 'string' },
              summary: { type: 'string' },
              category: { type: 'string' },
              jurisdiction: { type: 'string' },
              regulator: { type: 'string' },
              effective_date: { type: 'string' },
              compliance_deadline: { type: 'string' },
              risk_level: { type: 'string', enum: ['', 'Critical', 'High', 'Medium', 'Low'] },
              required_action: { type: 'string' }
            },
            required: ['short_headline', 'summary', 'category', 'jurisdiction', 'regulator', 'effective_date', 'compliance_deadline', 'risk_level', 'required_action']
          }
        },
        required: ['source_id', 'source_name', 'source_excerpt', 'confidence', 'rationale', 'fields']
      }
    }
  },
  required: ['suggestions']
};

export default async (request) => {
  if (request.method !== 'POST') return json(405, { error: 'Method not allowed.' });
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) return json(503, { error: 'AI is not configured. Add OPENAI_API_KEY in Netlify environment variables.' });

  let body;
  try { body = await request.json(); }
  catch { return json(400, { error: 'The request body must be valid JSON.' }); }

  const sources = sanitizeSources(body?.sources);
  if (!sources.length) return json(400, { error: 'No usable source text was provided.' });

  const instruction = String(body?.instruction || '').trim().slice(0, 4000);
  const existingUpdates = Array.isArray(body?.existingUpdates) ? body.existingUpdates.slice(0, 30) : [];
  const focusSuggestion = body?.focusSuggestion || null;
  const sourceBundle = sources.map((source, index) =>
    `SOURCE ${index + 1}\nID: ${source.id}\nNAME: ${source.name}\nTYPE: ${source.type}\nTEXT:\n${source.text}`
  ).join('\n\n---\n\n');

  const task = focusSuggestion
    ? 'Regenerate one improved suggestion for the same underlying regulatory development. Return exactly one suggestion.'
    : 'Identify distinct, material regulatory developments suitable for a corporate compliance newsletter. Avoid duplicates.';

  const prompt = `You are preparing structured regulatory-newsletter proposals for human review.\n\n${task}\n\nUser instruction:\n${instruction || 'Extract material regulatory developments. Use concise, neutral, legally careful language. Preserve explicit dates, jurisdictions, regulator names and required actions. Do not invent facts.'}\n\nExisting newsletter updates (avoid duplicates unless the source clearly updates them):\n${JSON.stringify(existingUpdates)}\n\n${focusSuggestion ? `Prior suggestion to improve:\n${JSON.stringify(focusSuggestion)}\n\n` : ''}Rules:\n- Every factual field must be supported by the named source.\n- Use an empty string when a field is not supported.\n- Headline: ideally 45-100 characters, maximum 180.\n- Summary: ideally 2-4 concise sentences, maximum 1,100 characters.\n- Required action: practical and cautious; do not present legal advice as certainty.\n- source_excerpt must quote or closely preserve a short supporting passage, maximum 650 characters.\n- confidence reflects source clarity and completeness, not importance.\n- Return no more than 10 suggestions.\n\nSources:\n${sourceBundle}`;

  try {
    const response = await fetch('https://api.openai.com/v1/responses', {
      method: 'POST',
      headers: { 'authorization': `Bearer ${apiKey}`, 'content-type': 'application/json' },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || 'gpt-5-mini',
        store: false,
        input: [{ role: 'user', content: [{ type: 'input_text', text: prompt }] }],
        text: {
          format: {
            type: 'json_schema',
            name: 'regulatory_newsletter_suggestions',
            description: 'Human-reviewable regulatory newsletter field suggestions grounded in uploaded sources.',
            strict: true,
            schema: suggestionSchema
          },
          verbosity: 'low'
        }
      })
    });

    const payload = await response.json();
    if (!response.ok) {
      const message = payload?.error?.message || `OpenAI returned ${response.status}.`;
      return json(response.status >= 500 ? 502 : 400, { error: message });
    }
    const outputText = payload?.output_text || payload?.output?.flatMap(item => item?.content || []).find(item => item?.type === 'output_text')?.text;
    if (!outputText) return json(502, { error: 'The AI service returned no structured output.' });
    const parsed = JSON.parse(outputText);
    return json(200, { suggestions: parsed.suggestions || [], mode: 'api', model: process.env.OPENAI_MODEL || 'gpt-5-mini' });
  } catch (error) {
    return json(500, { error: `AI analysis failed: ${error.message}` });
  }
};
