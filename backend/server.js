require('dotenv').config();
const express = require('express');
const mysql = require('mysql2/promise');
const bodyParser = require('body-parser');
const cors = require('cors');
const OpenAI = require('openai');
const { createClient } = require('@supabase/supabase-js');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = 3008;

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
);

app.use(bodyParser.json({ limit: '10mb' }));
app.use(express.json());
app.use(cors({
  origin: ["https://lineitems.bulkscraper.cloud", "https://clickycube.lovable.app","https://rent-roll-oracle.lovable.app"],
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
}));

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const pool = mysql.createPool({
  host: process.env.DB_HOST,
  port: process.env.DB_PORT,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,
  waitForConnections: true,
  connectionLimit: 10,
});

async function trackEmbeddingUsage(response, endpoint = 'unknown') {
  await supabase.from('token_usage').insert({
    model: 'text-embedding-3-large',
    endpoint,
    total_tokens: response?.usage?.total_tokens || 0,
  });
}

async function trackChatUsage(response, endpoint = 'unknown') {
  await supabase.from('token_usage').insert({
    model: 'gpt-4o-mini',
    endpoint,
    prompt_tokens: response?.usage?.prompt_tokens || 0,
    completion_tokens: response?.usage?.completion_tokens || 0,
    total_tokens: response?.usage?.total_tokens || 0,
  });
}


// ────────────────────────────────────────────────
// GET /api/token-usage - View token usage stats
// DELETE /api/token-usage - Reset all rows
// ────────────────────────────────────────────────
app.get('/api/token-usage', async (req, res) => {
  const { data, error } = await supabase
    .from('token_usage')
    .select('model, endpoint, prompt_tokens, completion_tokens, total_tokens, created_at');

  if (error) return res.status(500).json({ error: error.message });

  const COSTS = {
    'text-embedding-3-large': { total: 0.13 },
    'gpt-4o-mini': { input: 0.15, output: 0.60 },
  };

  // Aggregate by model
  const byModel = {};
  data.forEach(row => {
    if (!byModel[row.model]) {
      byModel[row.model] = { requests: 0, prompt_tokens: 0, completion_tokens: 0, total_tokens: 0 };
    }
    byModel[row.model].requests += 1;
    byModel[row.model].prompt_tokens += row.prompt_tokens || 0;
    byModel[row.model].completion_tokens += row.completion_tokens || 0;
    byModel[row.model].total_tokens += row.total_tokens || 0;
  });

  // Aggregate by endpoint
  const byEndpoint = {};
  data.forEach(row => {
    if (!byEndpoint[row.endpoint]) {
      byEndpoint[row.endpoint] = { requests: 0, total_tokens: 0 };
    }
    byEndpoint[row.endpoint].requests += 1;
    byEndpoint[row.endpoint].total_tokens += row.total_tokens || 0;
  });

  // Calculate costs
  let totalCost = 0;
  Object.entries(byModel).forEach(([model, stats]) => {
    if (model === 'text-embedding-3-large') {
      stats.estimated_cost_usd = ((stats.total_tokens / 1_000_000) * COSTS[model].total).toFixed(6);
      totalCost += parseFloat(stats.estimated_cost_usd);
    } else if (model === 'gpt-4o-mini') {
      const cost =
        (stats.prompt_tokens / 1_000_000) * COSTS[model].input +
        (stats.completion_tokens / 1_000_000) * COSTS[model].output;
      stats.estimated_cost_usd = cost.toFixed(6);
      totalCost += cost;
    }
  });

  res.json({
    by_model: byModel,
    by_endpoint: byEndpoint,
    total_estimated_cost_usd: totalCost.toFixed(6),
    total_rows: data.length,
  });
});

app.delete('/api/token-usage', async (req, res) => {
  const { error } = await supabase
    .from('token_usage')
    .delete()
    .neq('id', '00000000-0000-0000-0000-000000000000'); // matches all rows

  if (error) return res.status(500).json({ error: error.message });
  res.json({ message: 'Token usage cleared.' });
});


// ────────────────────────────────────────────────
// POST /api/test-mapping
// ────────────────────────────────────────────────
app.post('/api/test-mapping', async (req, res) => {
  const { itemsToMap, targetSchema } = req.body;

  if (!itemsToMap || !targetSchema ||
      !itemsToMap.income || !Array.isArray(itemsToMap.income) ||
      !itemsToMap.expense || !Array.isArray(itemsToMap.expense)) {
    return res.status(400).json({ error: 'Invalid payload: itemsToMap must contain income & expense arrays' });
  }

  try {
    // ── STEP 1: Fetch metadata for dynamic keyword extraction ──
    const { data: supabaseItems, error: fetchError } = await supabase
      .from('lineitems')
      .select('account_name, type, target_key');

    if (fetchError) throw fetchError;
    if (!supabaseItems || supabaseItems.length === 0) {
      return res.status(500).json({ error: 'No reference line items in database' });
    }

    console.log(`Fetched ${supabaseItems.length} reference items for keyword extraction`);

    const keywordMap = {};
    for (const item of supabaseItems) {
      const { type, target_key, account_name } = item;
      if (!type || !target_key || !account_name) continue;
      const category = type.toLowerCase();
      const normalizedKey = target_key.toLowerCase().replace(/\s+/g, '_');
      const fullPath = `${category}.${normalizedKey}`;
      if (!keywordMap[fullPath]) keywordMap[fullPath] = [];
      keywordMap[fullPath].push(account_name.toLowerCase());
    }

    const extractedKeywords = {};
    for (const [fullPath, accounts] of Object.entries(keywordMap)) {
      const wordFreq = {};
      accounts.forEach(acc => {
        const words = acc.split(/\s+/).filter(w => w.length > 3);
        words.forEach(word => { wordFreq[word] = (wordFreq[word] || 0) + 1; });
      });
      const threshold = Math.max(1, Math.floor(accounts.length * 0.1));
      extractedKeywords[fullPath] = Object.entries(wordFreq)
        .filter(([, count]) => count >= threshold)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 20)
        .map(([word]) => word);
    }

    console.log(`Extracted keywords for ${Object.keys(extractedKeywords).length} target keys`);

    const flatItems = [];
    for (const [category, items] of Object.entries(itemsToMap)) {
      items.forEach(source => flatItems.push({ category, source }));
    }

    // ── STEP 2: Keyword matching ──
    const results = [];
    const unmapped = [];

    flatItems.forEach(({ category, source }) => {
      const lower = source.toLowerCase();
      let targetKey = null;
      let confidence = 0;
      let method = "unmapped";
      let bestMatch = null;
      let bestKeywordScore = 0;

      for (const [fullPath, keywords] of Object.entries(extractedKeywords)) {
        if (!fullPath.startsWith(`${category}.`)) continue;
        const matchCount = keywords.filter(kw => lower.includes(kw)).length;
        if (matchCount > bestKeywordScore) {
          bestKeywordScore = matchCount;
          bestMatch = fullPath;
        }
      }

      if (bestKeywordScore >= 3 && extractedKeywords[bestMatch]?.length >= 5) {
        const [, key] = bestMatch.split('.');
        targetKey = key;
        confidence = Math.min(0.95, 0.7 + (bestKeywordScore * 0.05));
        method = "keyword";
      }

      const resultRow = {
        category,
        source,
        target: targetKey,
        full_path: targetKey ? `${category}.${targetKey}` : null,
        confidence: Number(confidence.toFixed(3)),
        method
      };

      results.push(resultRow);

      if (!targetKey) {
        unmapped.push({ category, source, index: results.length - 1 });
      }
    });

    console.log(`After keyword matching: ${unmapped.length} unmapped out of ${flatItems.length}`);

    if (unmapped.length === 0) {
      return res.json({
        metadata: {
          total_income_items: itemsToMap.income.length,
          total_expense_items: itemsToMap.expense.length,
          total_rows: results.length,
          keyword_mapped: results.length,
          embedding_mapped: 0,
          ai_mapped: 0,
          timestamp: new Date().toISOString()
        },
        results
      });
    }

    // ── STEP 3: Generate embeddings for unmapped items ──
    const embeddingPromises = unmapped.map(async (item) => {
      const response = await openai.embeddings.create({
        model: "text-embedding-3-large",
        input: item.source,
        dimensions: 1536,
      });
      trackEmbeddingUsage(response, '/api/test-mapping'); // ✅ track
      return { ...item, embedding: response.data[0].embedding };
    });

    const itemsWithEmbeddings = await Promise.all(embeddingPromises);
    console.log(`Generated embeddings for ${itemsWithEmbeddings.length} unmapped items`);

    // ── STEP 4: Vector search ──
    const embeddingCandidates = [];

    for (const userItem of itemsWithEmbeddings) {
      try {
        const embeddingStr = `[${userItem.embedding.join(',')}]`;

        const { data: matches, error: searchError } = await supabase.rpc('match_lineitems_raw', {
          p_query_embedding: embeddingStr,
          p_match_threshold: 0.45,
          p_match_count: 5,
          p_filter_type: userItem.category.charAt(0).toUpperCase() + userItem.category.slice(1)
        });

        if (searchError) {
          console.error(`❌ Vector search error for "${userItem.source}":`, searchError.message);
          embeddingCandidates.push({ source: userItem.source, category: userItem.category, index: userItem.index, candidates: [] });
          continue;
        }

        if (!matches || matches.length === 0) {
          console.log(`⚠️  No embedding matches found for "${userItem.source}" (category: ${userItem.category})`);
          embeddingCandidates.push({ source: userItem.source, category: userItem.category, index: userItem.index, candidates: [] });
          continue;
        }

        console.log(`\n🔍 Top matches for "${userItem.source}":`);
        matches.forEach((match, idx) => {
          console.log(`   ${idx + 1}. "${match.account_name}" → ${match.target_key} (similarity: ${match.similarity.toFixed(3)})`);
        });

        const topMatch = matches[0];
        const normalizedTargetKey = topMatch.target_key.toLowerCase().replace(/\s+/g, '_');

        if (topMatch.similarity > 0.75) {
          console.log(`   ✅ High confidence — using "${topMatch.target_key}" directly\n`);
          results[userItem.index] = {
            ...results[userItem.index],
            target: normalizedTargetKey,
            full_path: `${userItem.category}.${normalizedTargetKey}`,
            confidence: Number(Math.min(0.95, topMatch.similarity).toFixed(3)),
            method: "embedding",
            matched_account: topMatch.account_name
          };
        } else {
          console.log(`   ⚠️  Low confidence (${topMatch.similarity.toFixed(3)}) — sending to AI\n`);
          embeddingCandidates.push({
            source: userItem.source,
            category: userItem.category,
            index: userItem.index,
            candidates: matches.slice(0, 3).map(m => m.target_key.toLowerCase().replace(/\s+/g, '_')),
            candidateDetails: matches.slice(0, 3).map(m => ({
              target_key: m.target_key,
              account_name: m.account_name,
              similarity: m.similarity
            }))
          });
        }
      } catch (err) {
        console.error(`❌ Vector search failed for "${userItem.source}":`, err.message);
        embeddingCandidates.push({ source: userItem.source, category: userItem.category, index: userItem.index, candidates: [] });
      }
    }

    console.log(`\n📊 Summary: ${results.filter(r => r.method === 'embedding').length} embedding mapped, ${embeddingCandidates.length} going to AI\n`);

    // ── STEP 5: AI fallback ──
    if (embeddingCandidates.length > 0) {
      console.log(`🤖 Sending ${embeddingCandidates.length} items to AI...\n`);
      const BATCH_SIZE = 45;

      for (let i = 0; i < embeddingCandidates.length; i += BATCH_SIZE) {
        const batch = embeddingCandidates.slice(i, i + BATCH_SIZE);

        console.log(`📤 AI Batch ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(embeddingCandidates.length / BATCH_SIZE)}:`);
        batch.forEach(item => {
          if (item.candidateDetails?.length > 0) {
            console.log(`   • "${item.source}" → Candidates from embedding:`);
            item.candidateDetails.forEach((d, idx) => {
              console.log(`      ${idx + 1}. ${d.target_key} (similarity: ${d.similarity.toFixed(3)}, account: "${d.account_name}")`);
            });
          } else {
            console.log(`   • "${item.source}" → No embedding matches, using full schema`);
          }
        });

        const promptItems = batch.map(item => {
          if (item.candidates.length > 0) {
            return `Source: "${item.source}"\nCategory: ${item.category}\nCandidates: ${item.candidates.join(', ')}`;
          } else {
            const schemaKeys = Object.keys(targetSchema[item.category] || {});
            return `Source: "${item.source}"\nCategory: ${item.category}\nAvailable keys: ${schemaKeys.join(', ')}`;
          }
        }).join('\n\n');

        const prompt = `
          You are a financial line-item classifier. For each item, choose the BEST target_key.
          If candidates are provided, choose from those. Otherwise, choose from the available keys.
          ALWAYS choose a key — never return null.

          ${promptItems}

          Return ONLY valid JSON array with NO markdown, NO backticks:
          [
            {"source": "exact source text", "target": "chosen_key", "confidence": 0.3-1.0}
          ]
          `;

        let attempts = 0;
        let parsed = null;

        while (attempts < 3 && !parsed) {
          attempts++;
          try {
            const response = await openai.chat.completions.create({
              model: "gpt-4o-mini",
              messages: [
                { role: "system", content: "Return ONLY valid JSON array. No markdown, no extra text." },
                { role: "user", content: prompt }
              ],
              temperature: 0.2,
              max_tokens: 4096,
            });

            trackChatUsage(response, '/api/test-mapping'); // ✅ track

            const content = response.choices[0].message.content.trim();
            const cleanContent = content.replace(/```json\n?|\n?```/g, '').trim();
            try {
              parsed = JSON.parse(cleanContent);
            } catch (parseErr) {
              console.error(`   ❌ AI parse attempt ${attempts} failed:`, cleanContent.substring(0, 300));
            }
          } catch (err) {
            console.error(`   ❌ AI call attempt ${attempts} failed:`, err.message);
          }
        }

        if (parsed && Array.isArray(parsed)) {
          console.log(`   ✅ AI returned ${parsed.length} results`);
          parsed.forEach(aiResult => {
            const candidate = batch.find(c => c.source === aiResult.source);
            if (candidate && aiResult.target) {
              console.log(`      • "${aiResult.source}" → ${aiResult.target} (confidence: ${aiResult.confidence})`);
              results[candidate.index] = {
                ...results[candidate.index],
                target: aiResult.target,
                full_path: `${candidate.category}.${aiResult.target}`,
                confidence: aiResult.confidence || 0.65,
                method: "embed"
              };
            }
          });
        } else {
          console.log(`   ❌ AI returned invalid response after ${attempts} attempts\n`);
        }
      }
    }

    // Final fallback
    results.forEach(r => {
      if (!r.target) {
        r.target = r.category === "income" ? "other_income" : "other_expense";
        r.full_path = `${r.category}.${r.target}`;
        r.confidence = 0.3;
        r.method = "fallback";
      }
    });

    res.json({
      metadata: {
        total_income_items: itemsToMap.income.length,
        total_expense_items: itemsToMap.expense.length,
        total_rows: results.length,
        keyword_mapped: results.filter(r => r.method === "keyword").length,
        embedding_mapped: results.filter(r => r.method === "embedding").length,
        ai_mapped: results.filter(r => r.method === "embed").length,
        fallback_mapped: results.filter(r => r.method === "fallback").length,
        timestamp: new Date().toISOString()
      },
      results
    });

  } catch (error) {
    console.error('Mapping failed:', error);
    res.status(500).json({ error: 'Failed to map line items', details: error.message || 'Unknown error' });
  }
});


// ────────────────────────────────────────────────
// POST /api/lineitems - Insert a single row with embedding
// ────────────────────────────────────────────────
app.post('/api/lineitems', async (req, res) => {
  const { account_code, account_name, header, description, type, property_type, portfolio = false, target_key } = req.body;

  if (!account_name || !type || !target_key) {
    return res.status(400).json({ error: 'Missing required fields: account_name, type, target_key' });
  }
  if (!['Income', 'Expense', 'Capital'].includes(type)) {
    return res.status(400).json({ error: 'type must be one of: Income, Expense, Capital' });
  }

  try {
    const textToEmbed = [account_name, account_code || '', header || '', description || '', type, target_key].filter(Boolean).join(' | ');

    const embeddingResponse = await openai.embeddings.create({
      model: "text-embedding-3-large",
      input: textToEmbed,
      dimensions: 1536,
    });
    trackEmbeddingUsage(embeddingResponse, '/api/lineitems'); // ✅ track

    const embedding = embeddingResponse.data[0].embedding;

    const { data, error } = await supabase
      .from('lineitems')
      .insert({ id: uuidv4(), account_code: account_code || null, account_name, header: header || null, description: description || null, type, property_type: property_type || null, portfolio, target_key, embedding, created_at: new Date().toISOString() })
      .select()
      .single();

    if (error) throw error;
    res.status(201).json({ message: 'Line item created successfully', data });
  } catch (error) {
    console.error('Insert failed:', error);
    res.status(500).json({ error: 'Failed to insert line item', details: error.message || error.details || 'Unknown error' });
  }
});


// ────────────────────────────────────────────────
// POST /api/lineitems/batch
// ────────────────────────────────────────────────
app.post('/api/lineitems/batch', async (req, res) => {
  const items = req.body;

  if (!Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: 'Request body must be a non-empty array of line item objects' });
  }

  const BATCH_SIZE = 20;
  const results = [];
  const errors = [];

  try {
    for (let i = 0; i < items.length; i += BATCH_SIZE) {
      const batch = items.slice(i, i + BATCH_SIZE);
      const batchResults = [];

      const validations = batch.map((item, idx) => {
        const { account_code, account_name, header, description, type, property_type, portfolio = false, target_key } = item;
        if (!account_name || !type || !target_key) return { index: i + idx, error: 'Missing required fields: account_name, type, target_key', item };
        if (!['Income', 'Expense', 'Capital'].includes(type)) return { index: i + idx, error: 'type must be one of: Income, Expense, Capital', item };
        return {
          index: i + idx,
          valid: true,
          textToEmbed: [account_name, account_code || '', header || '', description || '', type, target_key].filter(Boolean).join(' | '),
          data: { account_code: account_code || null, account_name, header: header || null, description: description || null, type, property_type: property_type || null, portfolio, target_key }
        };
      });

      const validItems = validations.filter(v => v.valid);
      const invalidItems = validations.filter(v => !v.valid);
      invalidItems.forEach(inv => errors.push({ index: inv.index, error: inv.error, item: inv.item }));
      if (validItems.length === 0) continue;

      const embeddingPromises = validItems.map(async (v) => {
        try {
          const response = await openai.embeddings.create({ model: "text-embedding-3-large", input: v.textToEmbed, dimensions: 1536 });
          trackEmbeddingUsage(response, '/api/lineitems/batch'); // ✅ track
          return { ...v, embedding: response.data[0].embedding };
        } catch (err) {
          return { index: v.index, error: `Embedding failed: ${err.message}`, item: batch[v.index - i] };
        }
      });

      const embeddedItems = await Promise.all(embeddingPromises);
      const successes = embeddedItems.filter(item => !item.error);
      const embeddingErrors = embeddedItems.filter(item => item.error);
      embeddingErrors.forEach(err => errors.push({ index: err.index, error: err.error, item: err.item }));
      if (successes.length === 0) continue;

      const toInsert = successes.map(item => ({ id: uuidv4(), ...item.data, embedding: item.embedding, created_at: new Date().toISOString() }));

      const { data: inserted, error: insertError } = await supabase.from('lineitems').insert(toInsert).select('id, account_name, target_key, created_at').neq('id', '');
      if (insertError) throw insertError;

      inserted.forEach((row, idx) => batchResults.push({ index: successes[idx].index, status: 'success', inserted: row }));
      results.push(...batchResults);
    }

    res.status(207).json({
      summary: {
        total_received: items.length,
        total_success: results.length,
        total_failed: errors.length,
        success_rate: items.length > 0 ? (results.length / items.length * 100).toFixed(1) + '%' : '0%'
      },
      successes: results,
      failures: errors
    });

  } catch (error) {
    console.error('Batch insert failed:', error);
    res.status(500).json({ error: 'Batch insert failed', details: error.message || error.details || 'Unknown error', partial_results: results.length > 0 ? results : undefined, partial_errors: errors.length > 0 ? errors : undefined });
  }
});


// ────────────────────────────────────────────────
// POST /api/lineitems/openai-batch
// ────────────────────────────────────────────────
app.post('/api/lineitems/openai-batch', async (req, res) => {
  const items = req.body;

  if (!Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: 'Request body must be a non-empty array of line item objects' });
  }

  const validItems = [];
  const errors = [];

  items.forEach((item, idx) => {
    const { account_code, account_name, header, description, type, property_type, portfolio = false, target_key } = item;
    if (!account_name || !type || !target_key) { errors.push({ index: idx, error: 'Missing required fields: account_name, type, target_key', item }); return; }
    if (!['Income', 'Expense', 'Capital'].includes(type)) { errors.push({ index: idx, error: 'type must be one of: Income, Expense, Capital', item }); return; }
    const textToEmbed = [account_name, account_code || '', header || '', description || '', type, target_key].filter(Boolean).join(' | ');
    validItems.push({ index: idx, textToEmbed, data: { account_code: account_code || null, account_name, header: header || null, description: description || null, type, property_type: property_type || null, portfolio, target_key } });
  });

  if (validItems.length === 0) return res.status(400).json({ error: 'No valid items to process', failures: errors });

  try {
    const jsonlContent = validItems.map(v => JSON.stringify({ custom_id: String(v.index), method: 'POST', url: '/v1/embeddings', body: { model: 'text-embedding-3-large', input: v.textToEmbed, dimensions: 1536 } })).join('\n');

    const fileUpload = await openai.files.create({ file: new File([jsonlContent], 'lineitems_embeddings.jsonl', { type: 'application/json' }), purpose: 'batch' });
    const batch = await openai.batches.create({ input_file_id: fileUpload.id, endpoint: '/v1/embeddings', completion_window: '24h', metadata: { description: 'Line item embeddings batch', item_count: String(validItems.length) } });

    if (!app.locals.pendingBatches) app.locals.pendingBatches = new Map();
    app.locals.pendingBatches.set(batch.id, { validItems, submittedAt: new Date().toISOString(), totalReceived: items.length, validationErrors: errors });

    return res.status(202).json({ message: 'Batch submitted to OpenAI.', batch_id: batch.id, status: batch.status, input_file_id: fileUpload.id, total_received: items.length, valid_items: validItems.length, validation_failures: errors.length, failures: errors.length > 0 ? errors : undefined, poll_url: `/api/lineitems/openai-batch/${batch.id}` });

  } catch (error) {
    console.error('OpenAI batch submit failed:', error);
    return res.status(500).json({ error: 'Failed to submit OpenAI batch job', details: error.message || 'Unknown error' });
  }
});


// ────────────────────────────────────────────────
// GET /api/lineitems/openai-batch/:batchId
// ────────────────────────────────────────────────
app.get('/api/lineitems/openai-batch/:batchId', async (req, res) => {
  const { batchId } = req.params;
  const pending = app.locals.pendingBatches?.get(batchId);
  if (!pending) return res.status(404).json({ error: 'Batch ID not found.', batch_id: batchId });

  try {
    const batch = await openai.batches.retrieve(batchId);

    if (batch.status !== 'completed') {
      return res.status(200).json({ batch_id: batchId, status: batch.status, request_counts: batch.request_counts, created_at: batch.created_at, message: ['failed', 'expired', 'cancelled'].includes(batch.status) ? `Batch ended with status: ${batch.status}.` : 'Batch is still processing. Poll again shortly.' });
    }

    if (!batch.output_file_id) return res.status(500).json({ error: 'Batch completed but output_file_id is missing', batch_id: batchId });

    const fileResponse = await openai.files.content(batch.output_file_id);
    const rawText = await fileResponse.text();
    const outputLines = rawText.split('\n').filter(line => line.trim()).map(line => JSON.parse(line));

    const { validItems, validationErrors } = pending;
    const itemsByIndex = new Map(validItems.map(v => [String(v.index), v]));
    const toInsert = [];
    const insertErrors = [];

    for (const line of outputLines) {
      const { custom_id, response, error } = line;
      if (error || response?.status_code !== 200) {
        insertErrors.push({ index: Number(custom_id), error: error?.message || `Embedding API returned status ${response?.status_code}`, item: itemsByIndex.get(custom_id)?.data });
        continue;
      }
      const embedding = response.body?.data?.[0]?.embedding;
      if (!embedding) { insertErrors.push({ index: Number(custom_id), error: 'No embedding returned', item: itemsByIndex.get(custom_id)?.data }); continue; }

      const v = itemsByIndex.get(custom_id);
      if (!v) continue;

      // ✅ Track batch embedding tokens (batch API returns usage per line)
      const usageTokens = response.body?.usage?.total_tokens || 0;
      tokenUsage['text-embedding-3-large'].requests += 1;
      tokenUsage['text-embedding-3-large'].total_tokens += usageTokens;

      toInsert.push({ id: uuidv4(), ...v.data, embedding, created_at: new Date().toISOString() });
    }

    let inserted = [];
    if (toInsert.length > 0) {
      const SUPABASE_CHUNK = 500;
      for (let i = 0; i < toInsert.length; i += SUPABASE_CHUNK) {
        const chunk = toInsert.slice(i, i + SUPABASE_CHUNK);
        const { data: chunkInserted, error: insertError } = await supabase.from('lineitems').insert(chunk).select('id, account_name, target_key, created_at');
        if (insertError) {
          console.error('Supabase chunk insert error:', insertError);
          chunk.forEach(row => insertErrors.push({ error: `Supabase insert failed: ${insertError.message}`, item: row }));
        } else {
          inserted.push(...(chunkInserted || []));
        }
      }
    }

    app.locals.pendingBatches.delete(batchId);
    const allErrors = [...(validationErrors || []), ...insertErrors];

    return res.status(207).json({
      batch_id: batchId,
      status: 'completed',
      summary: { total_received: pending.totalReceived, valid_submitted: validItems.length, total_success: inserted.length, total_failed: allErrors.length, success_rate: pending.totalReceived > 0 ? (inserted.length / pending.totalReceived * 100).toFixed(1) + '%' : '0%' },
      successes: inserted,
      failures: allErrors.length > 0 ? allErrors : undefined
    });

  } catch (error) {
    console.error('Batch poll/insert failed:', error);
    return res.status(500).json({ error: 'Failed to retrieve or process batch results', details: error.message || 'Unknown error', batch_id: batchId });
  }
});


// ────────────────────────────────────────────────
// DEBUG ENDPOINTS
// ────────────────────────────────────────────────
app.get('/api/debug/embeddings', async (req, res) => {
  const { data, error } = await supabase.from('lineitems').select('id, account_name, target_key, embedding').limit(5);
  if (error) return res.status(500).json({ error: error.message });
  res.json({
    total_sampled: data.length,
    analysis: data.map(item => ({
      id: item.id,
      account_name: item.account_name,
      target_key: item.target_key,
      has_embedding: !!item.embedding,
      embedding_length: Array.isArray(item.embedding) ? item.embedding.length : 0,
      embedding_sample: Array.isArray(item.embedding) ? item.embedding.slice(0, 3) : null
    }))
  });
});

app.get('/api/debug/vector-test', async (req, res) => {
  const results = {};

  try {
    const { data: sample, error: sampleError } = await supabase.from('lineitems').select('id, account_name, type, target_key, embedding').limit(3);
    results.sample_data = {
      count: sample?.length || 0,
      error: sampleError?.message || null,
      samples: (sample || []).map(s => {
        const emb = s.embedding;
        const isArray = Array.isArray(emb);
        const isString = typeof emb === 'string';
        return { id: s.id, account_name: s.account_name, type: s.type, target_key: s.target_key, embedding_type: typeof emb, embedding_is_array: isArray, embedding_is_string: isString, embedding_length: isArray ? emb.length : isString ? emb.length : 'N/A', embedding_sample: isArray ? emb.slice(0, 4) : isString ? emb.substring(0, 60) + '...' : null };
      }),
    };

    const { data: allTypes, error: typeError } = await supabase.from('lineitems').select('type');
    const typeFrequency = {};
    (allTypes || []).forEach(row => { const key = `"${row.type}"`; typeFrequency[key] = (typeFrequency[key] || 0) + 1; });
    results.type_frequency = { error: typeError?.message || null, counts: typeFrequency };

    const { data: realRow, error: realRowError } = await supabase.from('lineitems').select('id, account_name, embedding').not('embedding', 'is', null).limit(1).single();
    results.real_embedding_source = { error: realRowError?.message || null, id: realRow?.id || null, account_name: realRow?.account_name || null, has_embedding: !!realRow?.embedding };

    const dummyArray = new Array(1536).fill(0.01);
    const realEmbedding = realRow?.embedding ?? null;

    const rpcConfigs = [
      { label: 'real_embedding_as_array', query_embedding: Array.isArray(realEmbedding) ? realEmbedding : null, filter_type: 'Income', skip: !realEmbedding || !Array.isArray(realEmbedding), skip_reason: 'Real embedding not available or not an array' },
      { label: 'real_embedding_as_string', query_embedding: realEmbedding ? (Array.isArray(realEmbedding) ? `[${realEmbedding.join(',')}]` : realEmbedding) : null, filter_type: 'Income', skip: !realEmbedding, skip_reason: 'Real embedding not available' },
      { label: 'dummy_array_income', query_embedding: dummyArray, filter_type: 'Income', skip: false },
      { label: 'dummy_string_income', query_embedding: `[${dummyArray.join(',')}]`, filter_type: 'Income', skip: false },
      { label: 'dummy_array_expense', query_embedding: dummyArray, filter_type: 'Expense', skip: false },
    ];

    results.rpc_tests = {};
    for (const cfg of rpcConfigs) {
      if (cfg.skip) { results.rpc_tests[cfg.label] = { skipped: true, reason: cfg.skip_reason }; continue; }
      const { data, error } = await supabase.rpc('match_lineitems_raw', { p_query_embedding: cfg.query_embedding, p_match_threshold: 0.3, p_match_count: 3, p_filter_type: cfg.filter_type });
      results.rpc_tests[cfg.label] = { filter_type: cfg.filter_type, embedding_format: typeof cfg.query_embedding === 'string' ? 'string' : 'array', match_count: data?.length ?? 0, matches: (data || []).map(m => ({ account_name: m.account_name, target_key: m.target_key, similarity: m.similarity })), error: error?.message || null, error_code: error?.code || null, error_hint: error?.hint || null };
    }

    const workingFormats = Object.entries(results.rpc_tests).filter(([, v]) => !v.skipped && !v.error && v.match_count > 0).map(([label]) => label);
    results.verdict = { working_formats: workingFormats, recommendation: workingFormats.length === 0 ? 'No RPC calls returned results.' : `Use the "${workingFormats[0]}" format in your main mapping route.` };

    res.json(results);
  } catch (err) {
    res.status(500).json({ error: err.message, partial_results: results });
  }
});

// ────────────────────────────────────────────────
// POST /api/analyze-rent-roll
// Uses gemini-2.5-flash via gateway with SSE streaming
// and usage logging to MySQL
// ────────────────────────────────────────────────
const RENT_ROLL_MODEL = 'gpt-4o-mini';

const RENT_ROLL_SYSTEM_PROMPT = `You are a rent roll spreadsheet parser. Given an anonymized HTML table of an Excel rent roll, analyze the structure and output a JSON parsing instruction.

[THINKING]
Briefly identify: header row(s), data start row, the suite/unit/space ID column (the most critical anchor), and the overall layout.

[GROUPING]
Briefly describe: how new tenants start, what continuation rows look like (rows with no suite ID that belong to the same tenant), what rows to skip. Note any column groups visible from merged/parent headers.

[PARSING INSTRUCTION]
\`\`\`json
{
  "header_rows": [],
  "data_starts_at_row": null,
  "suite_id_col": "",
  "tenant_name_col": "",
  "scalar_fields": [
    { "field_id": "lease_start", "col": "C", "label": "Lease Commencement" },
    { "field_id": "lease_end",   "col": "D", "label": "Lease Expiration" },
    { "field_id": "gla_sqft",   "col": "E", "label": "NRA (SF)" }
  ],
  "groups": [
    {
      "id": "base_rent",
      "label": "Base Rent",
      "collection": false,
      "columns": [
        { "col": "F", "label": "Monthly" },
        { "col": "G", "label": "PSF" }
      ]
    },
    {
      "id": "current_charges",
      "label": "Current Charges",
      "collection": true,
      "columns": [
        { "col": "H", "label": "Code" },
        { "col": "I", "label": "Amount" },
        { "col": "J", "label": "PSF" }
      ]
    }
  ],
  "skip_row_patterns": [],
  "new_tenant_rule": "suite_id column non-empty",
  "confidence": "high",
  "notes": ""
}
\`\`\`

Rules:
- suite_id_col is REQUIRED — identify the column with the unique tenant space identifier (suite #, unit #, space #, bay #, parcel #, etc.). This is the most important field.
- scalar_fields: every column that has exactly one value per tenant. Use descriptive field_ids. Common known slugs: lease_start, lease_end, gla_sqft, sign_date, lease_term, occupancy_date, option_date, occupancy_status, floor, building, property_name. For unrecognized fields, create a descriptive snake_case id (e.g. "cam_start", "renewal_option").
- groups: columns that naturally belong together, discovered from the actual header structure — do NOT use a fixed/predefined list of groups.
  - Look for parent headers that span multiple sub-columns (e.g. a merged "Current Charges" cell above "Code", "Amount", "PSF"). Each such parent is a group.
  - collection: true when the group's columns appear on multiple continuation rows for one tenant (e.g. multiple charge lines, multiple rent escalation steps).
  - collection: false when the group has only one row of data per tenant (e.g. "Base Rent" with Monthly + PSF sub-columns side by side).
  - Discover as many groups as exist in the data — there is no limit.
- skip_row_patterns: regex patterns to skip total/subtotal/blank separator rows (e.g. "(?i)\\btotal\\b", "^\\s*$").
- Only include columns that actually exist in the data. Use "" for suite_id_col/tenant_name_col if not confidently identified.

[FLAGS]
List genuine ambiguities only. If none, say "None."

IMPORTANT: Keep text minimal — the JSON is what matters.`;

function estimateGeminiCost(model, promptTokens, completionTokens) {
  const rates = {
    'gpt-4o-mini':      { input: 0.15, output: 0.60 },
    'gemini-2.5-flash': { input: 0.15, output: 0.60 },
    'gemini-2.5-pro':   { input: 1.25, output: 10.00 },
  };
  const r = rates[model] || { input: 0.15, output: 0.60 };
  return (promptTokens / 1_000_000) * r.input + (completionTokens / 1_000_000) * r.output;
}

function logRentRollUsage({ provider, model, promptTokens, completionTokens, latencyMs, statusCode, isStream, error }) {
  pool.execute(
    `INSERT INTO usage_logs
      (service_key, service_name, provider, model,
       prompt_tokens, completion_tokens, total_tokens,
       estimated_cost_usd, latency_ms, status_code, is_stream, error)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [
      'analyze-rent-roll',
      'Analyze Rent Roll',
      provider,
      model,
      promptTokens,
      completionTokens,
      promptTokens + completionTokens,
      estimateGeminiCost(model, promptTokens, completionTokens),
      latencyMs,
      statusCode,
      isStream ? 1 : 0,
      error ?? null,
    ]
  ).catch(err => console.error("[rent-roll] failed to log usage:", err.message));
}

function extractStreamUsage(sseText) {
  const lines = sseText.split("\n").reverse();
  for (const line of lines) {
    if (!line.startsWith("data:")) continue;
    const data = line.slice(5).trim();
    if (data === "[DONE]") continue;
    try {
      const parsed = JSON.parse(data);
      if (parsed.usage?.prompt_tokens != null) return parsed.usage;
    } catch { /* skip */ }
  }
  return { prompt_tokens: 0, completion_tokens: 0 };
}

app.post('/api/analyze-rent-roll', async (req, res) => {
  const startedAt = Date.now();
  const { sampleHtml, contextNote = '' } = req.body;

  if (!sampleHtml) {
    return res.status(400).json({ error: 'sampleHtml is required' });
  }

  try {
    const upstream = await fetch('https://ai.gateway.bulkscraper.cloud/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${process.env.GATEWAY_API_KEY}`,
      },
      body: JSON.stringify({
        model: RENT_ROLL_MODEL,
        stream: true,
        messages: [
          { role: 'system', content: RENT_ROLL_SYSTEM_PROMPT },
          { role: 'user',   content: contextNote ? `${contextNote}\n\n${sampleHtml}` : sampleHtml },
        ],
      }),
    });

    if (!upstream.ok) {
      const text = await upstream.text();
      let parsed;
      try { parsed = JSON.parse(text); } catch { parsed = { message: text }; }
      const errMsg = parsed?.error?.message ?? parsed?.message ?? 'Upstream error';

      // logRentRollUsage({
      //   provider: 'gateway', model: RENT_ROLL_MODEL,
      //   promptTokens: 0, completionTokens: 0,
      //   latencyMs: Date.now() - startedAt,
      //   statusCode: upstream.status,
      //   isStream: true, error: errMsg,
      // });

      if (upstream.status === 429) return res.status(429).json({ error: 'Rate limit exceeded' });
      if (upstream.status === 402) return res.status(402).json({ error: 'Payment required' });
      return res.status(upstream.status).json({ error: errMsg });
    }

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('X-Gateway-Provider', 'gateway');
    res.flushHeaders();

    const reader  = upstream.body.getReader();
    const decoder = new TextDecoder();
    const chunks  = [];

    req.on('close', () => reader.cancel());

    while (true) {
      const { done, value } = await reader.read();
      if (done) {
        const usage = extractStreamUsage(chunks.join(''));
        // logRentRollUsage({
        //   provider: 'gateway', model: RENT_ROLL_MODEL,
        //   promptTokens: usage.prompt_tokens,
        //   completionTokens: usage.completion_tokens,
        //   latencyMs: Date.now() - startedAt,
        //   statusCode: 200, isStream: true,
        // });
        res.end();
        break;
      }
      const text = decoder.decode(value, { stream: true });
      chunks.push(text);
      res.write(text);
    }
  } catch (error) {
    console.error('analyze-rent-roll failed:', error);
    if (res.headersSent) {
      res.write(`data: ${JSON.stringify({ error: error.message || 'Upstream error' })}\n\n`);
      res.write('data: [DONE]\n\n');
      res.end();
    } else {
      res.status(500).json({ error: error.message || 'Failed to analyze rent roll' });
    }
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});