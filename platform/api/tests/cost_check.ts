/**
 * Check Anthropic API cost for parsing one WO PDF.
 * Usage: npx tsx tests/cost_check.ts
 */
import 'dotenv/config';
import { config } from '../src/config.js';
import { readFileSync } from 'fs';

const OPUS_INPUT_PER_M = 15;
const OPUS_OUTPUT_PER_M = 75;

async function run() {
  const pdfPath = '/Users/stamatiangelides/Desktop/Oneiro/RM-41594.pdf';
  const fileBytes = readFileSync(pdfPath);
  const encoded = fileBytes.toString('base64');

  console.log(`PDF: ${pdfPath}`);
  console.log(`Size: ${(fileBytes.length / 1024).toFixed(0)} KB\n`);

  // Minimal request to measure input token cost of the PDF
  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.anthropic.apiKey,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: 'claude-opus-4-6',
      max_tokens: 16,
      messages: [{
        role: 'user',
        content: [
          { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: encoded } },
          { type: 'text', text: 'Reply with just: ok' },
        ],
      }],
    }),
  });

  const result = await response.json() as any;
  const usage = result.usage;

  if (!usage) {
    console.error('No usage data:', JSON.stringify(result, null, 2));
    return;
  }

  // The real extraction prompt adds ~3K tokens to input, and output is ~800-1500 tokens
  const promptTokens = 3000;
  const realOutputTokens = 1200;

  const pdfInputTokens = usage.input_tokens;
  const totalInputTokens = pdfInputTokens + promptTokens;

  const inputCost = (totalInputTokens / 1_000_000) * OPUS_INPUT_PER_M;
  const outputCost = (realOutputTokens / 1_000_000) * OPUS_OUTPUT_PER_M;
  const totalCost = inputCost + outputCost;

  console.log('=== TOKEN USAGE (from API) ===');
  console.log(`  PDF input tokens:     ${pdfInputTokens.toLocaleString()}`);
  console.log(`  + Extraction prompt:  ~${promptTokens.toLocaleString()}`);
  console.log(`  = Total input:        ~${totalInputTokens.toLocaleString()}`);
  console.log(`  Output (typical):     ~${realOutputTokens.toLocaleString()}`);
  if (usage.cache_read_input_tokens) {
    console.log(`  Cache read tokens:    ${usage.cache_read_input_tokens.toLocaleString()}`);
  }
  console.log('');
  console.log('=== COST PER SINGLE WO SCAN (claude-opus-4-6) ===');
  console.log(`  Input:   $${inputCost.toFixed(4)}  ($${OPUS_INPUT_PER_M}/M tokens)`);
  console.log(`  Output:  $${outputCost.toFixed(4)}  ($${OPUS_OUTPUT_PER_M}/M tokens)`);
  console.log(`  TOTAL:   $${totalCost.toFixed(4)}`);
  console.log('');
  console.log('=== VOLUME ESTIMATES ===');
  console.log(`  10 WOs/day:    $${(totalCost * 10).toFixed(2)}/day  = $${(totalCost * 10 * 22).toFixed(2)}/month`);
  console.log(`  50 WOs/day:    $${(totalCost * 50).toFixed(2)}/day  = $${(totalCost * 50 * 22).toFixed(2)}/month`);
  console.log(`  100 WOs/day:   $${(totalCost * 100).toFixed(2)}/day = $${(totalCost * 100 * 22).toFixed(2)}/month`);
  console.log('');

  // Multi-WO stack estimate (adds detect pass)
  const detectOutputTokens = 300;
  const detectCost = (totalInputTokens / 1_000_000) * OPUS_INPUT_PER_M + (detectOutputTokens / 1_000_000) * OPUS_OUTPUT_PER_M;
  console.log('=== MULTI-WO STACK (5 WOs in one PDF) ===');
  console.log(`  Pass 1 (detect):   $${detectCost.toFixed(4)}`);
  console.log(`  Pass 2 (5 parses): $${(totalCost * 5).toFixed(4)}`);
  console.log(`  TOTAL:             $${(detectCost + totalCost * 5).toFixed(4)}`);
  console.log(`  Per WO:            $${((detectCost + totalCost * 5) / 5).toFixed(4)}`);
}

run().catch(console.error);
