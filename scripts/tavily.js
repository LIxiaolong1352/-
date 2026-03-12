#!/usr/bin/env node
// Tavily Search CLI
const API_KEY = "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";
const API_URL = "https://api.tavily.com/search";

async function search(query, options = {}) {
  const response = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      api_key: API_KEY,
      query: query,
      search_depth: options.depth || "basic",
      max_results: options.max || 5,
      include_answer: options.answer !== false
    })
  });
  return await response.json();
}

// CLI
const query = process.argv.slice(2).join(" ");
if (!query) {
  console.log("Usage: node tavily.js <query>");
  process.exit(1);
}

search(query).then(data => {
  if (data.answer) console.log("Answer:", data.answer, "\n");
  (data.results || []).forEach((r, i) => {
    console.log(`${i+1}. ${r.title}\n   ${r.url}\n   ${r.content?.slice(0, 150)}...\n`);
  });
}).catch(e => console.error("Error:", e.message));
