#!/usr/bin/env python3
"""简化版 summarize - 用 Tavily API 总结网页内容"""
import sys
import json
import urllib.request
import urllib.parse

API_KEY = "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"

def summarize_url(url):
    """用 Tavily 提取并总结网页"""
    data = json.dumps({
        "api_key": API_KEY,
        "query": f"总结这个网页的主要内容: {url}",
        "search_depth": "advanced",
        "max_results": 3,
        "include_answer": True
    }).encode()
    
    req = urllib.request.Request(
        "https://api.tavily.com/search",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    
    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode())
        
    if result.get("answer"):
        print("\n[总结]\n" + result["answer"])
    
    print("\n[来源]")
    for r in result.get("results", [])[:3]:
        print(f"  - {r.get('title', 'N/A')}")
        print(f"    {r.get('url', 'N/A')}")
        if r.get('content'):
            print(f"    {r['content'][:200]}...\n")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python summarize.py <网址>")
        sys.exit(1)
    
    url = sys.argv[1]
    if not url.startswith(("http://", "https://")):
        print("请输入完整网址 (http:// 或 https://)")
        sys.exit(1)
    
    summarize_url(url)
