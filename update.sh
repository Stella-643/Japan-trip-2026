#!/usr/bin/env bash
# 用法：改完 Japan_Trip_D1-D9_Template.xlsx 后执行
#   ./update.sh "改了D3午餐"
# 作用：重新生成 index.html → 提交 → 推送（Cloudflare Pages 会自动重新部署）
set -e
cd "$(dirname "$0")"

echo "① 重新生成页面…"
python3 build_travel.py

echo "② 提交改动…"
git add -A
git commit -m "${1:-update itinerary}" || echo "  （没有新改动可提交）"

echo "③ 推送到 GitHub…"
git push

echo ""
echo "✅ 完成。Cloudflare 会在约 30 秒内自动重新部署，稍后刷新你的站点即可。"
