# .github/workflows/daily_report.yml
# ★ 重點：cache/ 目錄必須 commit 回 repo，否則每次 Actions 都是全新環境，
#   快照凍結機制會失效（每天都變成「第一次執行」）。
name: Daily TW Strategy Report

on:
  schedule:
    # GitHub Actions 用 UTC；台灣 08:00 = UTC 00:00
    - cron: "0 0 * * 1-5"
  workflow_dispatch:
    inputs:
      rebuild_cache:
        description: "清空快取重建（策略邏輯變更後才需要）"
        required: false
        default: "0"

permissions:
  contents: write          # ★ 必要：允許把 cache/ 回寫 repo

jobs:
  report:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install deps
        run: |
          python -m pip install --upgrade pip
          pip install yfinance python-pptx pandas numpy matplotlib requests holidays openpyxl

      - name: Run report & send mail
        env:
          GMAIL_USER:    ${{ secrets.GMAIL_USER }}
          GMAIL_PASS:    ${{ secrets.GMAIL_PASS }}
          TO_EMAIL:      ${{ secrets.TO_EMAIL }}
          REBUILD_CACHE: ${{ github.event.inputs.rebuild_cache }}
        run: python run_and_send.py

      # ★ 關鍵步驟：把更新後的快照 commit 回 repo
      # if: always() → 即使寄信失敗，當天算出來的快照也要保存下來
      - name: Commit cache snapshots
        if: always()
        run: |
          git config user.name  "github-actions[bot]"
          git config user.email "41898282+github-actions[bot]@users.noreply.github.com"
          git add -f cache/ 2>/dev/null || true
          if git diff --staged --quiet; then
            echo "快照無變更，略過 commit"
          else
            git commit -m "chore: update signal/market cache $(date -u +%Y-%m-%d)"
            # 併發保護：若期間有其他 commit，先 rebase 再 push
            git pull --rebase --autostash origin ${{ github.ref_name }} || true
            git push
          fi

      - name: Upload artifacts
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: reports-${{ github.run_id }}
          path: reports/
          retention-days: 30
          if-no-files-found: ignore
