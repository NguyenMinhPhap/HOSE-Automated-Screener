name: HOSE Daily Screener and Email

on:
  schedule:
    # 12:15 AM giờ Việt Nam (GMT+7) tương ứng 05:15 UTC hàng ngày
    - cron: '15 5 * * *'
  workflow_dispatch:

jobs:
  run-pipeline:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout code
        uses: actions/checkout@v3

      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.10'

      - name: Install dependencies
        run: |
          pip install pandas openpyxl numpy

      - name: Run Master Screener & Send Email
        env:
          MY_EMAIL: ${{ secrets.nguyenminhphap8322@gmail.com }}
          MAIL_PASSWORD: ${{ secrets.escn nkau jwyx wssc }}
        run: |
          python master_screener.py
