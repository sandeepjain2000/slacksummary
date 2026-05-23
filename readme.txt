Purpose
-------
Slack reporting utilities in Python: improved Slack test script (zslack_test_improved.py), Excel mapping (SlackIds.xlsx), config.json / email_config.json, and archived Reports.

How to use
----------
1. Install Python dependencies implied by zslack_test_improved.py (requests, pandas/openpyxl if used, slack SDK, etc.).
2. Fill config.json with bot token, workspace IDs, and channels; keep tokens out of git.
3. Run python zslack_test_improved.py or use s.bat for a saved command line on Windows.
4. Store generated summaries under Reports/; use archive/ and old reports/ for dated history.
