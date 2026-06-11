**The editing surface — open this when you want to change content:**
platform_content.xlsx

The workbook has a `README` sheet and 8 entity sheets (`data_sources`, `foundation_features`, `value_categories`, `use_cases`, `use_case_features`, `personas`, `use_case_personas`, `maturity_levels`). Pre-filled with 6 sources, 13 features, 6 value categories, 12 use cases, 6 personas, and a maturity ladder for the most-used cases — so the visualization is fully populated out of the box and you can see the shape immediately.

**The viewing surface — open this to see the infographic:**
platform_infographic.html

Self-contained — no server needed. Click any use case to highlight the foundation features it relies on, dim the irrelevant ones, light up the personas concerned, and slide open a right-hand drawer with the value calculation, customer story, personas, features used, and the maturity ladder. The top filter buttons let you focus on a single value category at a time, and curved gradient lines visually link sources → hub → categories → use cases (and selected use case → its features).

**The pipeline — your one command to refresh after editing:**
xlsx_to_json.py and data.json

```
python3 xlsx_to_json.py
```

That single command reads the workbook, validates referential integrity (every `value_category_id`, `feature_id`, `persona_id`, `use_case_id` reference must resolve — you get a clear error and a non-zero exit if anything dangles), writes `data.json`, *and* injects the JSON inline into the HTML so the file works straight from disk with no server. Just refresh the browser tab.

Your day-to-day loop is now: edit a row in the spreadsheet → run one command → reload. When you have your real taxonomy of data sources, foundation features, value categories and use cases, paste them into the corresponding sheets and you're done — no code changes needed.