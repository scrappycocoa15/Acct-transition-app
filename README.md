# CSE Assignment Tool — Streamlit

Automates RSE → CSE account transitions by fetching live Salesforce data, applying the balancing algorithm, and exporting results as Excel.

---

## Option A — Deploy to Streamlit Community Cloud (no local hosting)

1. Create a free account at [streamlit.io](https://streamlit.io)
2. Push `streamlit_app.py` and `requirements.txt` to a **GitHub repo** (can be private)
3. In Streamlit Cloud, click **New app** → connect the repo → set main file to `streamlit_app.py`
4. Click **Deploy** — your app will be live at a `*.streamlit.app` URL

> Streamlit Community Cloud is free for public and private repos.

---

## Option B — Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

App opens at `http://localhost:8501`.

---

## How to get your Salesforce Session ID

1. Log into Salesforce in Chrome or Edge
2. Open **DevTools** (F12) → **Application** → **Cookies** → `https://sapconcur.my.salesforce.com`
3. Copy the value of the **`sid`** cookie
4. Paste it into the app and click **Fetch & Assign**

Session IDs expire after ~2 hours of inactivity or on logout.

---

## What the app does

| Step | Action |
|------|--------|
| 1 | Fetches **New Accounts Transitioning** report (`00O0e000005i4gg`) via Salesforce Async Analytics API |
| 2 | Fetches **Current Account Volumes** report (`00O7V000006IT4i`) |
| 3 | Builds active CSE pool per segment — excludes anyone with < 20 accounts |
| 4 | Runs the RSE → CSE assignment algorithm: volume balance → OB CSM partnership → ARR parity |
| 5 | Applies monthly overrides (Tom Wahl / Alex Capeloto) — toggleable in the sidebar |
| 6 | Displays results with highlighted new-owner columns |
| 7 | Exports `cse_assignments.xlsx` with three sheets: **Assignments**, **Exceptions**, **CSE Summary** |

---

## Updating monthly overrides

Toggle the checkboxes in the **Settings** sidebar before clicking Fetch & Assign. No code changes needed.

To permanently change the defaults, edit `MONTHLY_OVERRIDES` at the top of `streamlit_app.py`:

```python
active_overrides = {
    "Strategic": {"Thomas Wahl": 1},
    "Premier":   {"Alex Capeloto": 1},
}
```

---

## Updating the CSE roster or partnerships

The CSE roster and OB CSM partnerships are embedded in `streamlit_app.py`. Search for `_RAW_CSES` (roster) and `_STRAT_P` / `_PREMIER_P` / `_KEY_P` (partnerships) to make changes.
