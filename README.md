# CSE Account Assignment Tool

## Quickstart

### 1. Export reports from Salesforce (recommended workflow)

For each of the two reports below:
1. Open the report in Salesforce
2. Click **Export** (top-right corner)
3. Select **Details Only**
4. Click **Export** → choose **Formatted Report: CSV**

| Report | Salesforce ID |
|--------|--------------|
| New Accounts Transitioning | `00O0e000005i4gg` |
| Current Account Volumes | `00O7V000006IT4i` |

### 2. Run the app locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

### 3. Upload and run

- The app opens in **Upload CSV Files** mode by default
- Upload both CSV exports
- Click **Run Assignment**
- Download the Excel output

---

## Deploy to Streamlit Community Cloud (free)

1. Push this folder to a GitHub repo
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Point to your repo and `streamlit_app.py`
4. Deploy — no secrets or environment variables needed

---

## Live Salesforce Fetch (optional)

Switch to **Live Salesforce Fetch** mode and paste your `sid` cookie value.  
To get the `sid` cookie: DevTools → Application → Cookies → `sapconcur.my.salesforce.com` → copy `sid` value.

> **Note:** Live fetch may time out when the app is hosted on Streamlit Cloud because  
> Salesforce locks session tokens to the originating browser IP. CSV upload is more reliable.

---

## Output: cse_assignments.xlsx

| Sheet | Contents |
|-------|----------|
| Assignments | All assigned accounts — new-owner columns highlighted blue, override rows highlighted yellow |
| Exceptions | Accounts that could not be assigned (missing CS Team, no eligible CSE, etc.) |
| CSE Summary | Pre/post account counts and ARR per rep — rows receiving accounts highlighted green |
