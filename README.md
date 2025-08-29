# Orbico Invoice → Excel (Streamlit)

Upload an Orbico PDF invoice to extract line items and compute **Stvarna količina (kom)**.

## Rules
- `AxB(Unit)` → use **B** (e.g., `4X4L` → 4 L per bottle; `12X0,4KG` → 0.4 KG)
- Else use the **last** standalone number+unit (e.g., `... 55L` → 55 L)
- Supports **L** and **KG**; avoids false matches like `R4 L`.

## Run locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
