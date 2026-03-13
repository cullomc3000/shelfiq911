ShelfIQ 911 - McKinsey Dashboard Bundle

Included:
- streamlit_app.py
- requirements.txt

Use:
1. Upload both files to your GitHub repo.
2. Make sure the repo uses requirements.txt exactly.
3. Deploy streamlit_app.py in Streamlit Cloud.

This version includes:
- McKinsey / PowerBI-style executive dashboard
- Sidebar navigation
- KPI cards
- Executive narrative
- Strategic Action Center
- Demo data and client upload mode

---

## Features

• Executive dashboard  
• AI insights engine  
• Distribution gap detection  
• Store performance scoring  
• Momentum analysis (13W vs 52W)  
• Sell-in recommendations  
• McKinsey-style executive narrative  

---

## Required Client Data

Upload Excel file with sheets:

### Sales_History
| store_id | sku_id | week_end_date | units | sales_dollars |

### Products
| sku_id | brand | category |

### Stores
| store_id | retailer | state | format |

### Shelf_Snapshot (optional)
| store_id | sku_id | facings | shelf_share |

---

## Run Locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
