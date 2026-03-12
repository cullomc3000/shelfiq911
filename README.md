# ShelfIQ 911

AI-powered retail analytics platform for:

- Distribution Gap Analysis
- SKU Velocity Scoring
- Momentum Detection
- Sell-In Recommendation Engine
- Shelf Productivity Analysis

Built using **Python + Streamlit**.

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
