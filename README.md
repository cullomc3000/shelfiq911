# ShelfIQ 911 - Power BI Style Streamlit App

This package is set up for local use, GitHub, and Streamlit Community Cloud deployment.

## Included files

- `streamlit_app.py` -> main app file
- `requirements.txt` -> Python dependencies
- `assets/` -> static brand files like logos, icons, background images, and design assets
- `logo/` -> place your company logo here if you want a default logo in the repo
- `sample_data/` -> optional demo input files for testing uploads

## What is the `assets` folder?

Use `assets/` for files that support the app visually but are not user-uploaded at runtime. Examples:
- company logos
- app icons
- banner images
- background images
- PDF cover graphics
- branded reference images

Example paths:
- `assets/company_logo.png`
- `assets/dashboard_banner.jpg`

## Recommended repo structure

```text
streamlit_powerbi_repo/
├── streamlit_app.py
├── requirements.txt
├── README.md
├── assets/
│   └── put_static_files_here.txt
├── logo/
│   └── add_default_logo_here.txt
└── sample_data/
    └── add_sample_files_here.txt
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy to Streamlit Community Cloud

1. Upload this folder to GitHub
2. In Streamlit Community Cloud, connect your GitHub repo
3. Set the main file path to `streamlit_app.py`
4. Deploy

## Notes

- Your app already supports uploading a logo in the sidebar at runtime.
- The `logo/` folder is included so you can also keep a default branded logo in the repository.
- The `sample_data/` folder is optional but useful for demos and testing.
