# Campaign+Asset Multi-Tab Output (Excel Only, Figma S3 AmzId Columns)

This Streamlit app generates multi-campaign output sheets with product images and Figma S3 links from an Excel campaign input and a product CSV.

## Features

- **Excel upload:** Upload your campaign Excel file.
- **Product CSV upload:** CSV must have `MB_id` and `image_src` columns.
- **Output:** 
  - Separate tabs for each unique Campaign Name + Asset combination.
  - Each tab contains: Hub, Focus Grid, PID1, PID2, Img1, Img2, AmzId1, AmzId2.
  - Img1/Img2: Milkbasket product links.
  - AmzId1/AmzId2: Figma S3 image links (with `.png` extension).
  - All_PIDs tab: All unique PIDs, Milkbasket link, and Figma S3 link.
- **Image tools:** Download all images or remove backgrounds.

## How to Run

```bash
pip install -r requirements.txt
streamlit run main.py
