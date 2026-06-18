---
title: M-Pesa Contribution Tracker
emoji: 💰
colorFrom: blue
colorTo: green
sdk: streamlit
sdk_version: 1.30.0
app_file: app.py
pinned: false
---

# M-Pesa Contribution Tracker Pro

Multi-Church Sunday Contribution Ledger — Track M-Pesa and cash contributions with **AI-powered handwriting recognition** (EasyOCR).

## Features
- 📄 Parse M-Pesa PDF statements automatically
- 📸 Read handwritten cash contribution notes using EasyOCR
- 📊 Generate formatted Excel reports (Summary, Bible Talk, Combined, Income & Expenses)
- 🏛️ Multi-church support with Supabase auth
- 👤 Member matching with aliases

## Live Demo
[![Hugging Face Spaces](https://img.shields.io/badge/🤗%20Hugging%20Face-Spaces-blue)](https://huggingface.co/spaces/YOUR_USERNAME/mpesa-contribution-tracker)

## Local Development
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deployment to Hugging Face Spaces

### Prerequisites
1. A [Hugging Face](https://huggingface.co/) account (free)
2. Your Supabase project credentials
3. This repository pushed to GitHub

### Step-by-Step

1. **Go to** https://huggingface.co/new-space
2. **Configure:**
   - Space Name: `mpesa-contribution-tracker`
   - License: `MIT`
   - SDK: **Streamlit**
   - Hardware: **CPU (free)** — it's enough for EasyOCR
3. **Connect your GitHub repo** or upload files manually
4. **Add Secrets** (in Space Settings → Repository Secrets):
   ```
   SUPABASE_URL = https://wcrcfitxlkymhcodqdfb.supabase.co
   SUPABASE_KEY = eyJhbGciOiJIUzI1NiIs...
   SUPABASE_SERVICE_ROLE_KEY = eyJhbGciOiJIUzI1NiIs...
   ```
5. **Set `streamlit run app.py`** as the entry command
6. **Deploy** — first run will download EasyOCR model (~200MB), subsequent runs are instant

### Important Notes
- Hugging Face Spaces free tier gives **16GB RAM + 2 vCPU** — plenty for EasyOCR + PyTorch
- The EasyOCR model downloads automatically on first run, then stays cached
- Keep `easyocr>=1.7` in `requirements.txt`
- Do **not** include `packets.txt` with tesseract (not needed for EasyOCR)