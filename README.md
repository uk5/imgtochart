# Image to Excel Converter 📊

A Streamlit application that uses Google's Gemini AI to extract data from images of charts/tables and convert them into an Excel file with native charts (Bar, Line, Pie, Doughnut, etc.) and matching colors!

## Features
- **AI-Powered Extraction**: Uses Gemini Pro Vision / Flash to understand images.
- **Auto-Chart Detection**: Identifies whether it's a Bar, Pie, Doughnut, etc.
- **Native Excel Charts**: Generates a real Excel chart, not just a static image.
- **Color Matching**: Extracts colors from the image and applies them to the Excel chart.

## Setup Locally
1. Clone the repo.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the app:
   ```bash
   streamlit run app.py
   ```

## Deployment on Streamlit Cloud
1. Push this code to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io/).
3. Connect your GitHub and deploy this repo.
4. **Important**: set your `GOOGLE_API_KEY` in the Streamlit "Secrets" settings if you want to hardcode it, or just enter it in the UI.
