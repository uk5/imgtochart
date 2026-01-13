import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
from PIL import Image
import json
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList

def get_vision_model():
    """Finds the best available vision-capable model."""
    try:
        models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        
        # Priority list
        priorities = ['gemini-2.5-flash', 'gemini-1.5-pro', 'gemini-1.5-flash']
        
        for p in priorities:
            for m in models:
                if p in m:
                    return m
        
        # Fallback
        for m in models:
            if 'gemini' in m:
                return m
                
        return 'gemini-1.5-flash'
    except Exception as e:
        return 'gemini-1.5-flash'

def process_image(image, api_key):
    genai.configure(api_key=api_key)
    
    # Use specific model or auto-detect
    # using known working model for now to ensure stability
    model_name = 'models/gemini-2.5-flash' 
    
    try:
        model = genai.GenerativeModel(model_name)
    except:
        model_name = get_vision_model()
        model = genai.GenerativeModel(model_name)

    st.write(f"Using AI Model: `{model_name}`")

    prompt = """
    Analyze this image. It contains a chart or table. 
    1. Identify the type of chart. Choose strictly from: 
       ['bar', 'column', 'line', 'pie', 'doughnut', 'scatter', 'area', 'table']
       (Note: A 'doughnut' is a pie chart with a hole in the center).
    2. Extract the data efficiently and accurately.
    3. **Extract the colors** used for each category or series in the chart. Return them as Hex codes (e.g., #FF0000).
    
    Output the result ONLY as a VALID JSON object with the following structure:
    {
        "chart_type": "detected_type",
        "csv_data": "raw_csv_string",
        "colors": ["#HexCode1", "#HexCode2", ...] 
    }
    
    NOTE: The order of colors in the 'colors' list MUST match the order of rows in the csv_data (excluding the header).
    
    Do not include markdown code blocks (like ```json). 
    Do not include any introductory text. 
    """
    
    with st.spinner('Analyzing image and extracting data...'):
        response = model.generate_content([prompt, image])
        text_response = response.text
        
        # Clean response
        text_response = text_response.replace('```json', '').replace('```', '').strip()
        
        try:
            result = json.loads(text_response)
            chart_type = result.get("chart_type", "bar").lower()
            csv_string = result.get("csv_data", "")
            colors = result.get("colors", [])
        except json.JSONDecodeError:
            st.warning("JSON parsing failed. Attempting raw CSV retrieval.")
            csv_string = text_response
            chart_type = "bar"
            colors = []
            
    return chart_type, csv_string, colors

def generate_excel(df, chart_type, colors):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Chart Logic
        chart = None
        if "doughnut" in chart_type:
            chart = DoughnutChart()
            chart.style = 26
        elif "pie" in chart_type:
            chart = PieChart()
        elif "line" in chart_type:
            chart = LineChart()
        elif "scatter" in chart_type:
            chart = ScatterChart()
        elif "bar" in chart_type:
            chart = BarChart()
            chart.type = "bar"
            chart.style = 10
        else:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            
        chart.title = f"Extracted {chart_type.title()} Chart"
        
        # Data ranges
        min_col, min_row = 1, 1
        max_col, max_row = len(df.columns), len(df) + 1
        
        # Detect Categories (String column)
        cat_col_idx = -1
        for idx, dtype in enumerate(df.dtypes):
            if dtype == 'object' or dtype == 'string':
                cat_col_idx = idx + 1
                break
        if cat_col_idx == -1: cat_col_idx = 1
        
        cats = Reference(worksheet, min_col=cat_col_idx, min_row=2, max_row=max_row)
        
        added_series = False
        series_idx = 0
        for idx, dtype in enumerate(df.dtypes):
            col_idx = idx + 1
            if col_idx == cat_col_idx: continue
            
            if pd.api.types.is_numeric_dtype(dtype):
                data = Reference(worksheet, min_col=col_idx, min_row=1, max_row=max_row)
                chart.add_data(data, titles_from_data=True)
                added_series = True
                
        if added_series:
            chart.set_categories(cats)
            if "pie" in chart_type or "doughnut" in chart_type:
                chart.dataLabels = DataLabelList()
                chart.dataLabels.showVal = True
                chart.dataLabels.showPercent = True
                
                # Apply Colors to Pie/Doughnut Slices
                if colors:
                    from openpyxl.chart.shapes import GraphicalProperties
                    # Removed invalid import: openpyxl.drawing.header_footer 
                    
                    # For pie charts, colors specific to data points are a bit complex in openpyxl
                    # We iterate through the slices of the first series
                    try:
                        # Ensure we have enough colors
                        series = chart.series[0]
                        for i, pt in enumerate(series.dPt):
                            # dPt list might be empty initially, need to create DataPoints
                            pass
                        
                        # Re-constructing series with DataPoints for coloring
                        # openpyxl requires defining DataPoints explicitly to color them individually
                        from openpyxl.chart.series import DataPoint
                        from openpyxl.drawing.colors import ColorChoice 
                        
                        # Resetting dPt just in case or populating it
                        series.dPt = []
                        
                        for i in range(len(df)):
                            if i < len(colors):
                                color_hex = colors[i].replace("#", "")
                                pt = DataPoint(idx=i)
                                pt.graphicalProperties.solidFill = color_hex
                                series.dPt.append(pt)
                    except Exception as e:
                        print(f"Error applying colors: {e}")
            
            else:
                # For Bar/Line, colors are per series usually, or per point if categories are varied
                # Assuming simple 1 series behavior matching row colors for now if applicable
                # But typically bars are one color per series. 
                # If the image implies different colors per category (varyColors=True), we can try that.
                if colors and len(colors) > 0:
                     chart.varyColors = True
                     # Applying specific colors to points in a bar chart is similar to pie
                     try:
                        series = chart.series[0]
                        from openpyxl.chart.series import DataPoint
                        series.dPt = []
                        for i in range(len(df)):
                             if i < len(colors):
                                color_hex = colors[i].replace("#", "")
                                pt = DataPoint(idx=i)
                                pt.graphicalProperties.solidFill = color_hex
                                series.dPt.append(pt)
                     except:
                        pass

            worksheet.add_chart(chart, "E2")
            
    output.seek(0)
    return output

# --- STREAMLIT APP ---
st.set_page_config(page_title="Image to Excel Converter", page_icon="📊")

st.title("📊 Image to Excel Converter")
st.markdown("Upload an image of a chart or table, and we'll convert it to an Excel file with a **native chart**!")

# API Key Input
api_key = st.text_input("Enter Google Gemini API Key", type="password")
if api_key:
    api_key = api_key.strip()

uploaded_file = st.file_uploader("Choose an image...", type=["png", "jpg", "jpeg"])

if uploaded_file is not None and api_key:
    image = Image.open(uploaded_file)
    st.image(image, caption='Uploaded Image', use_column_width=True)
    
    if st.button('Convert to Excel'):
        chart_type, csv_string, colors = process_image(image, api_key)
        
        st.success(f"Detected Chart Type: **{chart_type.title()}**")
        if colors:
             st.write("Detected Colors:", colors)
        
        try:
            df = pd.read_csv(io.StringIO(csv_string))
            st.dataframe(df)
            
            excel_data = generate_excel(df, chart_type, colors)
            
            st.download_button(
                label="📥 Download Excel File",
                data=excel_data,
                file_name="converted_chart.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Error processing data: {e}")
            st.text(csv_string)

elif uploaded_file is not None and not api_key:
    st.warning("Please enter your API Key to proceed.")
