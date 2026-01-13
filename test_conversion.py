import os
import google.generativeai as genai
import pandas as pd
import io
from PIL import Image

def get_vision_model():
    """Finds the best available vision-capable model."""
    try:
        models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        print(f"Available models: {models}")
        
        # Priority list
        priorities = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro-vision', 'gemini-2.0-flash-exp']
        
        # Check for exact matches first (ignoring 'models/' prefix)
        for p in priorities:
            for m in models:
                if p in m:
                    return m
        
        # Fallback to any gemini model
        for m in models:
            if 'gemini' in m:
                return m
                
        return 'gemini-1.5-flash' # Absolute fallback
    except Exception as e:
        print(f"Error listing models: {e}")
        return 'gemini-1.5-flash'

def image_to_excel(image_path, output_path, api_key):
    # Configure Gemini
    genai.configure(api_key=api_key)
    
import time

def image_to_excel(image_path, output_path, api_key):
    # Configure Gemini
    genai.configure(api_key=api_key)
    
import time

def image_to_excel(image_path, output_path, api_key):
    # Configure Gemini
    genai.configure(api_key=api_key)
    
    print("Listing available models for this API key...")
    try:
        all_models = list(genai.list_models())
        vision_models = []
        for m in all_models:
            print(f"- {m.name} (Methods: {m.supported_generation_methods})")
            if 'generateContent' in m.supported_generation_methods:
                vision_models.append(m.name)
    except Exception as e:
        print(f"Error listing models: {e}")
        return

    print(f"\nFound {len(vision_models)} potential models: {vision_models}")
    
    if not vision_models:
        print("No models found with 'generateContent' capability.")
        return

    # Load the image
    try:
        img = Image.open(image_path)
    except Exception as e:
        print(f"Error loading image: {e}")
        return

    # Prompt
    prompt = """
    Analyze this image. It contains a chart or table. 
    1. Identify the type of chart. Choose strictly from: 
       ['bar', 'column', 'line', 'pie', 'doughnut', 'scatter', 'area', 'table']
       (Note: A 'doughnut' is a pie chart with a hole in the center).
    2. Extract the data efficiently and accurately.
    
    Output the result ONLY as a VALID JSON object with the following structure:
    {
        "chart_type": "detected_type",
        "csv_data": "raw_csv_string"
    }
    
    Do not include markdown code blocks (like ```json). 
    Do not include any introductory text. 
    """

    # Force specific model that worked
    vision_models = ['models/gemini-2.5-flash']
    
    for model_name in vision_models:   
        print(f"\nTrying model: {model_name}")
        model = genai.GenerativeModel(model_name)
        
        try:
            print("Sending request to Gemini...")
            response = model.generate_content([prompt, img])
            text_response = response.text
            
            # Cleanup potential markdown formatting
            text_response = text_response.replace('```json', '').replace('```', '').strip()
            
            import json
            try:
                result = json.loads(text_response)
                chart_type = result.get("chart_type", "bar").lower()
                csv_string = result.get("csv_data", "")
            except json.JSONDecodeError:
                print("Failed to parse JSON, assuming direct CSV output...")
                csv_string = text_response
                chart_type = "bar" 
            
            print(f"Detected Chart Type: {chart_type}")
            print("Parsing CSV...")
            
            # Convert CSV string to DataFrame
            df = pd.read_csv(io.StringIO(csv_string))
            
            # Save to Excel with Chart
            print(f"Saving to Excel with {chart_type} chart...")
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
                # Dynamic Chart Selection
                from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, DoughnutChart, Reference
                
                chart = None
                if "doughnut" in chart_type:
                    chart = DoughnutChart()
                    chart.style = 26 # Specific style for doughnut if available, or default
                elif "pie" in chart_type:
                    chart = PieChart()
                elif "line" in chart_type:
                    chart = LineChart()
                elif "scatter" in chart_type:
                    chart = ScatterChart()
                elif "bar" in chart_type: # Horizontal bar
                    chart = BarChart()
                    chart.type = "bar" 
                    chart.style = 10
                else: # Default column
                    chart = BarChart()
                    chart.type = "col" 
                    chart.style = 10
                
                chart.title = f"Extracted {chart_type.title()} Chart"
                
                # Find data ranges
                min_col, min_row = 1, 1
                max_col, max_row = len(df.columns), len(df) + 1 # +1 for header
                
                cats = None
                data = None
                
                # Heuristic: First object/string col is categories
                cat_col_idx = -1
                for idx, dtype in enumerate(df.dtypes):
                    if dtype == 'object' or dtype == 'string':
                        cat_col_idx = idx + 1 
                        break
                
                if cat_col_idx == -1: cat_col_idx = 1
                
                cats = Reference(worksheet, min_col=cat_col_idx, min_row=2, max_row=max_row)
                
                added_series = False
                for idx, dtype in enumerate(df.dtypes):
                    col_idx = idx + 1
                    if col_idx == cat_col_idx:
                        continue
                    
                    if pd.api.types.is_numeric_dtype(dtype):
                        data = Reference(worksheet, min_col=col_idx, min_row=1, max_row=max_row)
                        chart.add_data(data, titles_from_data=True)
                        added_series = True
                
                if added_series:
                    chart.set_categories(cats)
                    # Data labels for Pie/Doughnut
                    if "pie" in chart_type or "doughnut" in chart_type:
                         from openpyxl.chart.label import DataLabelList
                         chart.dataLabels = DataLabelList()
                         chart.dataLabels.showVal = True
                         
                    worksheet.add_chart(chart, "E2") 
                    print(f"{chart_type.title()} chart added to Excel sheet.")
                else:
                    print("Could not identify numeric columns for chart.")

            print(f"Successfully saved data to {output_path}")
            return # Success!
            
        except Exception as e:
            print(f"Error with {model_name}: {e}")
            if "429" in str(e) or "quota" in str(e).lower():
                print("Rate limit hit. Waiting 10 seconds before next attempt...")
                time.sleep(10)
            else:
                print("Trying next available model...")
                
    print("All available models failed.")

if __name__ == "__main__":
    # You can set this env var or hardcode it for testing
    api_key = os.environ.get("GOOGLE_API_KEY")
    
    image_file = "image (5).png"
    output_file = "output.xlsx"
    
    if not api_key:
        print("Error: GOOGLE_API_KEY environment variable not set.")
    else:
        if os.path.exists(image_file):
            image_to_excel(image_file, output_file, api_key)
        else:
            print(f"Error: Image file '{image_file}' not found.")
