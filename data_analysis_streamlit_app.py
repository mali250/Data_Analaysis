# import streamlit as st
# import pandas as pd
# import numpy as np
# import plotly.express as px
# import plotly.graph_objects as go
# from plotly.subplots import make_subplots
# import openai
# from openai import OpenAI
# import json
# import io
# from typing import Dict, List, Any
# import seaborn as sns
# import matplotlib.pyplot as plt
# import os
# from dotenv import load_dotenv

# load_dotenv()

# # Check for required dependencies
# try:
#     import openpyxl
# except ImportError:
#     st.error("Missing required dependency 'openpyxl'. Please install it using: pip install openpyxl")
#     st.stop()

# try:
#     import xlrd
# except ImportError:
#     st.warning("Optional dependency 'xlrd' not found. .xls files may not work. Install with: pip install xlrd")

# # Check for optional statistical dependencies
# HAS_STATSMODELS = True
# try:
#     import statsmodels.api as sm
# except ImportError:
#     HAS_STATSMODELS = False

# # Configure page
# st.set_page_config(
#     page_title="Excel Data Analysis Automation",
#     page_icon="📊",
#     layout="wide",
#     initial_sidebar_state="expanded"
# )

# class ExcelAnalyzer:
#     def __init__(self, api_key: str):
#         self.client = OpenAI(api_key=api_key)
        
#     def read_excel_file(self, uploaded_file) -> Dict[str, pd.DataFrame]:
#         """Read all sheets from uploaded Excel file"""
#         try:
#             # Determine file extension
#             file_name = uploaded_file.name.lower()
            
#             if file_name.endswith('.xlsx'):
#                 # Use openpyxl engine for .xlsx files
#                 excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
#             elif file_name.endswith('.xls'):
#                 # Use xlrd engine for .xls files
#                 try:
#                     excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine='xlrd')
#                 except Exception:
#                     st.warning("Could not read .xls file with xlrd. Trying with openpyxl...")
#                     excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
#             else:
#                 excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            
#             return excel_data
#         except ImportError as e:
#             if 'openpyxl' in str(e):
#                 st.error("Missing required dependency 'openpyxl'. Please install it using:\n```\npip install openpyxl\n```")
#             elif 'xlrd' in str(e):
#                 st.error("Missing required dependency 'xlrd' for .xls files. Please install it using:\n```\npip install xlrd\n```")
#             else:
#                 st.error(f"Missing dependency: {str(e)}")
#             return {}
#         except Exception as e:
#             st.error(f"Error reading Excel file: {str(e)}")
#             st.info("Please ensure your file is a valid Excel file (.xlsx or .xls)")
#             return {}
    
#     def get_data_summary(self, df: pd.DataFrame) -> Dict[str, Any]:
#         """Generate comprehensive data summary"""
#         summary = {
#             "shape": df.shape,
#             "columns": list(df.columns),
#             "dtypes": df.dtypes.to_dict(),
#             "missing_values": df.isnull().sum().to_dict(),
#             "numeric_columns": list(df.select_dtypes(include=[np.number]).columns),
#             "categorical_columns": list(df.select_dtypes(include=['object']).columns),
#             "memory_usage": df.memory_usage(deep=True).sum(),
#         }
        
#         # Add statistical summary for numeric columns
#         if summary["numeric_columns"]:
#             summary["statistics"] = df[summary["numeric_columns"]].describe().to_dict()
        
#         return summary
    
#     def analyze_with_gpt(self, data_summary: Dict, sample_data: str) -> Dict[str, Any]:
#         """Use GPT to analyze data and provide insights"""
        
#         prompt = f"""
#         As a data analyst, analyze this Excel dataset and provide comprehensive insights:
        
#         Dataset Summary:
#         - Shape: {data_summary['shape']}
#         - Columns: {data_summary['columns']}
#         - Data Types: {data_summary['dtypes']}
#         - Missing Values: {data_summary['missing_values']}
#         - Numeric Columns: {data_summary['numeric_columns']}
#         - Categorical Columns: {data_summary['categorical_columns']}
        
#         Sample Data:
#         {sample_data}
        
#         Statistical Summary:
#         {data_summary.get('statistics', 'No numeric data available')}
        
#         Please provide a JSON response with the following structure:
#         {{
#             "key_insights": [
#                 "insight 1",
#                 "insight 2",
#                 "insight 3"
#             ],
#             "data_quality_issues": [
#                 "issue 1",
#                 "issue 2"
#             ],
#             "recommended_visualizations": [
#                 {{
#                     "chart_type": "histogram",
#                     "columns": ["column_name"],
#                     "purpose": "distribution analysis"
#                 }},
#                 {{
#                     "chart_type": "scatter",
#                     "columns": ["x_column", "y_column"],
#                     "purpose": "correlation analysis"
#                 }}
#             ],
#             "business_recommendations": [
#                 "recommendation 1",
#                 "recommendation 2"
#             ],
#             "anomalies_detected": [
#                 "anomaly 1"
#             ]
#         }}
#         """
        
#         try:
#             response = self.client.chat.completions.create(
#                 model="gpt-4",
#                 messages=[
#                     {"role": "system", "content": "You are an expert data analyst. Provide insights in valid JSON format only and also Details Description."},
#                     {"role": "user", "content": prompt}
#                 ],
#                 temperature=0.6
#             )
            
#             analysis = json.loads(response.choices[0].message.content)
#             return analysis
#         except Exception as e:
#             st.error(f"Error in GPT analysis: {str(e)}")
#             return self.get_fallback_analysis(data_summary)
    
#     def get_fallback_analysis(self, data_summary: Dict) -> Dict[str, Any]:
#         """Provide basic analysis if GPT fails"""
#         return {
#             "key_insights": [
#                 f"Dataset contains {data_summary['shape'][0]} rows and {data_summary['shape'][1]} columns",
#                 f"Found {len(data_summary['numeric_columns'])} numeric columns for analysis",
#                 f"Data quality: {sum(data_summary['missing_values'].values())} missing values total"
#             ],
#             "data_quality_issues": [
#                 f"Missing values found in: {[k for k, v in data_summary['missing_values'].items() if v > 0]}"
#             ],
#             "recommended_visualizations": [
#                 {"chart_type": "histogram", "columns": data_summary['numeric_columns'][:2], "purpose": "distribution analysis"},
#                 {"chart_type": "correlation", "columns": data_summary['numeric_columns'], "purpose": "correlation analysis"}
#             ],
#             "business_recommendations": [
#                 "Clean missing values before analysis",
#                 "Focus on key numeric metrics for insights"
#             ],
#             "anomalies_detected": []
#         }
    
#     def create_visualizations(self, df: pd.DataFrame, recommendations: List[Dict]) -> List[go.Figure]:
#         """Generate comprehensive visualizations based on GPT recommendations and data analysis"""
#         figures = []
        
#         numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
#         categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
#         datetime_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
        
#         # Clean numeric columns - remove columns with all NaN values
#         clean_numeric_cols = []
#         for col in numeric_cols:
#             if not df[col].isna().all() and df[col].notna().sum() > 0:
#                 clean_numeric_cols.append(col)
#         numeric_cols = clean_numeric_cols
        
#         # Process GPT recommendations first
#         for rec in recommendations[:4]:  # Process some GPT recommendations
#             try:
#                 chart_type = rec.get('chart_type', '').lower()
#                 columns = rec.get('columns', [])
                
#                 if chart_type == 'histogram' and columns and columns[0] in numeric_cols:
#                     fig = px.histogram(df, x=columns[0], title=f"Distribution of {columns[0]}", 
#                                      marginal="box", nbins=30)
#                     figures.append(fig)
                
#                 elif chart_type == 'scatter' and len(columns) >= 2 and all(col in numeric_cols for col in columns[:2]):
#                     # Create scatter plot with or without trendline based on statsmodels availability
#                     if HAS_STATSMODELS:
#                         fig = px.scatter(df, x=columns[0], y=columns[1], title=f"{columns[0]} vs {columns[1]}",
#                                        trendline="ols", marginal_x="histogram", marginal_y="histogram")
#                     else:
#                         fig = px.scatter(df, x=columns[0], y=columns[1], title=f"{columns[0]} vs {columns[1]}",
#                                        marginal_x="histogram", marginal_y="histogram")
#                     figures.append(fig)
                
#                 elif chart_type == 'bar' and columns and columns[0] in categorical_cols:
#                     value_counts = df[columns[0]].value_counts().head(15)
#                     fig = px.bar(x=value_counts.index, y=value_counts.values, 
#                                title=f"Top 15 {columns[0]}", color=value_counts.values,
#                                color_continuous_scale="viridis")
#                     figures.append(fig)
                
#             except Exception as e:
#                 st.warning(f"Could not create {chart_type} chart: {str(e)}")
        
#         # Add comprehensive default visualizations
#         try:
#             # 1. Correlation Heatmap (if multiple numeric columns)
#             if len(numeric_cols) >= 2:
#                 # Filter out columns with insufficient data for correlation
#                 corr_data = df[numeric_cols].dropna()
#                 if len(corr_data) > 1:
#                     corr_matrix = corr_data.corr()
#                     fig = px.imshow(corr_matrix, text_auto=True, title="📊 Correlation Matrix",
#                                    color_continuous_scale="RdBu_r", aspect="auto")
#                     fig.update_layout(width=800, height=600)
#                     figures.append(fig)
            
#             # 2. Distribution Analysis for numeric columns
#             if len(numeric_cols) >= 1:
#                 # Create subplots for distributions
#                 cols_to_plot = numeric_cols[:6]  # Limit to 6 for readability
#                 if len(cols_to_plot) > 0:
#                     fig = make_subplots(
#                         rows=min(3, len(cols_to_plot)), 
#                         cols=min(2, max(1, len(cols_to_plot)//2 + 1)),
#                         subplot_titles=[f"Distribution of {col}" for col in cols_to_plot],
#                         vertical_spacing=0.15
#                     )
                    
#                     for i, col in enumerate(cols_to_plot):
#                         row = i // 2 + 1
#                         col_pos = i % 2 + 1
#                         # Filter out NaN values for histogram
#                         clean_data = df[col].dropna()
#                         if len(clean_data) > 0:
#                             fig.add_trace(
#                                 go.Histogram(x=clean_data, name=col, nbinsx=25),
#                                 row=row, col=col_pos
#                             )
                    
#                     fig.update_layout(height=800, title_text="📈 Distribution Analysis", showlegend=False)
#                     figures.append(fig)
            
#             # 3. Box Plots for Outlier Detection
#             if len(numeric_cols) >= 1:
#                 fig = go.Figure()
#                 for col in numeric_cols[:8]:  # Limit to 8 columns for readability
#                     clean_data = df[col].dropna()
#                     if len(clean_data) > 0:
#                         fig.add_trace(go.Box(y=clean_data, name=col, boxpoints="outliers"))
                
#                 fig.update_layout(title="📦 Box Plots - Outlier Detection", 
#                                 xaxis_title="Variables", yaxis_title="Values")
#                 figures.append(fig)
            
#             # 4. Categorical Analysis
#             if len(categorical_cols) >= 1:
#                 cols_to_plot = categorical_cols[:4]
#                 if len(cols_to_plot) > 0:
#                     fig = make_subplots(
#                         rows=min(2, len(cols_to_plot)), 
#                         cols=min(2, max(1, len(cols_to_plot)//2 + 1)),
#                         subplot_titles=[f"Distribution of {col}" for col in cols_to_plot],
#                         specs=[[{"type": "xy"}] * min(2, max(1, len(cols_to_plot)//2 + 1))] * min(2, len(cols_to_plot))
#                     )
                    
#                     for i, col in enumerate(cols_to_plot):
#                         row = i // 2 + 1
#                         col_pos = i % 2 + 1
#                         value_counts = df[col].value_counts().head(10)
                        
#                         if len(value_counts) > 0:
#                             fig.add_trace(
#                                 go.Bar(x=value_counts.index, y=value_counts.values, name=col),
#                                 row=row, col=col_pos
#                             )
                    
#                     fig.update_layout(height=600, title_text="📊 Categorical Variables Analysis", showlegend=False)
#                     figures.append(fig)
            
#             # 5. Statistical Summary Visualization
#             if len(numeric_cols) >= 2:
#                 stats_df = df[numeric_cols].describe().T
#                 fig = go.Figure()
                
#                 fig.add_trace(go.Bar(name='Mean', x=stats_df.index, y=stats_df['mean']))
#                 fig.add_trace(go.Bar(name='Std Dev', x=stats_df.index, y=stats_df['std']))
#                 fig.add_trace(go.Bar(name='Max', x=stats_df.index, y=stats_df['max']))
                
#                 fig.update_layout(title="📊 Statistical Summary Comparison", 
#                                 barmode='group', xaxis_title="Variables", yaxis_title="Values")
#                 figures.append(fig)
            
#             # 6. Missing Values Visualization
#             missing_data = df.isnull().sum()
#             if missing_data.sum() > 0:
#                 missing_data = missing_data[missing_data > 0].sort_values(ascending=True)
#                 if len(missing_data) > 0:
#                     fig = px.bar(x=missing_data.values, y=missing_data.index, 
#                                orientation='h', title="🚫 Missing Values by Column",
#                                color=missing_data.values, color_continuous_scale="Reds")
#                     fig.update_layout(xaxis_title="Number of Missing Values", yaxis_title="Columns")
#                     figures.append(fig)
            
#             # 7. Scatter Matrix (for key numeric variables)
#             if len(numeric_cols) >= 3:
#                 key_cols = numeric_cols[:4]  # Use top 4 numeric columns
#                 # Create scatter matrix with clean data
#                 clean_df = df[key_cols].dropna()
#                 if len(clean_df) > 1:
#                     fig = px.scatter_matrix(clean_df, title="🔍 Scatter Matrix - Key Variables", height=800)
#                     figures.append(fig)
            
#             # 8. Time Series Analysis (if datetime columns exist)
#             if datetime_cols and numeric_cols:
#                 try:
#                     df_time = df.copy()
#                     date_col = datetime_cols[0]
#                     value_col = numeric_cols[0]
                    
#                     # Clean the data for time series
#                     df_time = df_time[[date_col, value_col]].dropna()
#                     if len(df_time) > 1:
#                         df_time = df_time.sort_values(date_col)
#                         fig = px.line(df_time, x=date_col, y=value_col, 
#                                     title=f"📅 Time Series: {value_col} over {date_col}")
#                         figures.append(fig)
#                 except Exception as e:
#                     pass  # Skip time series if there are issues
            
#             # 9. Top/Bottom Analysis for numeric columns
#             if len(numeric_cols) >= 1 and df.shape[0] > 10:
#                 col = numeric_cols[0]
#                 clean_data = df[col].dropna()
#                 if len(clean_data) >= 20:  # Need sufficient data for top/bottom analysis
#                     top_10 = df.nlargest(10, col)
#                     bottom_10 = df.nsmallest(10, col)
                    
#                     fig = make_subplots(rows=1, cols=2, 
#                                       subplot_titles=[f"Top 10 - {col}", f"Bottom 10 - {col}"])
                    
#                     fig.add_trace(go.Bar(x=list(range(len(top_10))), y=top_10[col], 
#                                        name="Top 10", marker_color="green"), row=1, col=1)
#                     fig.add_trace(go.Bar(x=list(range(len(bottom_10))), y=bottom_10[col], 
#                                        name="Bottom 10", marker_color="red"), row=1, col=2)
                    
#                     fig.update_layout(title=f"🔝 Top vs Bottom Analysis - {col}", showlegend=False)
#                     figures.append(fig)
            
#             # 10. Data Quality Dashboard
#             quality_metrics = {
#                 'Total Rows': df.shape[0],
#                 'Total Columns': df.shape[1],
#                 'Missing Values': df.isnull().sum().sum(),
#                 'Duplicate Rows': df.duplicated().sum(),
#                 'Numeric Columns': len(numeric_cols),
#                 'Categorical Columns': len(categorical_cols)
#             }
            
#             fig = go.Figure(data=[
#                 go.Bar(name='Data Quality Metrics', 
#                       x=list(quality_metrics.keys()), 
#                       y=list(quality_metrics.values()),
#                       text=list(quality_metrics.values()),
#                       textposition='auto',
#                       marker_color=['blue', 'green', 'red', 'orange', 'purple', 'brown'])
#             ])
            
#             fig.update_layout(title="📋 Data Quality Dashboard", 
#                             xaxis_title="Metrics", yaxis_title="Count")
#             figures.append(fig)
            
#             # 11. Pie Charts for Categorical Variables
#             for col in categorical_cols[:2]:  # Top 2 categorical columns
#                 value_counts = df[col].value_counts().head(8)
#                 if len(value_counts) > 1:
#                     fig = px.pie(values=value_counts.values, names=value_counts.index,
#                                title=f"🥧 Distribution of {col}")
#                     figures.append(fig)
            
#             # 12. Advanced Scatter Plots with size and color (Fixed)
#             if len(numeric_cols) >= 3:
#                 try:
#                     # Prepare clean data for scatter plot
#                     x_col = numeric_cols[0]
#                     y_col = numeric_cols[1]
#                     size_col = numeric_cols[2] if len(numeric_cols) > 2 else None
#                     color_col = categorical_cols[0] if categorical_cols else None
                    
#                     # Create a clean dataset
#                     plot_cols = [x_col, y_col]
#                     if size_col:
#                         plot_cols.append(size_col)
#                     if color_col:
#                         plot_cols.append(color_col)
                    
#                     plot_df = df[plot_cols].dropna()
                    
#                     if len(plot_df) > 0:
#                         # Handle size column - ensure no NaN or negative values
#                         size_data = None
#                         if size_col and size_col in plot_df.columns:
#                             size_data = plot_df[size_col]
#                             # Replace any remaining NaN with median and ensure positive values
#                             size_data = size_data.fillna(size_data.median())
#                             size_data = np.abs(size_data) + 1  # Ensure positive values
                        
#                         fig = px.scatter(plot_df, x=x_col, y=y_col, 
#                                        size=size_data,
#                                        color=color_col,
#                                        title=f"🎯 Advanced Scatter: {x_col} vs {y_col}",
#                                        hover_data=numeric_cols[:4])
#                         figures.append(fig)
#                 except Exception as e:
#                     # Create simple scatter plot as fallback
#                     if len(numeric_cols) >= 2:
#                         clean_df = df[[numeric_cols[0], numeric_cols[1]]].dropna()
#                         if len(clean_df) > 0:
#                             fig = px.scatter(clean_df, x=numeric_cols[0], y=numeric_cols[1],
#                                            title=f"🎯 Scatter: {numeric_cols[0]} vs {numeric_cols[1]}")
#                             figures.append(fig)
            
#         except Exception as e:
#             st.warning(f"Error creating some visualizations: {str(e)}")
        
#         # Ensure we have at least some basic charts
#         if not figures and numeric_cols:
#             try:
#                 clean_data = df[numeric_cols[0]].dropna()
#                 if len(clean_data) > 0:
#                     fig = px.histogram(clean_data, title=f"Distribution of {numeric_cols[0]}")
#                     figures.append(fig)
#             except:
#                 pass
        
#         return figures

# def main():
#     st.title("📊 Excel Data Analysis Automation System")
#     st.markdown("Upload your Excel file and get AI-powered insights with automatic visualizations!")
    
#     # Sidebar configuration
#     with st.sidebar:
#         st.header("Configuration")

#                 # Try to get API key from environment variable first
#         api_key_env = os.getenv('OPENAI_API_KEY')
        
#         if api_key_env:
#             st.success("✅ OpenAI API Key loaded from environment")
#             api_key = api_key_env
#         else:
#             api_key = st.text_input("OpenAI API Key", type="password", help="Enter your OpenAI API key or set OPENAI_API_KEY environment variable")
            
#             if not api_key:
#                 st.warning("Please enter your OpenAI API key or set the OPENAI_API_KEY environment variable")
#                 st.info("To set environment variable, create a .env file with: OPENAI_API_KEY=your_key_here")
#                 st.stop()

        
#         # api_key = st.text_input("OpenAI API Key", type="password", help="Enter your OpenAI API key")
        
#         # if not api_key:
#         #     st.warning("Please enter your OpenAI API key to proceed")
#         #     st.stop()
    
#     # Initialize analyzer
#     analyzer = ExcelAnalyzer(api_key)
    
#     # File upload
#     uploaded_file = st.file_uploader(
#         "Choose an Excel file",
#         type=['xlsx', 'xls'],
#         help="Upload your Excel file for analysis"
#     )
    
#     if uploaded_file is not None:
#         with st.spinner("Reading Excel file..."):
#             excel_data = analyzer.read_excel_file(uploaded_file)
        
#         if not excel_data:
#             st.error("Failed to read the Excel file. Please check the format.")
#             return
        
#         # Sheet selection
#         sheet_names = list(excel_data.keys())
#         selected_sheet = st.selectbox("Select sheet to analyze:", sheet_names)
        
#         df = excel_data[selected_sheet]
        
#         # Display basic info
#         col1, col2, col3, col4 = st.columns(4)
#         with col1:
#             st.metric("Rows", df.shape[0])
#         with col2:
#             st.metric("Columns", df.shape[1])
#         with col3:
#             st.metric("Missing Values", df.isnull().sum().sum())
#         with col4:
#             st.metric("Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024:.1f} KB")
        
#         # Show sample data
#         st.subheader("📋 Sample Data")
#         st.dataframe(df.head(10), use_container_width=True)
        
#         # Analysis button
#         if st.button("🔍 Analyze Data with AI", type="primary"):
#             with st.spinner("Analyzing data with GPT..."):
#                 # Get data summary
#                 data_summary = analyzer.get_data_summary(df)
                
#                 # Prepare sample data for GPT
#                 sample_data = df.head(5).to_string()
                
#                 # Get GPT analysis
#                 analysis = analyzer.analyze_with_gpt(data_summary, sample_data)
                
#                 # Display results
#                 st.subheader("🎯 Key Insights")
#                 for insight in analysis.get('key_insights', []):
#                     st.write(f"• {insight}")
                
#                 col1, col2 = st.columns(2)
                
#                 with col1:
#                     st.subheader("⚠️ Data Quality Issues")
#                     issues = analysis.get('data_quality_issues', [])
#                     if issues:
#                         for issue in issues:
#                             st.warning(issue)
#                     else:
#                         st.success("No major data quality issues detected!")
                
#                 with col2:
#                     st.subheader("💡 Business Recommendations")
#                     for rec in analysis.get('business_recommendations', []):
#                         st.info(rec)
                
#                 # Anomalies
#                 anomalies = analysis.get('anomalies_detected', [])
#                 if anomalies:
#                     st.subheader("🚨 Anomalies Detected")
#                     for anomaly in anomalies:
#                         st.error(anomaly)
                
#                 # Generate and display visualizations
#                 st.subheader("📈 Comprehensive Data Visualizations")
                
#                 with st.spinner("Generating comprehensive charts and insights..."):
#                     viz_recommendations = analysis.get('recommended_visualizations', [])
#                     figures = analyzer.create_visualizations(df, viz_recommendations)
                
#                 if figures:
#                     # Create tabs for different chart categories
#                     tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "📈 Distributions", "🔍 Relationships", "📋 Data Quality"])
                    
#                     # Distribute charts across tabs
#                     overview_charts = []
#                     distribution_charts = []
#                     relationship_charts = []
#                     quality_charts = []
                    
#                     for i, fig in enumerate(figures):
#                         title = fig.layout.title.text.lower() if fig.layout.title else ""
                        
#                         if any(word in title for word in ['correlation', 'matrix', 'quality', 'dashboard']):
#                             overview_charts.append((i, fig))
#                         elif any(word in title for word in ['distribution', 'histogram', 'box', 'pie']):
#                             distribution_charts.append((i, fig))
#                         elif any(word in title for word in ['scatter', 'relationship', 'vs', 'time series']):
#                             relationship_charts.append((i, fig))
#                         elif any(word in title for word in ['missing', 'top', 'bottom', 'quality']):
#                             quality_charts.append((i, fig))
#                         else:
#                             overview_charts.append((i, fig))
                    
#                     with tab1:
#                         st.markdown("### 🎯 Key Overview Charts")
#                         for i, fig in overview_charts:
#                             st.plotly_chart(fig, use_container_width=True, key=f"overview_chart_{i}")
                    
#                     with tab2:
#                         st.markdown("### 📊 Distribution Analysis")
#                         for i, fig in distribution_charts:
#                             st.plotly_chart(fig, use_container_width=True, key=f"dist_chart_{i}")
                    
#                     with tab3:
#                         st.markdown("### 🔗 Relationships & Correlations")
#                         for i, fig in relationship_charts:
#                             st.plotly_chart(fig, use_container_width=True, key=f"rel_chart_{i}")
                    
#                     with tab4:
#                         st.markdown("### 🔍 Data Quality & Insights")
#                         for i, fig in quality_charts:
#                             st.plotly_chart(fig, use_container_width=True, key=f"quality_chart_{i}")
                    
#                     # Summary of visualizations created
#                     st.success(f"✅ Generated {len(figures)} comprehensive visualizations!")
                    
#                     # Chart summary
#                     with st.expander("📋 Chart Summary"):
#                         chart_types = {}
#                         for fig in figures:
#                             chart_type = "Unknown"
#                             if hasattr(fig, 'data') and fig.data:
#                                 chart_type = type(fig.data[0]).__name__
#                             chart_types[chart_type] = chart_types.get(chart_type, 0) + 1
                        
#                         for chart_type, count in chart_types.items():
#                             st.write(f"• **{chart_type}**: {count} chart(s)")
                    
#                 else:
#                     st.warning("No suitable visualizations could be generated for this dataset.")
                
#                 # Summary statistics
#                 st.subheader("📊 Statistical Summary")
#                 numeric_cols = df.select_dtypes(include=[np.number]).columns
#                 if len(numeric_cols) > 0:
#                     st.dataframe(df[numeric_cols].describe(), use_container_width=True)
#                 else:
#                     st.info("No numeric columns found for statistical summary.")
                
#                 # Download processed data
#                 st.subheader("💾 Download Results")
                
#                 # Create analysis report
#                 report = {
#                     "file_name": uploaded_file.name,
#                     "sheet_analyzed": selected_sheet,
#                     "analysis_results": analysis,
#                     "data_summary": {
#                         "shape": data_summary["shape"],
#                         "columns": data_summary["columns"],
#                         "missing_values": data_summary["missing_values"]
#                     }
#                 }
                
#                 report_json = json.dumps(report, indent=2)
#                 st.download_button(
#                     label="📄 Download Analysis Report (JSON)",
#                     data=report_json,
#                     file_name=f"analysis_report_{selected_sheet}.json",
#                     mime="application/json"
#                 )
                
#                 # Download cleaned data
#                 if df.isnull().sum().sum() > 0:
#                     cleaned_df = df.dropna()
#                     csv_buffer = io.StringIO()
#                     cleaned_df.to_csv(csv_buffer, index=False)
#                     st.download_button(
#                         label="🧹 Download Cleaned Data (CSV)",
#                         data=csv_buffer.getvalue(),
#                         file_name=f"cleaned_{selected_sheet}.csv",
#                         mime="text/csv"
#                     )

# if __name__ == "__main__":
#     main()


import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openai
from openai import OpenAI
import json
import io
from typing import Dict, List, Any
import seaborn as sns
import matplotlib.pyplot as plt
import os
from dotenv import load_dotenv

load_dotenv()

# Check for required dependencies
try:
    import openpyxl
except ImportError:
    st.error("Missing required dependency 'openpyxl'. Please install it using: pip install openpyxl")
    st.stop()

try:
    import xlrd
except ImportError:
    st.warning("Optional dependency 'xlrd' not found. .xls files may not work. Install with: pip install xlrd")

# Check for optional statistical dependencies
HAS_STATSMODELS = True
try:
    import statsmodels.api as sm
except ImportError:
    HAS_STATSMODELS = False

# Configure page
st.set_page_config(
    page_title="Excel Data Analysis Automation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class ExcelAnalyzer:
    def __init__(self, api_key: str):
        self.client = OpenAI(api_key=api_key)
        
    def read_excel_file(self, uploaded_file) -> Dict[str, pd.DataFrame]:
        """Read all sheets from uploaded Excel file"""
        try:
            # Determine file extension
            file_name = uploaded_file.name.lower()
            
            if file_name.endswith('.xlsx'):
                # Use openpyxl engine for .xlsx files
                excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
            elif file_name.endswith('.xls'):
                # Use xlrd engine for .xls files
                try:
                    excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine='xlrd')
                except Exception:
                    st.warning("Could not read .xls file with xlrd. Trying with openpyxl...")
                    excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
            else:
                excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            
            return excel_data
        except ImportError as e:
            if 'openpyxl' in str(e):
                st.error("Missing required dependency 'openpyxl'. Please install it using:\n```\npip install openpyxl\n```")
            elif 'xlrd' in str(e):
                st.error("Missing required dependency 'xlrd' for .xls files. Please install it using:\n```\npip install xlrd\n```")
            else:
                st.error(f"Missing dependency: {str(e)}")
            return {}
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            st.info("Please ensure your file is a valid Excel file (.xlsx or .xls)")
            return {}
    
    def get_data_summary(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Generate comprehensive data summary"""
        summary = {
            "shape": df.shape,
            "columns": list(df.columns),
            "dtypes": df.dtypes.to_dict(),
            "missing_values": df.isnull().sum().to_dict(),
            "numeric_columns": list(df.select_dtypes(include=[np.number]).columns),
            "categorical_columns": list(df.select_dtypes(include=['object']).columns),
            "memory_usage": df.memory_usage(deep=True).sum(),
        }
        
        # Add statistical summary for numeric columns
        if summary["numeric_columns"]:
            summary["statistics"] = df[summary["numeric_columns"]].describe().to_dict()
        
        return summary
    
    def analyze_with_gpt(self, data_summary: Dict, sample_data: str) -> Dict[str, Any]:
        """Use GPT to analyze data and provide insights"""
        
        prompt = f"""
        As a data analyst, analyze this Excel dataset and provide comprehensive insights:
        
        Dataset Summary:
        - Shape: {data_summary['shape']}
        - Columns: {data_summary['columns']}
        - Data Types: {data_summary['dtypes']}
        - Missing Values: {data_summary['missing_values']}
        - Numeric Columns: {data_summary['numeric_columns']}
        - Categorical Columns: {data_summary['categorical_columns']}
        
        Sample Data:
        {sample_data}
        
        Statistical Summary:
        {data_summary.get('statistics', 'No numeric data available')}
        
        Please provide a JSON response with the following structure:
        {{
            "key_insights": [
                "insight 1",
                "insight 2",
                "insight 3"
            ],
            "data_quality_issues": [
                "issue 1",
                "issue 2"
            ],
            "recommended_visualizations": [
                {{
                    "chart_type": "histogram",
                    "columns": ["column_name"],
                    "purpose": "distribution analysis"
                }},
                {{
                    "chart_type": "scatter",
                    "columns": ["x_column", "y_column"],
                    "purpose": "correlation analysis"
                }}
            ],
            "business_recommendations": [
                "recommendation 1",
                "recommendation 2"
            ],
            "anomalies_detected": [
                "anomaly 1"
            ]
        }}
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are an expert data analyst. Provide insights in valid JSON format only and also Details Description."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.6
            )
            
            analysis = json.loads(response.choices[0].message.content)
            return analysis
        except Exception as e:
            st.error(f"Error in GPT analysis: {str(e)}")
            return self.get_fallback_analysis(data_summary)
    
    def get_fallback_analysis(self, data_summary: Dict) -> Dict[str, Any]:
        """Provide basic analysis if GPT fails"""
        return {
            "key_insights": [
                f"Dataset contains {data_summary['shape'][0]} rows and {data_summary['shape'][1]} columns",
                f"Found {len(data_summary['numeric_columns'])} numeric columns for analysis",
                f"Data quality: {sum(data_summary['missing_values'].values())} missing values total"
            ],
            "data_quality_issues": [
                f"Missing values found in: {[k for k, v in data_summary['missing_values'].items() if v > 0]}"
            ],
            "recommended_visualizations": [
                {"chart_type": "histogram", "columns": data_summary['numeric_columns'][:2], "purpose": "distribution analysis"},
                {"chart_type": "correlation", "columns": data_summary['numeric_columns'], "purpose": "correlation analysis"}
            ],
            "business_recommendations": [
                "Clean missing values before analysis",
                "Focus on key numeric metrics for insights"
            ],
            "anomalies_detected": []
        }
    
    def create_visualizations(self, df: pd.DataFrame, recommendations: List[Dict]) -> List[go.Figure]:
        """Generate comprehensive visualizations based on GPT recommendations and data analysis"""
        figures = []
        
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        datetime_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
        
        # Clean numeric columns - remove columns with all NaN values
        clean_numeric_cols = []
        for col in numeric_cols:
            if not df[col].isna().all() and df[col].notna().sum() > 0:
                clean_numeric_cols.append(col)
        numeric_cols = clean_numeric_cols
        
        # Process GPT recommendations first
        for rec in recommendations[:4]:  # Process some GPT recommendations
            try:
                chart_type = rec.get('chart_type', '').lower()
                columns = rec.get('columns', [])
                
                if chart_type == 'histogram' and columns and columns[0] in numeric_cols:
                    fig = px.histogram(df, x=columns[0], title=f"Distribution of {columns[0]}", 
                                     marginal="box", nbins=30)
                    figures.append(fig)
                
                elif chart_type == 'scatter' and len(columns) >= 2 and all(col in numeric_cols for col in columns[:2]):
                    # Create scatter plot with or without trendline based on statsmodels availability
                    if HAS_STATSMODELS:
                        fig = px.scatter(df, x=columns[0], y=columns[1], title=f"{columns[0]} vs {columns[1]}",
                                       trendline="ols", marginal_x="histogram", marginal_y="histogram")
                    else:
                        fig = px.scatter(df, x=columns[0], y=columns[1], title=f"{columns[0]} vs {columns[1]}",
                                       marginal_x="histogram", marginal_y="histogram")
                    figures.append(fig)
                
                elif chart_type == 'bar' and columns and columns[0] in categorical_cols:
                    value_counts = df[columns[0]].value_counts().head(15)
                    fig = px.bar(x=value_counts.index, y=value_counts.values, 
                               title=f"Top 15 {columns[0]}", color=value_counts.values,
                               color_continuous_scale="viridis")
                    figures.append(fig)
                
            except Exception as e:
                st.warning(f"Could not create {chart_type} chart: {str(e)}")
        
        # Add comprehensive default visualizations
        try:
            # 1. Correlation Heatmap (if multiple numeric columns)
            if len(numeric_cols) >= 2:
                # Filter out columns with insufficient data for correlation
                corr_data = df[numeric_cols].dropna()
                if len(corr_data) > 1:
                    corr_matrix = corr_data.corr()
                    fig = px.imshow(corr_matrix, text_auto=True, title="📊 Correlation Matrix",
                                   color_continuous_scale="RdBu_r", aspect="auto")
                    fig.update_layout(width=800, height=600)
                    figures.append(fig)
            
            # 2. Distribution Analysis for numeric columns
            if len(numeric_cols) >= 1:
                # Create subplots for distributions
                cols_to_plot = numeric_cols[:6]  # Limit to 6 for readability
                if len(cols_to_plot) > 0:
                    fig = make_subplots(
                        rows=min(3, len(cols_to_plot)), 
                        cols=min(2, max(1, len(cols_to_plot)//2 + 1)),
                        subplot_titles=[f"Distribution of {col}" for col in cols_to_plot],
                        vertical_spacing=0.15
                    )
                    
                    for i, col in enumerate(cols_to_plot):
                        row = i // 2 + 1
                        col_pos = i % 2 + 1
                        # Filter out NaN values for histogram
                        clean_data = df[col].dropna()
                        if len(clean_data) > 0:
                            fig.add_trace(
                                go.Histogram(x=clean_data, name=col, nbinsx=25),
                                row=row, col=col_pos
                            )
                    
                    fig.update_layout(height=800, title_text="📈 Distribution Analysis", showlegend=False)
                    figures.append(fig)
            
            # 3. Box Plots for Outlier Detection
            if len(numeric_cols) >= 1:
                fig = go.Figure()
                for col in numeric_cols[:8]:  # Limit to 8 columns for readability
                    clean_data = df[col].dropna()
                    if len(clean_data) > 0:
                        fig.add_trace(go.Box(y=clean_data, name=col, boxpoints="outliers"))
                
                fig.update_layout(title="📦 Box Plots - Outlier Detection", 
                                xaxis_title="Variables", yaxis_title="Values")
                figures.append(fig)
            
            # 4. Categorical Analysis
            if len(categorical_cols) >= 1:
                cols_to_plot = categorical_cols[:4]
                if len(cols_to_plot) > 0:
                    fig = make_subplots(
                        rows=min(2, len(cols_to_plot)), 
                        cols=min(2, max(1, len(cols_to_plot)//2 + 1)),
                        subplot_titles=[f"Distribution of {col}" for col in cols_to_plot],
                        specs=[[{"type": "xy"}] * min(2, max(1, len(cols_to_plot)//2 + 1))] * min(2, len(cols_to_plot))
                    )
                    
                    for i, col in enumerate(cols_to_plot):
                        row = i // 2 + 1
                        col_pos = i % 2 + 1
                        value_counts = df[col].value_counts().head(10)
                        
                        if len(value_counts) > 0:
                            fig.add_trace(
                                go.Bar(x=value_counts.index, y=value_counts.values, name=col),
                                row=row, col=col_pos
                            )
                    
                    fig.update_layout(height=600, title_text="📊 Categorical Variables Analysis", showlegend=False)
                    figures.append(fig)
            
            # 5. Statistical Summary Visualization
            if len(numeric_cols) >= 2:
                stats_df = df[numeric_cols].describe().T
                fig = go.Figure()
                
                fig.add_trace(go.Bar(name='Mean', x=stats_df.index, y=stats_df['mean']))
                fig.add_trace(go.Bar(name='Std Dev', x=stats_df.index, y=stats_df['std']))
                fig.add_trace(go.Bar(name='Max', x=stats_df.index, y=stats_df['max']))
                
                fig.update_layout(title="📊 Statistical Summary Comparison", 
                                barmode='group', xaxis_title="Variables", yaxis_title="Values")
                figures.append(fig)
            
            # 6. Missing Values Visualization
            missing_data = df.isnull().sum()
            if missing_data.sum() > 0:
                missing_data = missing_data[missing_data > 0].sort_values(ascending=True)
                if len(missing_data) > 0:
                    fig = px.bar(x=missing_data.values, y=missing_data.index, 
                               orientation='h', title="🚫 Missing Values by Column",
                               color=missing_data.values, color_continuous_scale="Reds")
                    fig.update_layout(xaxis_title="Number of Missing Values", yaxis_title="Columns")
                    figures.append(fig)
            
            # 7. Scatter Matrix (for key numeric variables)
            if len(numeric_cols) >= 3:
                key_cols = numeric_cols[:4]  # Use top 4 numeric columns
                # Create scatter matrix with clean data
                clean_df = df[key_cols].dropna()
                if len(clean_df) > 1:
                    fig = px.scatter_matrix(clean_df, title="🔍 Scatter Matrix - Key Variables", height=800)
                    figures.append(fig)
            
            # 8. Time Series Analysis (if datetime columns exist)
            if datetime_cols and numeric_cols:
                try:
                    df_time = df.copy()
                    date_col = datetime_cols[0]
                    value_col = numeric_cols[0]
                    
                    # Clean the data for time series
                    df_time = df_time[[date_col, value_col]].dropna()
                    if len(df_time) > 1:
                        df_time = df_time.sort_values(date_col)
                        fig = px.line(df_time, x=date_col, y=value_col, 
                                    title=f"📅 Time Series: {value_col} over {date_col}")
                        figures.append(fig)
                except Exception as e:
                    pass  # Skip time series if there are issues
            
            # 9. Top/Bottom Analysis for numeric columns
            if len(numeric_cols) >= 1 and df.shape[0] > 10:
                col = numeric_cols[0]
                clean_data = df[col].dropna()
                if len(clean_data) >= 20:  # Need sufficient data for top/bottom analysis
                    top_10 = df.nlargest(10, col)
                    bottom_10 = df.nsmallest(10, col)
                    
                    fig = make_subplots(rows=1, cols=2, 
                                      subplot_titles=[f"Top 10 - {col}", f"Bottom 10 - {col}"])
                    
                    fig.add_trace(go.Bar(x=list(range(len(top_10))), y=top_10[col], 
                                       name="Top 10", marker_color="green"), row=1, col=1)
                    fig.add_trace(go.Bar(x=list(range(len(bottom_10))), y=bottom_10[col], 
                                       name="Bottom 10", marker_color="red"), row=1, col=2)
                    
                    fig.update_layout(title=f"🔝 Top vs Bottom Analysis - {col}", showlegend=False)
                    figures.append(fig)
            
            # 10. Data Quality Dashboard
            quality_metrics = {
                'Total Rows': df.shape[0],
                'Total Columns': df.shape[1],
                'Missing Values': df.isnull().sum().sum(),
                'Duplicate Rows': df.duplicated().sum(),
                'Numeric Columns': len(numeric_cols),
                'Categorical Columns': len(categorical_cols)
            }
            
            fig = go.Figure(data=[
                go.Bar(name='Data Quality Metrics', 
                      x=list(quality_metrics.keys()), 
                      y=list(quality_metrics.values()),
                      text=list(quality_metrics.values()),
                      textposition='auto',
                      marker_color=['blue', 'green', 'red', 'orange', 'purple', 'brown'])
            ])
            
            fig.update_layout(title="📋 Data Quality Dashboard", 
                            xaxis_title="Metrics", yaxis_title="Count")
            figures.append(fig)
            
            # 11. Pie Charts for Categorical Variables
            for col in categorical_cols[:2]:  # Top 2 categorical columns
                value_counts = df[col].value_counts().head(8)
                if len(value_counts) > 1:
                    fig = px.pie(values=value_counts.values, names=value_counts.index,
                               title=f"🥧 Distribution of {col}")
                    figures.append(fig)
            
            # 12. Advanced Scatter Plots with size and color (Fixed)
            if len(numeric_cols) >= 3:
                try:
                    # Prepare clean data for scatter plot
                    x_col = numeric_cols[0]
                    y_col = numeric_cols[1]
                    size_col = numeric_cols[2] if len(numeric_cols) > 2 else None
                    color_col = categorical_cols[0] if categorical_cols else None
                    
                    # Create a clean dataset
                    plot_cols = [x_col, y_col]
                    if size_col:
                        plot_cols.append(size_col)
                    if color_col:
                        plot_cols.append(color_col)
                    
                    plot_df = df[plot_cols].dropna()
                    
                    if len(plot_df) > 0:
                        # Handle size column - ensure no NaN or negative values
                        size_data = None
                        if size_col and size_col in plot_df.columns:
                            size_data = plot_df[size_col]
                            # Replace any remaining NaN with median and ensure positive values
                            size_data = size_data.fillna(size_data.median())
                            size_data = np.abs(size_data) + 1  # Ensure positive values
                        
                        fig = px.scatter(plot_df, x=x_col, y=y_col, 
                                       size=size_data,
                                       color=color_col,
                                       title=f"🎯 Advanced Scatter: {x_col} vs {y_col}",
                                       hover_data=numeric_cols[:4])
                        figures.append(fig)
                except Exception as e:
                    # Create simple scatter plot as fallback
                    if len(numeric_cols) >= 2:
                        clean_df = df[[numeric_cols[0], numeric_cols[1]]].dropna()
                        if len(clean_df) > 0:
                            fig = px.scatter(clean_df, x=numeric_cols[0], y=numeric_cols[1],
                                           title=f"🎯 Scatter: {numeric_cols[0]} vs {numeric_cols[1]}")
                            figures.append(fig)
            
        except Exception as e:
            st.warning(f"Error creating some visualizations: {str(e)}")
        
        # Ensure we have at least some basic charts
        if not figures and numeric_cols:
            try:
                clean_data = df[numeric_cols[0]].dropna()
                if len(clean_data) > 0:
                    fig = px.histogram(clean_data, title=f"Distribution of {numeric_cols[0]}")
                    figures.append(fig)
            except:
                pass
        
        return figures

def main():
    st.title("📊 Excel Data Analysis Automation System")
    st.markdown("Upload your Excel file and get AI-powered insights with automatic visualizations!")
    
    # Sidebar configuration
    with st.sidebar:
        st.header("Configuration")

                # Try to get API key from environment variable first
        api_key_env = os.getenv('OPENAI_API_KEY')
        
        if api_key_env:
            st.success("✅ OpenAI API Key loaded from environment")
            api_key = api_key_env
        else:
            api_key = st.text_input("OpenAI API Key", type="password", help="Enter your OpenAI API key or set OPENAI_API_KEY environment variable")
            
            if not api_key:
                st.warning("Please enter your OpenAI API key or set the OPENAI_API_KEY environment variable")
                st.info("To set environment variable, create a .env file with: OPENAI_API_KEY=your_key_here")
                st.stop()

        
        # api_key = st.text_input("OpenAI API Key", type="password", help="Enter your OpenAI API key")
        
        # if not api_key:
        #     st.warning("Please enter your OpenAI API key to proceed")
        #     st.stop()
    
    # Initialize analyzer
    analyzer = ExcelAnalyzer(api_key)
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload your Excel file for analysis"
    )
    
    if uploaded_file is not None:
        with st.spinner("Reading Excel file..."):
            excel_data = analyzer.read_excel_file(uploaded_file)
        
        if not excel_data:
            st.error("Failed to read the Excel file. Please check the format.")
            return
        
        # Sheet selection
        sheet_names = list(excel_data.keys())
        selected_sheet = st.selectbox("Select sheet to analyze:", sheet_names)
        
        df = excel_data[selected_sheet]
        
        # Display basic info
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Rows", df.shape[0])
        with col2:
            st.metric("Columns", df.shape[1])
        with col3:
            st.metric("Missing Values", df.isnull().sum().sum())
        with col4:
            st.metric("Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024:.1f} KB")
        
        # Show sample data
        st.subheader("📋 Sample Data")
        st.dataframe(df.head(10), use_container_width=True)
        
        # Analysis button
        if st.button("🔍 Analyze Data with AI", type="primary"):
            with st.spinner("Analyzing data with SMART GPT..."):
                # Get data summary
                data_summary = analyzer.get_data_summary(df)
                
                # Prepare sample data for GPT
                sample_data = df.head(5).to_string()
                
                # Get GPT analysis
                analysis = analyzer.analyze_with_gpt(data_summary, sample_data)
                
                # Display results
                st.subheader("🎯 Key Insights")
                for insight in analysis.get('key_insights', []):
                    st.write(f"• {insight}")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("⚠️ Data Quality Issues")
                    issues = analysis.get('data_quality_issues', [])
                    if issues:
                        for issue in issues:
                            st.warning(issue)
                    else:
                        st.success("No major data quality issues detected!")
                
                with col2:
                    st.subheader("💡 Business Recommendations")
                    for rec in analysis.get('business_recommendations', []):
                        st.info(rec)
                
                # Anomalies
                anomalies = analysis.get('anomalies_detected', [])
                if anomalies:
                    st.subheader("🚨 Anomalies Detected")
                    for anomaly in anomalies:
                        st.error(anomaly)
                
                # Generate and display visualizations
                st.subheader("📈 Comprehensive Data Visualizations")
                
                with st.spinner("Generating comprehensive charts and insights..."):
                    viz_recommendations = analysis.get('recommended_visualizations', [])
                    figures = analyzer.create_visualizations(df, viz_recommendations)
                
                if figures:
                    # Create tabs for different chart categories
                    tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "📈 Distributions", "🔍 Relationships", "📋 Data Quality"])
                    
                    # Distribute charts across tabs
                    overview_charts = []
                    distribution_charts = []
                    relationship_charts = []
                    quality_charts = []
                    
                    for i, fig in enumerate(figures):
                        title = fig.layout.title.text.lower() if fig.layout.title else ""
                        
                        if any(word in title for word in ['correlation', 'matrix', 'quality', 'dashboard']):
                            overview_charts.append((i, fig))
                        elif any(word in title for word in ['distribution', 'histogram', 'box', 'pie']):
                            distribution_charts.append((i, fig))
                        elif any(word in title for word in ['scatter', 'relationship', 'vs', 'time series']):
                            relationship_charts.append((i, fig))
                        elif any(word in title for word in ['missing', 'top', 'bottom', 'quality']):
                            quality_charts.append((i, fig))
                        else:
                            overview_charts.append((i, fig))
                    
                    with tab1:
                        st.markdown("### 🎯 Key Overview Charts")
                        for i, fig in overview_charts:
                            st.plotly_chart(fig, use_container_width=True, key=f"overview_chart_{i}")
                    
                    with tab2:
                        st.markdown("### 📊 Distribution Analysis")
                        for i, fig in distribution_charts:
                            st.plotly_chart(fig, use_container_width=True, key=f"dist_chart_{i}")
                    
                    with tab3:
                        st.markdown("### 🔗 Relationships & Correlations")
                        for i, fig in relationship_charts:
                            st.plotly_chart(fig, use_container_width=True, key=f"rel_chart_{i}")
                    
                    with tab4:
                        st.markdown("### 🔍 Data Quality & Insights")
                        for i, fig in quality_charts:
                            st.plotly_chart(fig, use_container_width=True, key=f"quality_chart_{i}")
                    
                    # Summary of visualizations created
                    st.success(f"✅ Generated {len(figures)} comprehensive visualizations!")
                    
                    # Chart summary
                    with st.expander("📋 Chart Summary"):
                        chart_types = {}
                        for fig in figures:
                            chart_type = "Unknown"
                            if hasattr(fig, 'data') and fig.data:
                                chart_type = type(fig.data[0]).__name__
                            chart_types[chart_type] = chart_types.get(chart_type, 0) + 1
                        
                        for chart_type, count in chart_types.items():
                            st.write(f"• **{chart_type}**: {count} chart(s)")
                    
                else:
                    st.warning("No suitable visualizations could be generated for this dataset.")
                
                # Summary statistics
                st.subheader("📊 Statistical Summary")
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    st.dataframe(df[numeric_cols].describe(), use_container_width=True)
                else:
                    st.info("No numeric columns found for statistical summary.")
                
                # Download processed data
                st.subheader("💾 Download Results")
                
                # Create analysis report
                report = {
                    "file_name": uploaded_file.name,
                    "sheet_analyzed": selected_sheet,
                    "analysis_results": analysis,
                    "data_summary": {
                        "shape": data_summary["shape"],
                        "columns": data_summary["columns"],
                        "missing_values": data_summary["missing_values"]
                    }
                }
                
                report_json = json.dumps(report, indent=2)
                st.download_button(
                    label="📄 Download Analysis Report (JSON)",
                    data=report_json,
                    file_name=f"analysis_report_{selected_sheet}.json",
                    mime="application/json"
                )
                
                # Download cleaned data
                if df.isnull().sum().sum() > 0:
                    cleaned_df = df.dropna()
                    csv_buffer = io.StringIO()
                    cleaned_df.to_csv(csv_buffer, index=False)
                    st.download_button(
                        label="🧹 Download Cleaned Data (CSV)",
                        data=csv_buffer.getvalue(),
                        file_name=f"cleaned_{selected_sheet}.csv",
                        mime="text/csv"
                    )

if __name__ == "__main__":
    main()
