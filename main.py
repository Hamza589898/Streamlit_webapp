import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import csv
from datetime import datetime
import numpy as np
import traceback
import warnings
import sys

# Suppress warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="Finance Automation",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS styling
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    [data-testid="column"]:first-child {
        background-color: #f3f4f6;
        padding: 25px;
        border-radius: 8px;
    }
    
    .section-header {
        font-size: 18px;
        font-weight: 600;
        color: #1f2937;
        margin-bottom: 15px;
        margin-top: 20px;
    }
    
    .console-box {
        background: #1e1e1e;
        color: #d4d4d4;
        padding: 15px;
        border-radius: 8px;
        font-family: 'Courier New', monospace;
        font-size: 12px;
        max-height: 400px;
        overflow-y: auto;
        line-height: 1.6;
        white-space: pre-wrap;
    }
    
    .error-console-box {
        background: #2d1818;
        color: #ff6b6b;
        padding: 15px;
        border-radius: 8px;
        font-family: 'Courier New', monospace;
        font-size: 11px;
        max-height: 400px;
        overflow-y: auto;
        border: 1px solid #ff4444;
        line-height: 1.5;
        white-space: pre-wrap;
    }
    
    .stDownloadButton button {
        background-color: #10b981;
        color: white;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
    
    .stTabs [data-baseweb="tab"] {
        font-size: 18px;
        font-weight: 600;
        padding: 12px 24px;
    }
    
    [data-testid="stFileUploader"] section div {
        display: none;
    }
    
    [data-testid="stFileUploader"] section {
        padding: 0;
    }
    
    .uploadedFile {
        margin-top: 10px;
    }
    
    .empty-state {
        display: flex;
        align-items: center;
        justify-content: center;
        height: 400px;
        color: #6b7280;
        font-size: 16px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'console_logs' not in st.session_state:
    st.session_state.console_logs = []
if 'error_logs' not in st.session_state:
    st.session_state.error_logs = []
if 'script_executed' not in st.session_state:
    st.session_state.script_executed = False
if 'combined_data' not in st.session_state:
    st.session_state.combined_data = None
if 'file_uploader_key' not in st.session_state:
    st.session_state.file_uploader_key = 0

def log_to_console(message, msg_type='info'):
    """Add messages to console log"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    icon = {'info': 'INFO', 'success': 'SUCCESS', 'error': 'ERROR', 'warning': 'WARNING'}.get(msg_type, 'INFO')
    st.session_state.console_logs.append(f"[{timestamp}] [{icon}] {message}")
    if len(st.session_state.console_logs) > 100:
        st.session_state.console_logs = st.session_state.console_logs[-100:]

def log_error(error_msg, traceback_str=None):
    """Add error to error console"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    error_entry = f"[{timestamp}] ERROR: {error_msg}"
    if traceback_str:
        error_entry += f"\n\nTraceback:\n{traceback_str}"
    st.session_state.error_logs.append(error_entry)
    if len(st.session_state.error_logs) > 20:
        st.session_state.error_logs = st.session_state.error_logs[-20:]

@st.cache_data(show_spinner=False)
def process_uploaded_file(file_bytes, file_name):
    """Process uploaded CSV or Excel file - cached for performance"""
    try:
        if file_name.endswith('.csv'):
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding='latin-1')
        elif file_name.endswith(('.xlsx', '.xls')):
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
            except:
                try:
                    df = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd')
                except:
                    df = pd.read_excel(io.BytesIO(file_bytes))
        else:
            return None
        
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        return None

def combine_uploaded_files(uploaded_files):
    """Combine multiple uploaded files into one dataframe"""
    try:
        if not uploaded_files:
            return None
        
        log_to_console(f"Processing {len(uploaded_files)} file(s)...", 'info')
        all_dfs = []
        
        for uploaded_file in uploaded_files:
            file_bytes = uploaded_file.read()
            df = process_uploaded_file(file_bytes, uploaded_file.name)
            if df is not None:
                all_dfs.append(df)
                log_to_console(f"Loaded: {uploaded_file.name} ({len(df)} rows)", 'success')
        
        if not all_dfs:
            return None
        
        if len(all_dfs) == 1:
            combined_df = all_dfs[0]
        else:
            combined_df = pd.concat(all_dfs, ignore_index=True)
            log_to_console(f"Combined {len(all_dfs)} files", 'success')
        
        log_to_console(f"Total: {len(combined_df)} rows, {len(combined_df.columns)} columns", 'info')
        return combined_df
    except Exception as e:
        error_msg = f"Error combining files: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None

class PrintCapture:
    """Capture print statements"""
    def __init__(self, log_func):
        self.log_func = log_func
        self.original_stdout = sys.stdout
        
    def write(self, text):
        if text.strip():
            self.log_func(text.strip(), 'info')
    
    def flush(self):
        pass

def execute_python_script(df, script):
    """Execute Python script on dataframe with better error handling"""
    try:
        log_to_console("=" * 50, 'info')
        log_to_console("Starting script execution...", 'info')
        log_to_console(f"Input data: {len(df)} rows × {len(df.columns)} columns", 'info')
        log_to_console("=" * 50, 'info')
        
        # Create execution environment
        exec_globals = {
            'input_df': df.copy(),
            'pd': pd,
            'np': np,
            'datetime': datetime,
            'io': io,
            'csv': csv,
        }
        
        exec_locals = {}
        
        # Redirect stdout to capture print statements
        old_stdout = sys.stdout
        sys.stdout = PrintCapture(log_to_console)
        
        try:
            # Execute the script
            exec(script, exec_globals, exec_locals)
        finally:
            # Restore stdout
            sys.stdout = old_stdout
        
        # Check for output
        output_df = None
        
        if 'output_df' in exec_locals:
            output_df = exec_locals['output_df']
        elif 'output_df' in exec_globals:
            output_df = exec_globals['output_df']
        
        if output_df is None:
            error_msg = "Script must define 'output_df' variable"
            log_to_console(error_msg, 'error')
            log_error(error_msg, "No output_df found after script execution")
            return None
        
        if not isinstance(output_df, pd.DataFrame):
            error_msg = f"output_df must be a DataFrame, got {type(output_df)}"
            log_to_console(error_msg, 'error')
            log_error(error_msg)
            return None
        
        log_to_console("=" * 50, 'success')
        log_to_console(f"Output generated: {len(output_df)} rows × {len(output_df.columns)} columns", 'success')
        log_to_console("=" * 50, 'success')
        
        return output_df
            
    except SyntaxError as e:
        error_msg = f"Syntax Error on line {e.lineno}: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None
    except NameError as e:
        error_msg = f"Name Error: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None
    except Exception as e:
        error_msg = f"Execution failed: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None
    finally:
        sys.stdout = old_stdout

def create_dynamic_chart(df):
    """Create interactive chart from dataframe"""
    if df is None or len(df) == 0:
        return None
    
    try:
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        if not numeric_cols:
            return None
        
        date_cols = [col for col in df.columns if any(term in col.lower() 
                     for term in ['date', 'time', 'timestamp', 'day', 'month', 'year'])]
        
        if date_cols:
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    x_data = pd.to_datetime(df[date_cols[0]], errors='coerce')
                
                if x_data.isna().all():
                    x_data = list(range(len(df)))
                    x_label = 'Row Index'
                else:
                    x_label = date_cols[0]
            except:
                x_data = list(range(len(df)))
                x_label = 'Row Index'
        else:
            x_data = list(range(len(df)))
            x_label = 'Row Index'
        
        fig = go.Figure()
        colors = ['#f59e0b', '#10b981', '#3b82f6', '#dc2626', '#6366f1']
        plot_limit = min(len(df), 100)
        
        for idx, col in enumerate(numeric_cols[:3]):
            fig.add_trace(go.Scatter(
                x=x_data[:plot_limit] if not isinstance(x_data, list) else x_data[:plot_limit],
                y=df[col][:plot_limit],
                mode='lines+markers',
                name=col.capitalize(),
                line=dict(color=colors[idx], width=2),
                marker=dict(size=6)
            ))
        
        fig.update_layout(
            title='Output Data Visualization',
            xaxis_title=x_label,
            yaxis_title='Value',
            height=400,
            hovermode='x unified',
            template='plotly_white',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        
        return fig
    except Exception as e:
        log_error(f"Chart error: {str(e)}", traceback.format_exc())
        return None

# Main Layout
st.title("Finance Automation")
st.markdown("<br>", unsafe_allow_html=True)

left_col, right_col = st.columns([1, 2])

with left_col:
    st.markdown('<p class="section-header">Upload Files</p>', unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader(
        "Upload Files",
        type=['csv', 'xlsx', 'xls'],
        accept_multiple_files=True,
        help="Maximum file size: 200MB per file",
        label_visibility="collapsed",
        key=f"file_uploader_{st.session_state.file_uploader_key}"
    )
    
    if uploaded_files:
        file_names = [f.name for f in uploaded_files]
        current_files = getattr(st.session_state, 'current_file_names', [])
        
        if file_names != current_files:
            st.session_state.current_file_names = file_names
            combined_df = combine_uploaded_files(uploaded_files)
            if combined_df is not None:
                st.session_state.combined_data = combined_df
        
        st.success(f"{len(uploaded_files)} file(s) uploaded")
        for file in uploaded_files:
            st.text(f"• {file.name}")
        
        if st.session_state.combined_data is not None:
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Rows", len(st.session_state.combined_data))
            with col2:
                st.metric("Total Columns", len(st.session_state.combined_data.columns))
        
        if st.button("Remove All Files", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.session_state.file_uploader_key = 1
            st.session_state.console_logs = []
            st.session_state.error_logs = []
            st.rerun()
    else:
        if st.session_state.combined_data is not None:
            st.session_state.combined_data = None
            st.session_state.processed_data = None
            st.session_state.script_executed = False
            st.session_state.current_file_names = []
    
    st.markdown("---")
    
    st.markdown('<p class="section-header">Python Script</p>', unsafe_allow_html=True)
    
    default_script = """# Copy input data
df = input_df.copy()

# Show what we have
print(f"Input shape: {df.shape}")
print(f"Columns: {list(df.columns)}")

# Basic transformation
if len(df) > 0:
    # Add row number
    df['row_number'] = range(1, len(df) + 1)
    print(f"Added row_number column")

# REQUIRED: Set output
output_df = df
print("Script completed successfully!")"""
    
    code_editor = st.text_area(
        "Python Script",
        value=default_script,
        height=300,
        help="Script must define 'output_df' variable. Available: input_df, pd, np, datetime",
        label_visibility="collapsed"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        run_disabled = st.session_state.combined_data is None or not code_editor.strip()
        
        if st.button("Run Script", type="primary", use_container_width=True, disabled=run_disabled):
            if st.session_state.combined_data is not None and code_editor.strip():
                st.session_state.console_logs = []
                st.session_state.error_logs = []
                st.session_state.processed_data = None
                st.session_state.script_executed = False
                
                with st.spinner('Executing...'):
                    result = execute_python_script(st.session_state.combined_data, code_editor)
                    if result is not None:
                        st.session_state.processed_data = result
                        st.session_state.script_executed = True
                st.rerun()
    
    with col2:
        if st.button("Clear Logs", use_container_width=True):
            st.session_state.console_logs = []
            st.session_state.error_logs = []
            st.rerun()
    
    st.markdown("---")
    
    st.markdown('<p class="section-header">Download Output</p>', unsafe_allow_html=True)
    
    if st.session_state.processed_data is not None:
        col1, col2 = st.columns(2)
        
        with col1:
            csv_data = st.session_state.processed_data.to_csv(index=False)
            st.download_button(
                label="Download CSV",
                data=csv_data,
                file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                st.session_state.processed_data.to_excel(writer, index=False, sheet_name='Output')
            excel_data = buffer.getvalue()
            
            st.download_button(
                label="Download Excel",
                data=excel_data,
                file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        st.success("Output ready!")
    else:
        st.info("Run script to generate output")

with right_col:
    if st.session_state.script_executed and st.session_state.processed_data is not None:
        tab1, tab2, tab3 = st.tabs(["Output", "Execution Log", "Error Details"])
        
        with tab1:
            st.markdown("### Output Visualization")
            
            chart = create_dynamic_chart(st.session_state.processed_data)
            if chart:
                st.plotly_chart(chart, use_container_width=True)
            else:
                st.info("No numeric columns available for visualization")
            
            st.markdown("---")
            
            st.markdown("### Data Preview")
            preview_rows = st.slider("Preview rows:", 5, 50, 10, key="preview_slider")
            st.dataframe(st.session_state.processed_data.head(preview_rows), use_container_width=True)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(st.session_state.processed_data))
            with col2:
                st.metric("Total Columns", len(st.session_state.processed_data.columns))
            with col3:
                memory_kb = st.session_state.processed_data.memory_usage(deep=True).sum() / 1024
                st.metric("Memory", f"{memory_kb:.1f} KB")
        
        with tab2:
            if st.session_state.console_logs:
                console_html = '<div class="console-box">'
                console_html += '\n'.join(st.session_state.console_logs)
                console_html += '</div>'
                st.markdown(console_html, unsafe_allow_html=True)
            else:
                st.info("No execution logs available")
        
        with tab3:
            if st.session_state.error_logs:
                error_html = '<div class="error-console-box">'
                error_html += '\n\n'.join(st.session_state.error_logs)
                error_html += '</div>'
                st.markdown(error_html, unsafe_allow_html=True)
            else:
                st.success("No errors detected")
    else:
        st.markdown('<div class="empty-state">Upload files and run script to see output</div>', unsafe_allow_html=True)