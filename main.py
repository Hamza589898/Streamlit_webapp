import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import matplotlib.pyplot as plt
import io
import csv
from datetime import datetime
import numpy as np
import traceback
import warnings
import sys
import base64

warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="Finance Automation",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
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
</style>
""", unsafe_allow_html=True)

if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'captured_figures' not in st.session_state:
    st.session_state.captured_figures = []
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
    timestamp = datetime.now().strftime("%H:%M:%S")
    icon = {'info': 'INFO', 'success': 'SUCCESS', 'error': 'ERROR', 'warning': 'WARNING'}.get(msg_type, 'INFO')
    st.session_state.console_logs.append(f"[{timestamp}] [{icon}] {message}")
    if len(st.session_state.console_logs) > 100:
        st.session_state.console_logs = st.session_state.console_logs[-100:]

def log_error(error_msg, traceback_str=None):
    timestamp = datetime.now().strftime("%H:%M:%S")
    error_entry = f"[{timestamp}] ERROR: {error_msg}"
    if traceback_str:
        error_entry += f"\n\nTraceback:\n{traceback_str}"
    st.session_state.error_logs.append(error_entry)
    if len(st.session_state.error_logs) > 20:
        st.session_state.error_logs = st.session_state.error_logs[-20:]

@st.cache_data(show_spinner=False)
def process_uploaded_file(file_bytes, file_name):
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
    def __init__(self, log_func):
        self.log_func = log_func
        self.original_stdout = sys.stdout
        
    def write(self, text):
        if text.strip():
            self.log_func(text.strip(), 'info')
    
    def flush(self):
        pass

class FigureCapture:
    def __init__(self):
        self.figures = []
        self.original_plt_show = None
    
    def capture_matplotlib(self):
        try:
            fig = plt.gcf()
            if fig.get_axes():
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=100, bbox_inches='tight')
                buf.seek(0)
                img_str = base64.b64encode(buf.read()).decode()
                self.figures.append(('matplotlib', img_str))
                log_to_console(f"Captured matplotlib figure", 'success')
                plt.close(fig)
        except Exception as e:
            log_error(f"Error capturing matplotlib figure: {str(e)}")
    
    def capture_plotly(self, fig):
        try:
            self.figures.append(('plotly', fig))
            log_to_console(f"Captured plotly figure", 'success')
        except Exception as e:
            log_error(f"Error capturing plotly figure: {str(e)}")

def execute_python_script(df, script):
    try:
        log_to_console("=" * 50, 'info')
        log_to_console("Starting script execution...", 'info')
        log_to_console(f"Input data: {len(df)} rows × {len(df.columns)} columns", 'info')
        log_to_console("=" * 50, 'info')
        
        fig_capture = FigureCapture()
        
        def custom_plt_show(*args, **kwargs):
            fig_capture.capture_matplotlib()
        
        original_plotly_show = go.Figure.show
        def custom_plotly_show(self, *args, **kwargs):
            fig_capture.capture_plotly(self)
        
        exec_globals = {
            'input_df': df.copy(),
            'pd': pd,
            'np': np,
            'datetime': datetime,
            'io': io,
            'csv': csv,
            'plt': plt,
            'go': go,
            'px': px,
            'matplotlib': plt,
        }
        
        exec_locals = {}
        
        old_stdout = sys.stdout
        sys.stdout = PrintCapture(log_to_console)
        
        original_plt_show = plt.show
        plt.show = custom_plt_show
        
        go.Figure.show = custom_plotly_show
        
        try:
            exec(script, exec_globals, exec_locals)
            
            if plt.get_fignums():
                fig_capture.capture_matplotlib()
            
        finally:
            sys.stdout = old_stdout
            plt.show = original_plt_show
            go.Figure.show = original_plotly_show
        
        output_df = None
        
        if 'output_df' in exec_locals:
            output_df = exec_locals['output_df']
        elif 'output_df' in exec_globals:
            output_df = exec_globals['output_df']
        
        if output_df is None:
            error_msg = "Script must define 'output_df' variable"
            log_to_console(error_msg, 'error')
            log_error(error_msg, "No output_df found after script execution")
            return None, []
        
        if not isinstance(output_df, pd.DataFrame):
            error_msg = f"output_df must be a DataFrame, got {type(output_df)}"
            log_to_console(error_msg, 'error')
            log_error(error_msg)
            return None, []
        
        log_to_console("=" * 50, 'success')
        log_to_console(f"Output generated: {len(output_df)} rows × {len(output_df.columns)} columns", 'success')
        if fig_capture.figures:
            log_to_console(f"Captured {len(fig_capture.figures)} visualization(s)", 'success')
        log_to_console("=" * 50, 'success')
        
        return output_df, fig_capture.figures
            
    except SyntaxError as e:
        error_msg = f"Syntax Error on line {e.lineno}: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None, []
    except NameError as e:
        error_msg = f"Name Error: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None, []
    except Exception as e:
        error_msg = f"Execution failed: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None, []
    finally:
        sys.stdout = old_stdout
        plt.close('all')

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
            st.session_state.captured_figures = []
            st.session_state.script_executed = False
            st.session_state.current_file_names = []
    
    st.markdown("---")
    
    st.markdown('<p class="section-header">Python Script</p>', unsafe_allow_html=True)
    
    default_script = ""
    
    code_editor = st.text_area(
        "Python Script",
        value=default_script,
        height=400,
        help="Write Python code. Use plt.show() or fig.show() to display charts",
        label_visibility="collapsed",
        placeholder="# Write your Python code here\n# Your data is available as 'input_df'\n# Set 'output_df' to define the output"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        run_disabled = st.session_state.combined_data is None or not code_editor.strip()
        
        if st.button("Run Script", type="primary", use_container_width=True, disabled=run_disabled):
            if st.session_state.combined_data is not None and code_editor.strip():
                st.session_state.console_logs = []
                st.session_state.error_logs = []
                st.session_state.processed_data = None
                st.session_state.captured_figures = []
                st.session_state.script_executed = False
                
                with st.spinner('Executing...'):
                    result, figures = execute_python_script(st.session_state.combined_data, code_editor)
                    if result is not None:
                        st.session_state.processed_data = result
                        st.session_state.captured_figures = figures
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
            if st.session_state.captured_figures:
                st.markdown("### Visualizations")
                st.markdown(f"*{len(st.session_state.captured_figures)} chart(s) generated*")
                st.markdown("")
                
                for idx, (fig_type, fig_data) in enumerate(st.session_state.captured_figures):
                    if fig_type == 'matplotlib':
                        st.image(f"data:image/png;base64,{fig_data}", use_container_width=True)
                    elif fig_type == 'plotly':
                        st.plotly_chart(fig_data, use_container_width=True)
                    
                    if idx < len(st.session_state.captured_figures) - 1:
                        st.markdown("<br>", unsafe_allow_html=True)
                
                st.markdown("---")
                st.markdown("")
            else:
                st.info("No visualizations. Add plt.show() or fig.show() to your script.")
                st.markdown("---")
            
            st.markdown("### Data Preview")
            preview_rows = st.slider("Preview rows:", 5, 50, 10, key="preview_slider")
            st.dataframe(st.session_state.processed_data.head(preview_rows), use_container_width=True)
            
            st.markdown("")
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
        st.markdown('<div style="display: flex; align-items: center; justify-content: center; height: 400px; color: #6b7280; font-size: 16px;">Upload files and run script to see output</div>', unsafe_allow_html=True)
