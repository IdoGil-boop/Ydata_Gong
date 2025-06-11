import streamlit as st
import pandas as pd
import numpy as np
from openai import OpenAI
import json
import os
from datasets import load_dataset
import re
import matplotlib.pyplot as plt
import io
import base64
from dotenv import load_dotenv

# Set page config must be the first Streamlit command
st.set_page_config(page_title="Data Analysis Agent", layout="wide")

# Load environment variables from .env file
load_dotenv()

# Get the Nebius API key from environment variables
NUBIUS_API_KEY = os.getenv('NUBIUS_API_KEY')

if NUBIUS_API_KEY:
    try:
        client = OpenAI(
            base_url='https://api.studio.nebius.com/v1/',
            api_key=NUBIUS_API_KEY
        )
    except Exception as e:
        st.error(f"Error initializing OpenAI client: {e}")
        client = None
else:
    st.error("NUBIUS_API_KEY not found in environment variables")
    client = None

MODEL_NAME = 'Qwen/Qwen3-30B-A3B'

st.markdown(
    """
    <style>
    .stChatMessage {
        margin-bottom: 1.5em;
    }
    .user-message {
        background: #e6f0fa;
        border-radius: 8px;
        padding: 0.7em 1em;
        margin-bottom: 0.5em;
        width: fit-content;
        max-width: 80%;
        align-self: flex-end;
    }
    .assistant-message {
        background: #f6f6f6;
        border-radius: 8px;
        padding: 0.7em 1em;
        margin-bottom: 0.5em;
        width: fit-content;
        max-width: 80%;
        align-self: flex-start;
    }
    .stChatContainer {
        display: flex;
        flex-direction: column;
        height: 60vh;
        overflow-y: auto;
        border: 1px solid #eee;
        border-radius: 8px;
        padding: 1em;
        background: #fff;
        margin-bottom: 1em;
    }
    .stChatInputBox {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background: #fafbfc;
        padding: 1em 0.5em 0.5em 0.5em;
        border-top: 1px solid #eee;
        z-index: 100;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

@st.cache_data
def load_data():
    try:
        ds = load_dataset("bitext/Bitext-customer-support-llm-chatbot-training-dataset")
        df = ds['train'].to_pandas()
        return df
    except Exception as e:
        st.error(f"Error loading dataset: {str(e)}")
        return None

df_global = load_data()
conversation_history = []

if df_global is not None:
    df_columns = df_global.columns.tolist()
    quantitative_operations_list = [
        'count', 'sum', 'frequency', 'avg', 'max', 'min', 'median',
        'unique_counts', 'most_frequent', 'least_frequent'
    ]
    operation_types = quantitative_operations_list + ['example', 'summerize', 'unique_values']

    categories = df_global['category'].unique().tolist()
    intents = df_global['intent'].unique().tolist()
    flags = df_global['flags'].unique().tolist()

    system_prompt_text = f'''
    You are a senior data analyst.
    Your primary goal is to process business questions into data queries using the provided tools.
    You are deeply familiar with our dataset and its columns: {df_columns}.
    If a user's question can be answered using the 'convert_to_query' tool, you MUST use this tool.
    You can perform the following operations: {operation_types}.
    If the question is quantitative (e.g. how many..., what is the avg/sum...) make sure you use a quantitative operation from: {quantitative_operations_list}.
    Distribution question should invoke frequency operation.
    When relevant information is available in the conversation history, use it to improve the query.
    Trust previous findings in the conversation history (e.g. 
        if the user asks 'what is the most frequent category?' and follow-ups with 'give example from the most frequent category' - filter the data based on the previous answer and use the filtered data for the example operation; 
        if the user asks about the same question again, use the previous answer).
    If the user question is out-of-scope or not related to the dataset, include 'OOS' in your response.
    For 'most_frequent' and 'least_frequent' operations, you can specify the number of top/bottom values to return (default is 3 if not specified).
    '''

    tools = [
        {
            "type": "function",
            "function": {
                "name": "convert_to_query",
                "description": "Filters, groups, and performs operations on a column based on user request",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "target_col": {
                            "type": "string",
                            "enum": df_columns,
                            "description": "The column to analyze"
                        },
                        "operation": {
                            "type": "string",
                            "enum": operation_types,
                            "description": "The operation to perform"
                        },
                        "group by": {
                            "type": "array",
                            "items": {"type": "string", "enum": df_columns},
                            "description": "Optional column(s) to group by"
                        },
                        "filter": {
                            "type": "object",
                             "properties": {
                                "category": {
                                    "type": "array",
                                    "items": {"type": "string", "enum": categories},
                                    "description": "Filter by category values"
                                },
                                "intent": {
                                    "type": "array",
                                    "items": {"type": "string", "enum": intents},
                                    "description": "Filter by intent values"
                                },
                                "flag": {
                                    "type": "array",
                                    "items": {"type": "string", "enum": flags},
                                    "description": "Filter by flag values"
                                }
                            },
                            "additionalProperties": {
                                "type": "array",
                                "items": {"type": "string"},
                            },
                            "description": "Key-value filter conditions"
                        },
                        "top_n": {
                            "type": "integer",
                            "description": "Number of top/bottom frequent values to return (for most_frequent/least_frequent). Optional."
                        }
                    },
                    "required": ["target_col", "operation"]
                }
            }
        }
    ]
else:
    df_columns = []
    query_types_enum_list = []
    quantitative_operations_list = []
    system_prompt_text = "Error: Dataset not loaded. System prompt cannot be initialized."
    tools = []

def summerize_with_model_func(df_to_summarize, target_col_to_summarize):
    if not client:
        return "OpenAI client not initialized."
    system_prompt_summarizer = (
        "You are an expert text summarization AI.\n"
        "Your sole task is to provide a concise summary of the given text.\n"
        "Do not cite the given text itself - but rather give a short summary of it in your own words.\n"
        "Output ONLY the summary. Do NOT include any meta-commentary, thinking steps (e.g., <think> tags), or self-correction.\n"
        "Begin the response directly with the summary text. Again - DO NOT INCLUDE THINKING PROCESS"
    )
    try:
        if isinstance(df_to_summarize, pd.DataFrame) and not df_to_summarize.empty:
            target_column_series = df_to_summarize[target_col_to_summarize].astype(str)
            input_text = "\n".join(target_column_series.tolist())
            if len(input_text) > 20000:
                avg_len_char_per_row = target_column_series.map(len).mean()
                n_rows_to_sample = int(15000 // avg_len_char_per_row)
                n_rows_to_sample = min(n_rows_to_sample, len(df_to_summarize))
                sampled_df_rows = df_to_summarize.sample(n=n_rows_to_sample)
                input_text = "\n".join(sampled_df_rows[target_col_to_summarize].astype(str).tolist())
        else:
            return "No data provided or data in incorrect format for summarization."
    except KeyError:
        return f"Column '{target_col_to_summarize}' not found in the data provided for summarization."
    except Exception as e:
        return f"Error preparing text for summarization: {e}"

    user_message_summarizer = f'Summarize the following user inputs: {input_text}'
    messages_summarizer = [
        {'role': 'system', 'content': system_prompt_summarizer},
        {'role': 'user', 'content': user_message_summarizer}
    ]

    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=messages_summarizer,
            temperature=0.2,
            max_tokens=512
        )
        summary_content = response.choices[0].message.content
        summary_content_cleaned = re.sub(r'<think>.*?</think>\s*', '', summary_content, flags=re.DOTALL)
        return summary_content_cleaned.strip()
    except Exception as e:
        return f"Error during summarization API call: {e}"

def convert_to_query_func(function_args, df_data, quant_ops):
    operation = function_args.get('operation')
    target_col = function_args.get('target_col')
    group_by_cols_arg = function_args.get('group by')
    filter_conditions = function_args.get('filter')
    top_n = function_args.get('top_n', 3)
    try:
        top_n = int(top_n)
    except (ValueError, TypeError):
        top_n = 3

    if not operation:
        return "Missing operation."
    if operation != 'example':
        if not target_col:
            return "Missing target_col for this operation."
        if target_col not in df_data.columns:
            return f"Target column '{target_col}' not found."

    processed_group_by_cols = []
    if group_by_cols_arg:
        if isinstance(group_by_cols_arg, dict):
            processed_group_by_cols = [
                group_by_cols_arg[str(i)]
                for i in sorted(int(k) for k in group_by_cols_arg.keys())
                if str(i) in group_by_cols_arg
            ]
        elif isinstance(group_by_cols_arg, list):
            processed_group_by_cols = [col for col in group_by_cols_arg if isinstance(col, str)]
        elif isinstance(group_by_cols_arg, str):
            processed_group_by_cols = [group_by_cols_arg]
        else:
            return "Invalid format for 'group by' columns."
        for gb_col in processed_group_by_cols:
            if gb_col not in df_data.columns:
                return f"Group by column '{gb_col}' not found."

    internal_df = df_data.copy()
    if filter_conditions:
        for col, values in filter_conditions.items():
            if col not in internal_df.columns:
                return f"Filter column '{col}' not found."
            internal_df = internal_df[
                internal_df[col].astype(str).str.lower().isin([str(v).lower() for v in values])
            ]

    data_source_for_col_ops = None
    is_grouped_for_col_ops = False

    if operation != 'example':
        if processed_group_by_cols:
            if target_col and target_col in internal_df.columns:
                grouped_df_obj_for_col_ops = internal_df.groupby(processed_group_by_cols)
                data_source_for_col_ops = grouped_df_obj_for_col_ops[target_col]
                is_grouped_for_col_ops = True
            elif target_col:
                return f"Target column '{target_col}' not found for grouped operation."
        else:
            if target_col and target_col in internal_df.columns:
                data_source_for_col_ops = internal_df[target_col]
            elif target_col:
                return f"Target column '{target_col}' not found."
        is_numeric_col_for_operation = (
            pd.api.types.is_numeric_dtype(internal_df[target_col])
            if target_col and target_col in internal_df.columns else False
        )

    try:
        result = None
        if operation == 'example':
            if processed_group_by_cols:
                if internal_df.empty:
                    result = "No data available after filtering to provide grouped examples."
                else:
                    example_grouped_df_obj = internal_df.groupby(processed_group_by_cols)
                    if not example_grouped_df_obj.groups:
                        result = "No groups formed from the data to provide examples."
                    else:
                        result = example_grouped_df_obj.first()
                        if result.empty:
                            result = {"message": "Groups are empty; no example rows to provide."}
                        else:
                            result = result.to_dict()
            else:
                if not internal_df.empty:
                    sample_df = internal_df.sample(n=min(1, len(internal_df)))
                    result = sample_df.to_dict('records')[0] if not sample_df.empty else {"message": "No data to sample"}
                else:
                    result = "No data available to sample an example row from."
        elif data_source_for_col_ops is None:
            return f"Data source for column '{target_col}' could not be prepared for operation '{operation}'."
        elif operation == 'count':
            result = data_source_for_col_ops.count()
        elif operation == 'sum':
            result = data_source_for_col_ops.sum() if is_numeric_col_for_operation else "Sum requires a numeric column."
        elif operation == 'avg':
            result = data_source_for_col_ops.mean() if is_numeric_col_for_operation else "Average requires a numeric column."
        elif operation == 'max':
            result = data_source_for_col_ops.max()
        elif operation == 'min':
            result = data_source_for_col_ops.min()
        elif operation == 'median':
            result = data_source_for_col_ops.median() if is_numeric_col_for_operation else "Median requires a numeric column."
        elif operation == 'unique_values':
            if is_grouped_for_col_ops:
                result = data_source_for_col_ops.apply(lambda x: x.unique().tolist())
            else:
                result = data_source_for_col_ops.unique().tolist()
        elif operation == 'frequency':
            result = data_source_for_col_ops.value_counts()
        elif operation == 'unique_counts':
            result = data_source_for_col_ops.nunique()
        elif operation == 'most_frequent':
            if is_grouped_for_col_ops:
                result = data_source_for_col_ops.value_counts().groupby(
                    level=list(range(len(processed_group_by_cols))),
                    group_keys=False
                ).head(top_n)
            else:
                result = data_source_for_col_ops.value_counts().head(top_n)
        elif operation == 'least_frequent':
            if is_grouped_for_col_ops:
                result = data_source_for_col_ops.value_counts(ascending=True).groupby(
                    level=list(range(len(processed_group_by_cols))),
                    group_keys=False
                ).head(top_n)
            else:
                result = data_source_for_col_ops.value_counts(ascending=True).head(top_n)
        elif operation == 'summerize':
            df_for_summary = internal_df
            if not target_col:
                return "Target column required for summarization."
            return summerize_with_model_func(df_for_summary, target_col)
        else:
            return f"Operation '{operation}' not recognized or target column data source is missing."

        if isinstance(result, pd.Series):
            if result.empty:
                return {}
            if isinstance(result.index, pd.MultiIndex):
                return {str(key): value for key, value in result.items()}
            return result.to_dict()
        elif isinstance(result, pd.DataFrame):
            if result.empty:
                return []
            if isinstance(result.index, pd.MultiIndex):
                return {str(key): row.to_dict() for key, row in result.iterrows()}
            return result.to_dict('records')
        elif isinstance(result, np.generic):
            return result.item()
        elif isinstance(result, list):
            return result
        return result

    except Exception as e:
        return f"Error during operation '{operation}' (target: '{target_col}'): {str(e)}"

def create_plot_html(result, tool_call):
    try:
        op = json.loads(tool_call.arguments).get("operation", "")
        target_col = json.loads(tool_call.arguments).get("target_col", None)
        group_by = json.loads(tool_call.arguments).get("group by", None)
        
        fig, ax = plt.subplots(figsize=(8, 6))
        plot_created = False
        
        # If result is a DataFrame and has numeric columns, plot
        if isinstance(result, pd.DataFrame):
            df = result
            # If only one or two columns, try bar plot
            if df.shape[1] == 2:
                x, y = df.columns[0], df.columns[1]
                df.plot(kind="bar", x=x, y=y, ax=ax, legend=False)
                ax.set_xlabel(x)
                ax.set_ylabel(y)
                plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
                plot_created = True
            elif df.shape[1] == 1 and target_col:
                # Single column, plot value counts
                value_counts = df[target_col].value_counts()
                value_counts.plot(kind="bar", ax=ax)
                ax.set_xlabel(target_col)
                ax.set_ylabel("Count")
                plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
                plot_created = True
        elif isinstance(result, dict):
            # If dict of numeric values, plot bar
            vals = list(result.values())
            if all(isinstance(v, (int, float, np.integer, np.floating)) for v in vals):
                ax.bar(result.keys(), result.values())
                ax.set_xlabel("Category")
                ax.set_ylabel("Value")
                plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
                plot_created = True
        elif isinstance(result, list):
            # If list of dicts with same keys, try to plot
            if len(result) > 0 and isinstance(result[0], dict):
                keys = result[0].keys()
                # Only plot if all dicts have same keys and at least one numeric
                if all(set(d.keys()) == set(keys) for d in result):
                    # Try to plot first numeric key
                    for k in keys:
                        if all(isinstance(d[k], (int, float, np.integer, np.floating)) for d in result):
                            ax.bar(range(len(result)), [d[k] for d in result])
                            ax.set_xlabel("Index")
                            ax.set_ylabel(k)
                            plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
                            plot_created = True
                            break
        
        if plot_created:
            plt.tight_layout()
            # Convert plot to base64 string
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', bbox_inches='tight', dpi=100)
            buffer.seek(0)
            plot_data = buffer.getvalue()
            buffer.close()
            plt.close(fig)
            
            plot_base64 = base64.b64encode(plot_data).decode()
            return f'<img src="data:image/png;base64,{plot_base64}" style="max-width: 100%; height: auto;">'
        else:
            plt.close(fig)
            return None
    except Exception as e:
        plt.close(fig)
        return None

def execute_function_call_wrapper(function_call_obj, df_data, quant_ops):
    function_name = function_call_obj.name
    try:
        function_args_dict = json.loads(function_call_obj.arguments)
    except json.JSONDecodeError as e:
        return f"Error decoding JSON arguments: {e}"

    if function_name == 'convert_to_query':
        return convert_to_query_func(function_args_dict, df_data, quant_ops)
    return f"Function '{function_name}' not found."

def analyze_query(user_query, conversation_history):
    if not user_query:
        st.warning("Please enter a question.")
        return None
    else:
        with st.spinner("🤖 Agent is thinking..."):
            messages_for_llm = [
                {'role': 'system', 'content': f'{system_prompt_text}\n\nConversation History: {conversation_history}'},
                {'role': 'user', 'content': user_query}
            ]
            try:
                response_from_llm = client.chat.completions.create(
                    model=MODEL_NAME,
                    messages=messages_for_llm,
                    temperature=0,
                    tools=tools,
                    tool_choice='auto',
                )
                assistant_message_obj = response_from_llm.choices[0].message

                if 'OOS' in (assistant_message_obj.content or ''):
                    return {'type': 'error', 'content': 'out of scope'}
                elif assistant_message_obj.tool_calls:
                    tool_call = assistant_message_obj.tool_calls[0].function
                    function_result_val = execute_function_call_wrapper(
                        tool_call, df_global, quantitative_operations_list
                    )
                    return {
                        'type': 'tool',
                        'tool_call': {
                            "function_name": tool_call.name,
                            "arguments": json.loads(tool_call.arguments)
                        },
                        'result': function_result_val,
                        'content': assistant_message_obj.content
                    }
                else:
                    return {
                        'type': 'text',
                        'content': assistant_message_obj.content
                        if assistant_message_obj.content
                        else "_No direct textual response provided._"
                    }
            except Exception as e:
                return {'type': 'error', 'content': f"An error occurred during LLM call or function execution: {e}"}

if not client or df_global is None:
    st.error("Application cannot start. OpenAI client or Dataset is not initialized properly. Check API key and data loading.")
else:
    st.sidebar.header("Dataset Information")
    st.sidebar.write(f"**Rows:** {df_global.shape[0]}, **Columns:** {df_global.shape[1]}")
    st.sidebar.write("**Available Columns:**", df_columns)
    if st.sidebar.checkbox("Show sample data", True):
        st.sidebar.dataframe(df_global.head())

    if 'conversation_history' not in st.session_state:
        st.session_state.conversation_history = []

    # Main chat container - everything goes inside here
    chat_html = '<div class="stChatContainer" id="chatContainer" style="height: 60vh; overflow-y: auto; border: 1px solid #eee; border-radius: 8px; padding: 1em; background: #fff; margin-bottom: 1em;">'
    
    for i, msg in enumerate(st.session_state.conversation_history):
        if msg['role'] == 'user':
            chat_html += f'<div class="user-message stChatMessage" style="background: #e6f0fa; border-radius: 8px; padding: 0.7em 1em; margin-bottom: 0.5em; width: fit-content; max-width: 80%; margin-left: auto;"><b>You:</b> {msg["content"]}</div>'
        elif msg['role'] == 'assistant':
            # msg['content'] can be dict (tool) or str (text)
            if isinstance(msg['content'], dict):
                content = msg['content']
                if content.get('type') == 'error':
                    chat_html += f'<div class="assistant-message stChatMessage" style="background: #f6f6f6; border-radius: 8px; padding: 0.7em 1em; margin-bottom: 0.5em; width: fit-content; max-width: 80%;"><b>Agent:</b> <span style="color:red;">{content["content"]}</span></div>'
                elif content.get('type') == 'tool':
                    chat_html += f'<div class="assistant-message stChatMessage" style="background: #f6f6f6; border-radius: 8px; padding: 0.7em 1em; margin-bottom: 0.5em; width: fit-content; max-width: 80%;"><b>Agent (Tool Call):</b><br>'
                    chat_html += f'<pre style="background: #f0f0f0; padding: 0.5em; border-radius: 4px; font-size: 0.9em; overflow-x: auto;">{json.dumps(content["tool_call"], indent=2)}</pre>'
                    chat_html += '<b>Result:</b><br>'
                    
                    if isinstance(content['result'], (dict, list)):
                        chat_html += f'<pre style="background: #f0f0f0; padding: 0.5em; border-radius: 4px; font-size: 0.9em; overflow-x: auto;">{json.dumps(content["result"], indent=2)}</pre>'
                        # Try to create plot
                        try:
                            class DummyToolCall:
                                def __init__(self, name, arguments):
                                    self.name = name
                                    self.arguments = json.dumps(arguments)
                            dummy_tool_call = DummyToolCall(
                                content['tool_call']['function_name'],
                                content['tool_call']['arguments']
                            )
                            plot_html = create_plot_html(content['result'], dummy_tool_call)
                            if plot_html:
                                chat_html += f'<br><b>Visualization:</b><br>{plot_html}'
                        except Exception:
                            pass
                    elif isinstance(content['result'], pd.DataFrame):
                        # Convert DataFrame to HTML table
                        df_html = content['result'].to_html(classes='dataframe', table_id=None, escape=False)
                        chat_html += f'<div style="max-height: 300px; overflow-y: auto;">{df_html}</div>'
                        # Try to create plot
                        try:
                            class DummyToolCall:
                                def __init__(self, name, arguments):
                                    self.name = name
                                    self.arguments = json.dumps(arguments)
                            dummy_tool_call = DummyToolCall(
                                content['tool_call']['function_name'],
                                content['tool_call']['arguments']
                            )
                            plot_html = create_plot_html(content['result'], dummy_tool_call)
                            if plot_html:
                                chat_html += f'<br><b>Visualization:</b><br>{plot_html}'
                        except Exception:
                            pass
                    else:
                        chat_html += f'<div style="background: #f0f0f0; padding: 0.5em; border-radius: 4px;">{str(content["result"])}</div>'
                    chat_html += '</div>'
                elif content.get('type') == 'text':
                    chat_html += f'<div class="assistant-message stChatMessage" style="background: #f6f6f6; border-radius: 8px; padding: 0.7em 1em; margin-bottom: 0.5em; width: fit-content; max-width: 80%;"><b>Agent:</b> {content["content"]}</div>'
            else:
                chat_html += f'<div class="assistant-message stChatMessage" style="background: #f6f6f6; border-radius: 8px; padding: 0.7em 1em; margin-bottom: 0.5em; width: fit-content; max-width: 80%;"><b>Agent:</b> {msg["content"]}</div>'
    
    chat_html += '</div>'
    st.markdown(chat_html, unsafe_allow_html=True)

    # Chat input at the bottom
    with st.form(key="chat_input_form", clear_on_submit=True):
        user_query = st.text_input(
            "Ask your question about the customer support data:",
            placeholder="Type your question and press Enter...",
            key="chat_input"
        )
        submitted = st.form_submit_button("Send", use_container_width=True)
        if submitted and user_query.strip():
            st.session_state.conversation_history.append({"role": "user", "content": user_query})
            response = analyze_query(user_query, st.session_state.conversation_history)
            st.session_state.conversation_history.append({"role": "assistant", "content": response})
            st.rerun()