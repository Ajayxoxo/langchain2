import pandas as pd
from langchain.llms import Cohere
from langchain_experimental.agents import create_pandas_dataframe_agent
from langchain.agents.agent_types import AgentType
import streamlit as st

# Initialize Cohere LLM
llm = Cohere(cohere_api_key="DWYlDQorZ6fUuT2Hfzh02vlWdWML0l6jXLuUhFVX", temperature=0)

# Streamlit UI
st.title("Excel Data Analyzer")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)
    
    # Display the dataframe
    st.write("Excel Data:")
    st.dataframe(df)

    # Create Pandas DataFrame Agent
    agent = create_pandas_dataframe_agent(
        llm,
        df,
        verbose=True,
        agent_type=AgentType.ZERO_SHOT_REACT_DESCRIPTION,
        allow_dangerous_code=True
    )

    # User prompt input
    user_prompt = st.text_input("Enter your question or task for the Excel data:")

    if user_prompt:
        with st.spinner("Analyzing..."):
            # Run the agent
            response = agent.run(user_prompt)
            
            # Display the response
            st.write("Analysis Result:")
            st.write(response)

    # Option to download modified Excel file
    if st.button("Download Modified Excel"):
        # Save the potentially modified dataframe
        df.to_excel("modified_excel.xlsx", index=False)
        st.download_button(
            label="Click to Download",
            data=open("modified_excel.xlsx", "rb"),
            file_name="modified_excel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
