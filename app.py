import streamlit as st
import pandas as pd
import os
import smtplib
from email.message import EmailMessage
from topsis_neeraj_102317014 import topsis  

st.set_page_config(page_title="TOPSIS Analysis Studio", page_icon="📊")
st.title("📊 TOPSIS Analysis Studio")

with st.form("main_form"):
    uploaded_file = st.file_uploader("Upload your Excel data file", type="xlsx")
    weights = st.text_input("Weights", placeholder="e.g., 1,1,1,2")
    impacts = st.text_input("Impacts", placeholder="e.g., +,+,-,+")
    email_id = st.text_input("Receiver Email ID", placeholder="example@gmail.com")
    submit = st.form_submit_button("PROCESS & EMAIL RESULTS")

if submit:
    if not uploaded_file or not weights or not impacts or not email_id:
        st.warning("Please fill in all the fields!")
    else:
        try:
            with st.spinner("Calculating and Sending..."):
                input_file = "data.xlsx"
                output_file = "result.xlsx"
                
                with open(input_file, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                import sys
                sys.argv = ["topsis", input_file, weights, impacts, output_file]
                topsis.main()

                if os.path.exists(output_file):
                    msg = EmailMessage()
                    msg['Subject'] = 'TOPSIS Result - Neeraj 102317014'
                    msg['From'] = 'imneerajsir@gmail.com'
                    msg['To'] = email_id
                    msg.set_content("Please find the attached TOPSIS analysis results.")

                    with open(output_file, 'rb') as f:
                        msg.add_attachment(
                            f.read(), 
                            maintype='application', 
                            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                            filename='result.xlsx'
                        )

                    email_pass = st.secrets["EMAIL_PASS"]
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login('imneerajsir@gmail.com', email_pass)
                        smtp.send_message(msg)

                    st.success(f"Done! Result sent to {email_id}")
                    st.balloons()
                    
                    os.remove(input_file)
                    os.remove(output_file)
                else:
                    st.error("The result file was not generated. Please check your data format (Weights/Impacts count).")

        except Exception as e:
            st.error(f"Error occurred: {e}")
