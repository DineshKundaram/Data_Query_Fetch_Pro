import streamlit as st
import pandas as pd
import oracledb
import xlsxwriter
from io import BytesIO
import warnings
import time

warnings.filterwarnings("ignore")

# Constants for Excel limitations
MAX_ROWS = 1048576  # Excel's max rows
MAX_COLS = 16384    # Excel's max columns

# Database connection pools  #add user, password, dsn for database
pools = {
    'Database_1': oracledb.create_pool(user='', password='', 
                                              dsn='', 
                                              min=10, max=10, increment=0, threaded=True,
                                              getmode=oracledb.SPOOL_ATTRVAL_FORCEGET),
    'Database_2': oracledb.create_pool(user='', password='', 
                                       dsn='', 
                                       min=10, max=10, increment=0, threaded=True,
                                       getmode=oracledb.SPOOL_ATTRVAL_FORCEGET),
    
}

st.title("SQL Query Executor")

# User input for selecting the country
country = st.selectbox("Select the country:", options=list(pools.keys()))

# User input for SQL query with a larger text area
sql_query = st.text_area("Enter your SQL query here:", height=300)

# User input for file name
file_name = st.text_input("Enter the file name for the Excel output (without extension):")

# Execute button
if st.button("Execute and Download"):
    if sql_query and file_name:
        try:
            con = pools[country].acquire()
            cur = con.cursor(scrollable=True)
            cur.prefetchrows = 1000
            cur.arraysize = 1000

            # Execute query and fetch column names
            cur.execute(sql_query)
            columns = [desc[0] for desc in cur.description]
            
            # Fetch data in batches
            batch_size = 1000
            data = pd.DataFrame(columns=columns)
            while True:
                rows = cur.fetchmany(batch_size)
                if not rows:
                    break
                batch_data = pd.DataFrame(rows, columns=columns)
                data = pd.concat([data, batch_data], ignore_index=True)
            
            # Get dimensions
            num_rows, num_cols = data.shape
            
            # Create output buffer
            output = BytesIO()
            
            if num_rows > MAX_ROWS or num_cols > MAX_COLS:
                # Calculate number of chunks needed
                num_row_chunks = -(-num_rows // MAX_ROWS)  # Ceiling division
                num_col_chunks = -(-num_cols // MAX_COLS)  # Ceiling division
                
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for i in range(num_row_chunks):
                        start_row = i * MAX_ROWS
                        end_row = min((i + 1) * MAX_ROWS, num_rows)
                        
                        for j in range(num_col_chunks):
                            start_col = j * MAX_COLS
                            end_col = min((j + 1) * MAX_COLS, num_cols)
                            
                            chunk = data.iloc[start_row:end_row, start_col:end_col]
                            sheet_name = f'Sheet_{i}_{j}'
                            chunk.to_excel(writer, index=False, sheet_name=sheet_name)
                            
                            # Update progress bar
                            progress = ((i * num_col_chunks + j + 1) / (num_row_chunks * num_col_chunks)) * 100
                            st.progress(progress / 100)
                            time.sleep(0.1)  # Simulate processing time
                            
                    # Get the xlsxwriter workbook and worksheet objects
                    workbook = writer.book
                    workbook.nan_inf_to_errors = True
            else:
                # For smaller datasets, use direct xlsxwriter approach
                workbook = xlsxwriter.Workbook(output, {'nan_inf_to_errors': True})
                worksheet = workbook.add_worksheet()
                
                # Write headers
                for col_idx, col_name in enumerate(data.columns):
                    worksheet.write(0, col_idx, col_name)
                
                # Write data
                for r_idx, row in enumerate(data.values, start=1):
                    for c_idx, value in enumerate(row):
                        if pd.isnull(value) or pd.isna(value) or value == float('inf') or value == float('-inf'):
                            worksheet.write(r_idx, c_idx, '')
                        else:
                            worksheet.write(r_idx, c_idx, value)
                
                workbook.close()

            output.seek(0)
            st.success("Query executed successfully!")
            st.download_button(
                label="Download Excel file",
                data=output,
                file_name=f"{file_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            cur.close()
            pools[country].release(con)
        except oracledb.DatabaseError as e:
            st.error(f"Database error: {e}")
        except pd.errors.EmptyDataError:
            st.error("No data returned from the query.")
        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")
    else:
        st.warning("Please enter both a SQL query and a file name.")



