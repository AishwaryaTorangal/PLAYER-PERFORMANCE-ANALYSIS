import pandas as pd
import streamlit as st
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

def save_changes(df, file_path, sheet_name): 
    try:
        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        st.success("Changes saved successfully!")
    except PermissionError:
        st.error("Error saving the file. Please check file permissions.")
    except Exception as e:
        st.error(f"An error occurred while saving: {e}")

def create_data_in_dataframe(df, sheet_name):
    file_path = 'final.xlsx'
    new_data = {}
    st.write("Enter new data:")
    for column in df.columns:
        new_value = st.text_input(f"{column}: ", key=f"input_{column}")
        new_data[column] = new_value
    if st.button("Add new data"):
        df = df._append(new_data, ignore_index=True)
        save_changes(df, file_path, sheet_name) 
        st.write("Updated Data:")
        display_data(df)

def update_data_in_dataframe(df, sheet_name):
    file_path = 'final.xlsx'
    st.write("Update existing data:")
    st.write("Columns available for update:")
    st.write(df.columns)
    
    columns_to_update = st.text_input("Enter column(s) to update (comma-separated): ", key="update_columns").split(',')
    columns_to_update = [col.strip() for col in columns_to_update]  
    
    for col in columns_to_update:
        if col not in df.columns:
            st.warning(f"Column '{col}' not found. Skipping...")
            continue
        
        row_index = int(st.number_input(f"Enter row index to update for column '{col}': ", key=f"update_row_{col}"))
        if row_index < 0 or row_index >= len(df):
            st.warning("Invalid row index. Skipping update for this column.")
            continue
        
        new_value = st.text_input(f"Enter new value for column '{col}' in row {row_index}: ", key=f"update_value_{col}_{row_index}")
        df.at[row_index, col] = new_value
    
    save_changes(df, file_path, sheet_name)
    st.write("Updated Data:")
    display_data(df)

def delete_data_in_dataframe(df, sheet_name):
    file_path = 'final.xlsx'
    player_name = st.text_input("Enter player name to delete: ", key="delete_player_name").strip()
    player_name_lower = player_name.lower()  
    matching_rows = df[df['Player Name'].str.lower() == player_name_lower]

    if matching_rows.empty:
        st.warning(f"No player found with the name '{player_name}'. No data deleted.")
    else:
        st.write("Matching rows to delete:")
        st.write(matching_rows)

        confirmation = st.radio("Are you sure you want to delete the above rows?", ('Yes', 'No'), index=1, key="delete_confirmation")
        if confirmation == 'Yes':
            df = df.drop(matching_rows.index)
            save_changes(df, file_path, sheet_name)
        else:
            st.info("Deletion aborted.")
    st.write("Updated Data:")
    display_data(df)
def evaluate_performance(total_matches_played, total_matches_won):
    if total_matches_played == 0:
        return "No matches played yet. Performance cannot be evaluated."

    win_percentage = (total_matches_won / total_matches_played) * 100
    
    if win_percentage >= 80:
        feedback = "Excellent performance! Keep up the good work!"
    elif win_percentage >= 60:
        feedback = "Good performance! There's always room for improvement."
    elif win_percentage >= 40:
        feedback = "Fair performance. Keep practicing to improve."
    else:
        feedback = "Needs improvement. Identify areas of weakness and work on them."
    
    return feedback


def display_data(df):
    st.write(df)


def plot_data(df):
    st.write("Select plot type:")
    plot_type = st.selectbox(" ", ['Countplot', 'Line Plot', 'Histogram'])
    
    if plot_type == 'Countplot':
        st.write("Select column for countplot:")
        column_to_plot = st.selectbox(" ", df.columns)
        
        fig, ax = plt.subplots(figsize=(10, 6))
        if df[column_to_plot].dtype == 'object':
            sns.countplot(data=df, x=column_to_plot, ax=ax)
            ax.tick_params(axis='x', rotation=45)
            st.pyplot(fig)
        else:
            st.error("Selected column is not categorical. Please choose a categorical column for plotting.")
            return
        
        
    elif plot_type == 'Line Plot':
        st.write("Select columns for line plot:")
        x_column = st.selectbox("X-axis", df.columns)
        y_column = st.selectbox("Y-axis", df.columns)
        
        if np.issubdtype(df[x_column].dtype, np.number) and np.issubdtype(df[y_column].dtype, np.number):
            plot_line(df, x_column, y_column)
        else:
            st.error("Selected columns are not numeric. Please choose numeric columns for plotting.")
    elif plot_type == 'Histogram':
        st.write("Select column for histogram:")
        column_to_plot = st.selectbox(" ", df.columns)
    
        fig, ax = plt.subplots(figsize=(10, 6))
        if np.issubdtype(df[column_to_plot].dtype, np.number):
            sns.histplot(df[column_to_plot], ax=ax)
            ax.set_xlabel(column_to_plot)
            ax.set_ylabel('Frequency')
            ax.set_title(f'Histogram of {column_to_plot}')
            st.pyplot(fig)
        else:
            st.error("Selected column is not numeric. Please choose a numeric column for plotting.")
            return
    

def plot_line(df, x_column, y_column):
    fig, ax = plt.subplots(figsize=(8, 6))
    ax.plot(df[x_column], df[y_column], marker='o', linestyle='-')
    ax.set_title(f'Line Plot of {y_column} vs {x_column}')
    ax.set_xlabel(x_column)
    ax.set_ylabel(y_column)
    ax.grid(True)
    st.pyplot(fig)

def main():
    file_path = 'final.xlsx'
    try:
        df_2024 = pd.read_excel(file_path, sheet_name='2024')
        df_2023 = pd.read_excel(file_path, sheet_name='2023')
        df_2022 = pd.read_excel(file_path, sheet_name='2022')
    except FileNotFoundError:
        st.error("File not found. Please provide the correct file path.")
        return 
    except pd.errors.ParserError:
        st.error("Error parsing the Excel file. Please check the file format.")
        return 

    st.sidebar.title("Choose action")
    selected_option = st.sidebar.radio(" ", ['2024', '2023', '2022', 'Create', 'Update', 'Delete', 'Display', 'Plot'])

    if selected_option == '2024':
        display_data(df_2024)
        
        total_matches_played = st.number_input("Enter total matches played: ")
        total_matches_won = st.number_input("Enter total matches won: ")
    
        performance_feedback = evaluate_performance(total_matches_played, total_matches_won)
        st.write("Performance Feedback:", performance_feedback)

    elif selected_option == '2023':
        display_data(df_2023)
        
        total_matches_played = st.number_input("Enter total matches played: ")
        total_matches_won = st.number_input("Enter total matches won: ")
    
        performance_feedback = evaluate_performance(total_matches_played, total_matches_won)
        st.write("Performance Feedback:", performance_feedback)

    elif selected_option == '2022':
        display_data(df_2022)
        
        total_matches_played = st.number_input("Enter total matches played: ")
        total_matches_won = st.number_input("Enter total matches won: ")
    
        performance_feedback = evaluate_performance(total_matches_played, total_matches_won)
        st.write("Performance Feedback:", performance_feedback)

    elif selected_option == 'Create':
        selected_sheet_name = st.selectbox("Select sheet to add data:", ['2024', '2023', '2022'])
        if selected_sheet_name == '2024':
            create_data_in_dataframe(df_2024, '2024')
        elif selected_sheet_name == '2023':
            create_data_in_dataframe(df_2023, '2023')
        elif selected_sheet_name == '2022':
            create_data_in_dataframe(df_2022, '2022')
    elif selected_option == 'Update':
        selected_sheet_name = st.selectbox("Select sheet to update data:", ['2024', '2023', '2022'])
        if selected_sheet_name == '2024':
            update_data_in_dataframe(df_2024, '2024')
        elif selected_sheet_name == '2023':
            update_data_in_dataframe(df_2023, '2023')
        elif selected_sheet_name == '2022':
            update_data_in_dataframe(df_2022, '2022')
    elif selected_option == 'Delete':
        selected_sheet_name = st.selectbox("Select sheet to delete data:", ['2024', '2023', '2022'])
        if selected_sheet_name == '2024':
            delete_data_in_dataframe(df_2024, '2024')
        elif selected_sheet_name == '2023':
            delete_data_in_dataframe(df_2023, '2023')
        elif selected_sheet_name == '2022':
            delete_data_in_dataframe(df_2022, '2022')
    elif selected_option == 'Display':
        selected_sheet_name = st.selectbox("Select sheet to display:", ['2024', '2023', '2022'])
        if selected_sheet_name == '2024':
            display_data(df_2024)
        elif selected_sheet_name == '2023':
            display_data(df_2023)
        elif selected_sheet_name == '2022':
            display_data(df_2022)
    elif selected_option == 'Plot':
        selected_sheet_name = st.selectbox("Select sheet to plot:", ['2024', '2023', '2022'])
        if selected_sheet_name == '2024':
            plot_data(df_2024)
        elif selected_sheet_name == '2023':
            plot_data(df_2023)
        elif selected_sheet_name == '2022':
            plot_data(df_2022)

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.write(f"Welcome! Current Date and Time: {current_time}")

if __name__ == "__main__":
    main()
