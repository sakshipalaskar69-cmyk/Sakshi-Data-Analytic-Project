import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import openpyxl



#File Detail Path
INPUT_FILE = "Attendance1.xlsx"
OUTPUT_DIR = "Attendance_output"
FIG_DIR = os.path.join(OUTPUT_DIR, "figures")
os.makedirs(FIG_DIR,exist_ok=True)

#load file
df=pd.read_excel(INPUT_FILE)
print(df.head())

# A Data Cleaning Operations
# Check Duplicates and drop
print("After Removing Duplicates Rows")
df_cleaned = df.drop_duplicates()

# Rename or format column name
print("Rename or format column name")
df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
print(df.columns)

#Missing value
print("Missing values per column:")
df = df.fillna({
    "NAME": "NULL",               
    "STATUS": "NULL",             
    "CLASS": "Not Assigned",         
    "ATTENDANCE_SCORE": 0,           
})


#Data B Analysis
#Average
average_attendance = df['ATTENDANCE_SCORE'].mean()
print(f"Average Attendance: {average_attendance:.2%}")

#Min/Max Value
max_attendance = df['ATTENDANCE_SCORE'].max()
min_attendance = df['ATTENDANCE_SCORE'].min()

print(f"Maximum Attendance Value: {max_attendance}")
print(f"Minimum Attendance Value: {min_attendance}")

#Group By
print("Group by")
g=df.groupby('NAME')['ID'].mean()
print(g)

#Filter Data
high_scores = df[df['ATTENDANCE_SCORE'] > 50]
print(high_scores)

#Save Modified Data
clean_file = os.path.join(OUTPUT_DIR, "cleaned_data.xlsx")
df.to_excel(clean_file, index=False)
print(f"\n Cleaned data saved to {clean_file}")

df_reset = df.reset_index(drop=True)
df_reset.to_excel(clean_file, index=False)


#Create Charts 
#Bar 
df['CLASS'].value_counts().plot(kind='bar', color='blue')
plt.title("Attendance Count")
plt.xlabel("CLASS")
plt.ylabel("STATUS")
plt.tight_layout()
plt.show()

# Pie
plt.title('Pie Chart: Class Distribution')
df['CLASS'].value_counts().plot(
    kind='pie',
    autopct='%1.1f%%',
    startangle=90,
    ylabel=''  
)
plt.show()

#Line
df.plot(marker='o', figsize=(10, 6))
plt.title('Student Status Over Class')
plt.xlabel('CLASS')
plt.ylabel('STATUS')
plt.grid(True)
plt.legend(title='Students')
plt.tight_layout()
plt.show()










    











































