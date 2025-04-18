import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = r"C:\Users\mailt\Downloads\Number of Schools by Availability of Infrastructure and Facilities, School Management and School Category_Report type - State-wise_22.xlsx"

xls = pd.ExcelFile(file_path)

# Load the actual data sheet
df = xls.parse('ag-grid')

# Extract the header and data starting from the correct row
df_cleaned = df[2:].copy()
df_cleaned.columns = df_cleaned.iloc[0]
df_cleaned = df_cleaned[1:]
df_cleaned.reset_index(drop=True, inplace=True)

# Rename important columns
df_cleaned = df_cleaned.rename(columns={
    'Location': 'State/UT',
    'Rural/Urban': 'Location',
    'Total No. of Schools': 'Total Schools',
    'Electricity': 'Electricity Available',
    'Internet': 'Internet Available',
    'WASH Facility(Drinking Water, Toilet and Handwash)': 'WASH Facilities',
    'School Type': 'School Type'
})

# Convert important columns to numeric
numeric_columns = ['Total Schools', 'Electricity Available', 'Internet Available', 'WASH Facilities']
df_cleaned[numeric_columns] = df_cleaned[numeric_columns].apply(pd.to_numeric, errors='coerce')

# Set visual style
sns.set(style="whitegrid")

# Objective 1: Total schools by location (Rural vs Urban)
plt.figure(figsize=(6,4))
sns.barplot(data=df_cleaned, x='Location', y='Total Schools', estimator=sum, errorbar=None)
plt.title('Total Number of Schools by Location')
plt.ylabel('Number of Schools')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# Objective 2: Distribution of Electricity Availability
plt.figure(figsize=(6,4))
sns.histplot(df_cleaned['Electricity Available'].dropna(), bins=20, kde=True)
plt.title('Distribution of Electricity Availability')
plt.xlabel('Number of Schools with Electricity')
plt.tight_layout()
plt.show()

# Objective 3: Internet availability by location
plt.figure(figsize=(8,5))
sns.boxplot(data=df_cleaned, x='Location', y='Internet Available')
plt.title('Internet Availability by Location')
plt.ylabel('Number of Schools with Internet')
plt.tight_layout()
plt.show()

# Objective 4: Distribution of Schools with WASH Facilities
plt.figure(figsize=(6,4))
sns.histplot(df_cleaned['WASH Facilities'].dropna(), bins=20, color='green')
plt.title('Schools with WASH Facilities')
plt.xlabel('Number of Schools')
plt.tight_layout()
plt.show()

# Objective 5: Distribution by School Type
plt.figure(figsize=(6,4))
sns.countplot(data=df_cleaned, x='School Type', order=df_cleaned['School Type'].value_counts().index)
plt.title('Distribution of Schools by Type')
plt.ylabel('Number of Entries')
plt.tight_layout()
plt.show()

# Objective 6: Top 5 States/UTs with Highest Number of Schools
top_states = df_cleaned.groupby('State/UT')['Total Schools'].sum().sort_values(ascending=False).head(5)
plt.figure(figsize=(8,5))
sns.barplot(x=top_states.values, y=top_states.index, errorbar=None)
plt.title('Top 5 States/UTs with Highest Number of Schools')
plt.xlabel('Number of Schools')
plt.tight_layout()
plt.show()
