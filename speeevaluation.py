import subprocess
import pandas as pd

# Load the Excel file
input_file = "company_websites_with_usability.xlsx"
output_file = "company_websites_with_usability_updated.xlsx"
df = pd.read_excel(input_file)

# Ensure there's a "Page performance" column and "performance Category" column
if "performance Evaluation" not in df.columns:
    df["performance Evaluation"] = None
if "performance Category" not in df.columns:
    df["performance Category"] = None

# Function to run Lighthouse
def run_lighthouse(url):
    try:
        lighthouse_path = r"C:\Users\30694\AppData\Roaming\npm\lighthouse.cmd"
        result = subprocess.check_output([
    lighthouse_path,
    url,
    "--quiet",
    "--output=json",
    "--only-categories=performance",
    "--chrome-flags=--headless"
])





        # Parse the JSON result for performance score
        import json
        report = json.loads(result)
        return report["categories"]["performance"]["score"] * 100  # Convert to percentage
    except Exception as e:
        print(f"Error running Lighthouse for {url}: {e}")
        return None

# Function to categorize page performance based on Lighthouse score
def categorize_performance(score):
    if score >= 90:
        return "Fast"
    elif score >= 50:
        return "Moderate"
    else:
        return "Slow"

# Iterate through the URLs
for index, row in df.iterrows():
    website = row["Website"]
    if pd.notna(website):
        print(f"Analyzing performance for: {website}")
        performance_score = run_lighthouse(website)
        if performance_score is not None:
            df.at[index, "Performance Evaluation"] = performance_score
            # Categorize the performance based on the Lighthouse score
            performance_category = categorize_performance(performance_score)
            df.at[index, "Performance Category"] = performance_category
            print(f"Performance Evaluation for {website}: {performance_score} - {performance_category}")
        else:
            df.at[index, "Performance Evaluation"] = None
            df.at[index, "Performance Category"] = "Error"
    else:
        print(f"No website found for: {row['Company Name']}")
        df.at[index, "Performance Category"] = "No URL"

# Save the updated data
df.to_excel(output_file, index=False)
print(f"Updated data saved to '{output_file}'")