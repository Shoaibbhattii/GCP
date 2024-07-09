import pandas as pd
from namefu import fuzz

def find_matches(addresses, city_names):
    exact_matches = []
    approximate_matches = [[] for _ in range(5)]  # Change 5 to a higher number if you expect more variations

    for address in addresses:
        exact_match_found = False
        for city in city_names:
            if address == city:
                exact_matches.append(city)
                exact_match_found = True
                break
            elif fuzz.ratio(address, city) >= 80:
                approximate_matches[0].append((address, city))
                exact_match_found = True
                break
        if not exact_match_found:
            exact_matches.append(None)
            for matches in approximate_matches:
                matches.append((address, None))
    return exact_matches, approximate_matches

# Read the Excel file
file_path = 'fuzzy.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path)

# Assuming the addresses are in a column named 'Address' and city names are in a column named 'City'
addresses = df['Address'].tolist()
city_names = df['City'].tolist()

# Find matches
exact_matches, approximate_matches = find_matches(addresses, city_names)

# Add exact matches to DataFrame
df['Exact_Match'] = exact_matches

# Add approximate matches to DataFrame
for i, matches in enumerate(approximate_matches):
    column_name = f'Approx_Match_{i+1}'
    df[column_name] = [match[1] for match in matches]

# Save the result back to a new Excel file
output_file_path = 'matched_excel_file.xlsx'  # Replace with your desired output file path
df.to_excel(output_file_path, index=False)

print(f"Processed file saved as {output_file_path}")
