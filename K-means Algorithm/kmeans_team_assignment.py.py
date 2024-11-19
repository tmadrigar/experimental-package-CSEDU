# K-means Clustering and Team Distribution for Learning Styles
# This script consolidates the functionality of K-means clustering, cluster analysis, interpretation,
# and team formation from the provided data. 

import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
import seaborn as sns
from collections import defaultdict
import os

# --- STEP 1: Load Data ---
# Replace 'Learning_Styles_Data.xlsx' with the actual file name containing your data.
data_path = 'Learning_Styles_Data.xlsx'
data = pd.read_excel(data_path)

# --- STEP 2: Calculate Aggregate Scores ---
# Adjust column ranges as needed based on your data structure.
# EC = Abstract Conceptualization, OR = Reflective Observation, CA = Concrete Experience, EA = Active Experimentation
ec_columns = data.iloc[:, 1:11]  # Replace with the actual range for EC-related columns
or_columns = data.iloc[:, 11:22]
ca_columns = data.iloc[:, 22:34]
ea_columns = data.iloc[:, 34:45]

# Aggregate scores for each dimension
data['EC Score'] = ec_columns.sum(axis=1)
data['OR Score'] = or_columns.sum(axis=1)
data['CA Score'] = ca_columns.sum(axis=1)
data['EA Score'] = ea_columns.sum(axis=1)

# --- STEP 3: Normalize Data and Apply K-means ---
# Select columns for clustering
scores_data = data[['EC Score', 'OR Score', 'CA Score', 'EA Score']]

# Normalize scores
scaler = StandardScaler()
scaled_scores = scaler.fit_transform(scores_data)

# Apply K-means with 4 clusters
kmeans = KMeans(n_clusters=4, random_state=42)
data['Cluster'] = kmeans.fit_predict(scaled_scores)

# --- STEP 4: Analyze Clusters ---
# Visualize cluster centers
centers = pd.DataFrame(kmeans.cluster_centers_, columns=['EC Score', 'OR Score', 'CA Score', 'EA Score'])
print("Cluster Centers:\n", centers)

plt.figure(figsize=(10, 6))
sns.heatmap(centers.T, annot=True, cmap="YlGnBu", fmt=".2f")
plt.title('Cluster Centers Heatmap')
plt.show()

# --- STEP 5: Interpret Learning Styles ---
# Determine predominant and secondary learning styles for each individual
def interpret_learning_styles(row, styles):
    sorted_styles = row[styles].sort_values(ascending=False)
    return pd.Series([sorted_styles.index[0], sorted_styles.index[1]], index=['Predominant Style', 'Secondary Style'])

learning_styles = ['EC Score', 'OR Score', 'CA Score', 'EA Score']
data[['Predominant Style', 'Secondary Style']] = data.apply(
    lambda row: interpret_learning_styles(row, learning_styles), axis=1
)

# Save interpreted data to an output file
output_interpreted_path = 'Interpreted_Learning_Styles.xlsx'
data.to_excel(output_interpreted_path, index=False)

# --- STEP 6: Distribute Students into Teams ---
# Assume there are 4 teams
num_teams = 4
teams = defaultdict(list)

# Distribute students among teams based on clusters and learning styles
def distribute_students(data, num_teams):
    for cluster in data['Cluster'].unique():
        cluster_data = data[data['Cluster'] == cluster]
        for _, row in cluster_data.iterrows():
            teams_counts = {team: len(teams[team]) for team in range(1, num_teams + 1)}
            teams_styles = {team: [member['Predominant Style'] for member in teams[team]] for team in range(1, num_teams + 1)}
            possible_teams = [team for team in range(1, num_teams + 1) if row['Predominant Style'] not in teams_styles[team]]

            if possible_teams:
                selected_team = min(possible_teams, key=lambda t: teams_counts[t])
            else:
                selected_team = min(teams_counts, key=teams_counts.get)

            teams[selected_team].append(row.to_dict())

distribute_students(data, num_teams)

# Save team assignments to an output file
teams_data = []
for team, members in teams.items():
    for member in members:
        teams_data.append({
            'Team': team,
            'Student': f"Student_{member.name}",
            'Cluster': member['Cluster'],
            'Predominant Style': member['Predominant Style'],
            'Secondary Style': member['Secondary Style']
        })

teams_df = pd.DataFrame(teams_data)
output_teams_path = 'Formed_Teams.xlsx'
teams_df.to_excel(output_teams_path, index=False)

print(f"Team assignments saved to: {output_teams_path}")
