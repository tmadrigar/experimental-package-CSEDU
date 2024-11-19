# Analysis of Preliminary Study: Learning Styles Classification
# This script processes participant responses to classify their learning styles 
# and generates visualizations and summary results.

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# --- STEP 1: Load Data ---
# Replace 'responses_data.xlsx' and 'classified_participants.xlsx' with the actual file names.
responses_data_path = 'responses_data.xlsx'  # File with participants' responses
classified_participants_path = 'classified_participants.xlsx'  # File with pre-classified participants (optional)

data_responses = pd.read_excel(responses_data_path)

# --- STEP 2: Calculate Scores for Each Learning Style ---
# Define the columns related to each learning style (adjust ranges as necessary)
accommodator_columns = data_responses.columns[1:13]  # Columns for Accommodator
converger_columns = data_responses.columns[13:24]  # Columns for Converger
assimilator_columns = data_responses.columns[24:35]  # Columns for Assimilator
diverger_columns = data_responses.columns[35:]  # Columns for Diverger

# Compute total scores for each style
data_responses['Accommodator'] = data_responses[accommodator_columns].sum(axis=1)
data_responses['Converger'] = data_responses[converger_columns].sum(axis=1)
data_responses['Assimilator'] = data_responses[assimilator_columns].sum(axis=1)
data_responses['Diverger'] = data_responses[diverger_columns].sum(axis=1)

# --- STEP 3: Determine Predominant and Secondary Styles ---
# Identify predominant and secondary learning styles based on scores
data_responses['Predominant Style'] = data_responses[
    ['Accommodator', 'Converger', 'Assimilator', 'Diverger']
].idxmax(axis=1)

data_responses['Secondary Style'] = data_responses[
    ['Accommodator', 'Converger', 'Assimilator', 'Diverger']
].apply(lambda row: row.nlargest(2).idxmin(), axis=1)

# --- STEP 4: Visualize Results ---
# Plot the distribution of predominant learning styles
def plot_predominant_styles(data):
    freq_predominant = data['Predominant Style'].value_counts(normalize=True) * 100  # Percentage
    plt.figure(figsize=(8, 6))
    freq_predominant.plot(kind='bar', color='skyblue')
    plt.title('Distribution of Predominant Learning Styles')
    plt.xlabel('Learning Styles')
    plt.ylabel('Percentage of Participants')
    plt.xticks(rotation=45)
    for index, value in enumerate(freq_predominant):
        plt.text(index, value + 1, f'{value:.1f}%', ha='center', va='bottom')
    plt.tight_layout()
    plt.show()

# Plot the distribution of secondary learning styles
def plot_secondary_styles(data):
    freq_secondary = data['Secondary Style'].value_counts(normalize=True) * 100  # Percentage
    plt.figure(figsize=(8, 6))
    freq_secondary.plot(kind='bar', color='lightgreen')
    plt.title('Distribution of Secondary Learning Styles')
    plt.xlabel('Learning Styles')
    plt.ylabel('Percentage of Participants')
    plt.xticks(rotation=45)
    for index, value in enumerate(freq_secondary):
        plt.text(index, value + 1, f'{value:.1f}%', ha='center', va='bottom')
    plt.tight_layout()
    plt.show()

# Plot the correlation matrix of learning styles
def plot_correlation_matrix(data):
    correlation_matrix = data[['Accommodator', 'Converger', 'Assimilator', 'Diverger']].corr()
    plt.figure(figsize=(8, 6))
    sns.heatmap(correlation_matrix, annot=True, cmap='Blues', fmt='.2f', cbar=True)
    plt.title('Correlation Matrix of Learning Styles')
    plt.tight_layout()
    plt.show()

# Call functions to generate visualizations
plot_predominant_styles(data_responses)
plot_secondary_styles(data_responses)
plot_correlation_matrix(data_responses)

# --- STEP 5: Save Results ---
# Save classified results to an output file
output_path = 'classified_learning_styles_results.xlsx'  # Replace with your desired file name
data_responses.to_excel(output_path, index=False)

print(f"Results have been saved to: {output_path}")
