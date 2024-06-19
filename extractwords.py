import docx
from collections import Counter
import pandas as pd

# Load the document MODIFY THE PATH TO YOUR FILE
doc_path = './output.docx'
doc = docx.Document(doc_path)

# Extract text from the document
text = []
for para in doc.paragraphs:
    text.append(para.text)

# Combine all paragraphs into a single string
full_text = ' '.join(text)

# Function to find single word frequency
def find_single_word_frequency(text):
    words = text.split()
    word_counts = Counter(words)
    return word_counts

# Find single word frequency
single_word_counts = find_single_word_frequency(full_text)

# Convert to DataFrame and get the top 50
df_single_word = pd.DataFrame(single_word_counts.items(), columns=['Word', 'Frequency']).sort_values(by='Frequency', ascending=False).head(50)
print("Top 50 Single Words")
print(df_single_word)

# Function to find phrases of a given length
def find_phrases(text, length):
    words = text.split()
    phrases = [' '.join(words[i:i+length]) for i in range(len(words)-length+1)]
    return phrases

# Find phrases of length 2, 3, 4, 5, 6, and 7
phrases_2 = find_phrases(full_text, 2)
phrases_3 = find_phrases(full_text, 3)
phrases_4 = find_phrases(full_text, 4)
phrases_5 = find_phrases(full_text, 5)
phrases_6 = find_phrases(full_text, 6)
phrases_7 = find_phrases(full_text, 7)

# Count the frequency of each phrase
count_2 = Counter(phrases_2)
count_3 = Counter(phrases_3)
count_4 = Counter(phrases_4)
count_5 = Counter(phrases_5)
count_6 = Counter(phrases_6)
count_7 = Counter(phrases_7)

# Convert to DataFrame and get the top 10 for each
df_2 = pd.DataFrame(count_2.items(), columns=['Phrase', 'Frequency']).sort_values(by='Frequency', ascending=False).head(10)
df_3 = pd.DataFrame(count_3.items(), columns=['Phrase', 'Frequency']).sort_values(by='Frequency', ascending=False).head(10)
df_4 = pd.DataFrame(count_4.items(), columns=['Phrase', 'Frequency']).sort_values(by='Frequency', ascending=False).head(10)
df_5 = pd.DataFrame(count_5.items(), columns=['Phrase', 'Frequency']).sort_values(by='Frequency', ascending=False).head(10)
df_6 = pd.DataFrame(count_6.items(), columns=['Phrase', 'Frequency']).sort_values(by='Frequency', ascending=False).head(10)
df_7 = pd.DataFrame(count_7.items(), columns=['Phrase', 'Frequency']).sort_values(by='Frequency', ascending=False).head(10)

print("\nTop 2-word Phrases")
print(df_2)
print("\nTop 3-word Phrases")
print(df_3)
print("\nTop 4-word Phrases")
print(df_4)
print("\nTop 5-word Phrases")
print(df_5)
print("\nTop 6-word Phrases")
print(df_6)
print("\nTop 7-word Phrases")
print(df_7)
