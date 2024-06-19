# Document Word and Phrase Frequency Analyzer
###File: duplicatewords.py

This Python script analyzes the frequency of words and phrases in a Microsoft Word document (.docx). It extracts the text from the document, calculates the frequency of individual words, and identifies the most common phrases of varying lengths. The results are displayed as data frames, showing the top occurrences for each category.

## Features

- Extracts text from a .docx file.
- Calculates the frequency of individual words.
- Identifies and counts the frequency of phrases with lengths ranging from 2 to 7 words.
- Outputs the top 50 most frequent individual words.
- Outputs the top 10 most frequent phrases for each phrase length (2 to 7 words).

## Dependencies

The script requires the following Python packages:
- `python-docx`
- `collections`
- `pandas`

You can install these dependencies using pip:

```bash
pip install python-docx pandas
```

Usage

    - Place your Word document in the same directory as the script and name it output.docx or modify the doc_path variable to point to your document's location.
    - Run the script:
```bash
python word_phrase_frequency_analyzer.py
```

Example output:
Top 50 Single Words
| Word   | Frequency |
|--------|-----------|
| the    | 120       |
| of     | 85        |
| and    | 75        |
...

Top 2-word Phrases
| Phrase          | Frequency |
|-----------------|-----------|
| of the          | 30        |
| in the          | 25        |
...

Top 3-word Phrases
| Phrase             | Frequency |
|--------------------|-----------|
| one of the         | 15        |
| as well as         | 12        |
...

...

