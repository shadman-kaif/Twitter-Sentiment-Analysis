# Import necessary libraries
import pandas as pd
from textblob import TextBlob
import numpy as np

# Read Excel sheet
df = pd.read_excel("C:\\Users\\Shadman\\Desktop\\Twitter\\TWEETS.xlsx")

# Convert all the tweets into a list of strings
col_one_list = df['text'].tolist()

# Make lists
polarity_scores = []
polarity_fb = []
subjectivity = []
subjectivity_fb = []

for x in col_one_list:
    blob = TextBlob(x)
    polarity_scores.append(blob.sentiment.polarity)
    subjectivity.append(blob.sentiment.subjectivity)
    if blob.sentiment.polarity == 0:
        polarity_fb.append('Neutral')
    elif blob.sentiment.polarity > 0:
        polarity_fb.append('Positive')
    else: 
        polarity_fb.append('Negative')
    if blob.sentiment.subjectivity == 0.5:
        subjectivity_fb.append('Neutral')
    elif blob.sentiment.subjectivity < 0.5:
        subjectivity_fb.append('Objective')
    else: 
        subjectivity_fb.append('Subjective')

df.loc[:,'Polarity Score'] = polarity_scores
df.loc[:,'Polarity Feedback'] = polarity_fb
df.loc[:,'Subjectivity'] = subjectivity
df.loc[:,'Subjectivity Feedback'] = subjectivity_fb

print(df)

df.to_excel("C:\\Users\\Shadman\\Desktop\\Twitter\\TextBlobTest.xlsx", index = False)