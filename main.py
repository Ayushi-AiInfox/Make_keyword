import pandas as pd
import openai
import os
from dotenv import load_dotenv

load_dotenv()  

# Replace the hardcoded key with an environment variable
openai.api_key = os.getenv("OPENAI_API_KEY")
df = pd.read_excel('demo.xlsx')
print("df",df.columns)

# print("df['Description']",df['Description'])
# print("df['Specialties']",df['Specialties'])

df['Description_Specialties'] = df['Description'].fillna('') + ' ' + df['Specialties'].fillna('')
df['Description_Specialties'] = df['Description_Specialties'].str.strip()
# print("df['Description_Specialties']",df['Description_Specialties'][0])


def extract_focused_words(text):
    if isinstance(text, str): 
        # prompt = f"Extract key words from the following description: {text}"
        # prompt = f"Extract meaningful key phrases or short descriptions from the following text: {text}. Provide the output in a list of phrases that summarize the main concepts."
        # prompt = f"Extract concise and meaningful key phrases that summarize the following content. Avoid incomplete phrases and repeating phrases or using unnecessary separators: {text}"
        prompt = f"""You are an help ful agent that is expert in sumarizing the companies,
                     you will have the discription of companies and their specialities finde the Main focus of the companies.
                     return the result in 3-6 keywords in a meaningfull phrase.
                     Given discription and specialities : {text}"""
        
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini", 
            messages=[{"role": "user", "content": prompt}],
            max_tokens=15,  
            n=1,
            stop=None,
            temperature=0.1,  
        )
        
        focused_words = response['choices'][0]['message']['content'].strip()
        return focused_words
    return ""



# df['focused_words_description'] = df['Description'].apply(extract_focused_words)
# df['focused_words_specialties'] = df['Specialties'].apply(extract_focused_words)

# #handle missing value
# df['focused_words_specialties'] = df['focused_words_specialties'].fillna("")



# df['all_focused_words'] = df['focused_words_description'] + " " + df['focused_words_specialties']


# df['focused_words_description'] = df['focused_words_description'].str.replace('\n', ' ')
# df['focused_words_specialties'] = df['focused_words_specialties'].str.replace('\n', ' ')



# df['all_focused_words'] = df['focused_words_description'] + " " + df['focused_words_specialties']
# df['all_focused_words'] = df['all_focused_words'].str.replace(' - ', ' ', regex=False)
# df['all_focused_words'] = df['all_focused_words'].str.replace(r'\b\w\b', '', regex=True).str.strip()
# df['all_focused_words'] = df['all_focused_words'].str.replace(r'\s+', ' ', regex=True)

# print("*FOCUSED*", df[['focused_words_description', 'focused_words_specialties', 'all_focused_words']])



df['focused_words_Description_Specialties'] = df['Description_Specialties'].apply(extract_focused_words)

df['focused_words_Description_Specialties'] = df['focused_words_Description_Specialties'].str.replace(' - ', ' ', regex=False)
df['focused_words_Description_Specialties'] = df['focused_words_Description_Specialties'].str.replace(r'\b\w\b', '', regex=True).str.strip()
df['focused_words_Description_Specialties'] = df['focused_words_Description_Specialties'].str.replace(r'\s+', ' ', regex=True)



output_filename = "focused_words_merged_output05.xlsx"
df[['Description_Specialties', 'focused_words_Description_Specialties']].to_excel(output_filename, index=False)

print(f"Excel file with focused words saved as {output_filename}")


























