import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from rake_nltk import Rake
from openpyxl import load_workbook
import nltk

nltk.download('stopwords')
nltk.download('punkt')
nltk.download('punkt_tab')


def load_data(file_path):
    df = pd.read_excel(file_path)
    return df


def preprocess_text(text):
    if pd.isnull(text):
        return ""
    return text.lower().replace("\n", " ").replace("\r", " ")


def extract_keywords(text, max_words=4):
    r = Rake()
    r.extract_keywords_from_text(text)
    ranked_phrases = r.get_ranked_phrases()
    
    filtered_phrases = [phrase for phrase in ranked_phrases if len(phrase.split()) <= max_words]
    return ", ".join(filtered_phrases[:3])  


def extract_focused_word(text):
    r = Rake()
    r.extract_keywords_from_text(text)
    ranked_phrases = r.get_ranked_phrases()
    if ranked_phrases:
        focused_word = ranked_phrases[0]  
        return focused_word
    return "" 


def generate_focus_columns(df):
    focus_list = []
    focused_word_list = []
    for index, row in df.iterrows():
      
        combined_text = f"{preprocess_text(row['Description'])} {preprocess_text(row['Specialties'])}"
        focus = extract_keywords(combined_text)
        focused_word = extract_focused_word(combined_text)
        
        focus_list.append(focus)
        focused_word_list.append(focused_word)
    
    
    df['Focus'] = focus_list
    df['Focused Word'] = focused_word_list
    return df

def save_to_excel(df, file_path):
    
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet2')  
        

def main():
    input_file = "demo.xlsx"  

 
    df = load_data(input_file)
    df = generate_focus_columns(df)


    save_to_excel(df, input_file)
    print(f"Results saved to {input_file}")

if __name__ == "__main__":
    main()
